"""Gaussian fitting of interference fringes in optical spectra.

A thin, transparent layer (for example the sample chamber of a diamond anvil
cell, a film or a plate) acts as a low-finesse Fabry-Perot cavity, so its
transmission or reflection spectrum contains interference fringes. This script
finds every fringe and fits it with a Gaussian to obtain its centre.

The steps are:

1. Read one or more spectra (x, intensity) from a text, CSV or Excel file.
   The x axis can be wavelength (nm, um or Angstrom) or wavenumber (cm^-1 or
   um^-1). Header lines are skipped automatically.
2. Split the spectrum into continuous segments. A spectrum merged from
   several detector windows has overlap regions, where points from two windows
   interleave, and gaps; these are detected from the spacing of the x values.
   Overlap regions are discarded and nothing is interpolated across segments.
   A single continuous spectrum is left as one segment.
3. Convert each segment to wavenumber, where the fringes are evenly spaced,
   and divide it by a smooth Savitzky-Golay baseline.
4. Locate each fringe maximum (or minimum) and fit it, over one fringe period,
   with a Gaussian on a linear background. The Gaussian is used only to locate
   the fringe centre, not as a model of the line shape. By default a cosine
   fit over the same range is also made, to reject poor fringes and to check
   the centres.
5. Write a table of fringe centres and an overview plot for each spectrum.

Command line
------------
    python fringe_fitting.py examples/KCl_0.68GPa_full_range.csv --out results
    python fringe_fitting.py data.txt --x-unit cm-1 --y-col all --minima

Run ``python fringe_fitting.py -h`` for all options.

Python
------
    from fringe_fitting import fit_fringes
    table = fit_fringes(x, y, x_unit="nm")
"""
from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit
from scipy.signal import find_peaks, savgol_filter

log = logging.getLogger("fringe_fitting")

X_UNITS = ("nm", "um", "angstrom", "cm-1", "um-1")


@dataclass
class Settings:
    """Analysis settings. The defaults reproduce the analysis of the KCl paper."""

    split_segments: bool = True       # detect overlaps and gaps between detector windows
    min_segment_points: int = 50      # shorter segments are ignored
    edge_trim: int = 20               # points dropped at each end of a segment
    baseline_periods: float = 2.5     # Savitzky-Golay baseline window, in fringe periods
    smoothing_periods: float = 0.25   # smoothing used only to locate the maxima
    min_peak_separation: float = 0.6  # minimum distance between fringes, in periods
    fit_halfwidth: float = 0.5        # half-width of each fit range, in periods
    cosine_check: bool = True         # reject fringes whose cosine fit fails
    period_tolerance: float = 0.3     # accepted fractional deviation of the cosine period
    period_per_segment: bool = False  # estimate the fringe period separately for each segment
    minima: bool = False              # fit fringe minima instead of maxima (e.g. reflection)


# ----------------------------------------------------------------------------
# Units
# ----------------------------------------------------------------------------
def to_wavenumber(x: np.ndarray, unit: str) -> np.ndarray:
    """Convert x values to wavenumber in 1/um."""
    unit = unit.lower().replace("µ", "u")
    if unit == "nm":
        return 1000.0 / x
    if unit == "um":
        return 1.0 / x
    if unit in ("angstrom", "a"):
        return 1.0e4 / x
    if unit in ("cm-1", "1/cm"):
        return x / 1.0e4
    if unit in ("um-1", "1/um"):
        return x
    raise ValueError(f"unknown x unit {unit!r}; use one of {', '.join(X_UNITS)}")


# ----------------------------------------------------------------------------
# Input
# ----------------------------------------------------------------------------
def _numeric_block(df: pd.DataFrame) -> pd.DataFrame:
    """Drop header and footer rows: keep rows whose first two columns are numbers."""
    num = df.apply(pd.to_numeric, errors="coerce")
    good = num.iloc[:, 0].notna() & num.iloc[:, 1].notna()
    if not good.any():
        raise ValueError("no numeric data found")
    return num[good].dropna(axis=1, how="all")


def _read_table(path: Path) -> pd.DataFrame:
    """Read a delimited text file with an unknown delimiter and header."""
    for sep in (None, r"\s+", ",", ";", "\t"):
        try:
            df = pd.read_csv(path, sep=sep, header=None, engine="python", comment="#",
                             encoding="utf-8-sig", skip_blank_lines=True, on_bad_lines="skip")
        except (pd.errors.ParserError, UnicodeDecodeError, ValueError):
            continue
        if df.shape[1] >= 2:
            try:
                return _numeric_block(df)
            except ValueError:
                continue
    raise ValueError(f"could not read two numeric columns from {path}")


def read_spectra(path: Path, sheet: str | None = None, x_col: int = 0,
                 y_cols: str = "1") -> dict[str, tuple[np.ndarray, np.ndarray]]:
    """Read spectra from a .csv/.txt/.dat or .xlsx/.xls file.

    Returns {name: (x, y)}. ``y_cols`` is a comma-separated list of column
    numbers (counting from 0) or ``"all"`` for every column except ``x_col``.
    For a workbook every sheet is read unless ``sheet`` is given.
    """
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        raw = pd.read_excel(path, sheet_name=sheet if sheet else None, header=None)
        if isinstance(raw, pd.DataFrame):
            raw = {sheet: raw}
        blocks = {str(k): _numeric_block(v) for k, v in raw.items()}
    else:
        blocks = {path.stem: _read_table(path)}
    spectra = {}
    for name, df in blocks.items():
        cols = [c for c in range(df.shape[1]) if c != x_col] if y_cols == "all" else [int(c) for c in y_cols.split(",")]
        for c in cols:
            data = df.iloc[:, [x_col, c]].dropna().to_numpy(float)
            label = name if len(cols) == 1 else f"{name}_col{c}"
            spectra[label] = (data[:, 0], data[:, 1])
    return spectra


# ----------------------------------------------------------------------------
# Segments
# ----------------------------------------------------------------------------
def split_segments(x: np.ndarray, y: np.ndarray, s: Settings) -> list[tuple[np.ndarray, np.ndarray]]:
    """Split a sorted spectrum into continuous segments.

    Within a detector window the x step is regular. Where two windows overlap,
    their points interleave and the step becomes irregularly small; where
    there is a gap, the step is large. Overlap regions are dropped from both
    windows because their points cannot always be assigned unambiguously.
    """
    if not s.split_segments or len(x) < 3:
        return [(x, y)]
    step = np.diff(x)
    typical = np.median(step)
    small = np.where(step < 0.75 * typical)[0]
    overlaps = []
    if len(small):
        start = prev = small[0]
        for i in small[1:]:
            if i - prev > 8:
                overlaps.append((start, prev + 1))
                start = i
            prev = i
        overlaps.append((start, prev + 1))
    cuts = sorted([(a, b + 1) for a, b in overlaps] + [(g + 1, g + 1) for g in np.where(step > 3 * typical)[0]])
    segments, pos = [], 0
    for a, b in cuts:
        if a - pos > s.min_segment_points:
            segments.append((x[pos:a], y[pos:a]))
        pos = b
    if len(x) - pos > s.min_segment_points:
        segments.append((x[pos:], y[pos:]))
    return segments


# ----------------------------------------------------------------------------
# Fitting
# ----------------------------------------------------------------------------
def gaussian(x, offset, amplitude, slope, centre, width):
    return offset + slope * (x - centre) + amplitude * np.exp(-((x - centre) ** 2) / (2 * width ** 2))


def cosine(x, offset, amplitude, slope, centre, period):
    return offset + slope * (x - centre) + amplitude * np.cos(2 * np.pi * (x - centre) / period)


def _odd(n: float, minimum: int = 5) -> int:
    n = max(minimum, int(n))
    return n if n % 2 else n + 1


def _uniform(nu: np.ndarray, y: np.ndarray):
    """Sort by wavenumber, merge repeated points and resample onto a uniform grid."""
    order = np.argsort(nu)
    nu, y = nu[order], y[order]
    nu, inverse = np.unique(nu, return_inverse=True)
    y = np.bincount(inverse, weights=y) / np.bincount(inverse)
    step = np.median(np.diff(nu))
    grid = np.arange(nu[0], nu[-1], step)
    return grid, np.interp(grid, nu, y), step


def normalise_segment(nu: np.ndarray, y: np.ndarray, period: float, s: Settings):
    """Resample a segment onto a uniform wavenumber grid and divide by its baseline."""
    if s.edge_trim and len(nu) > 4 * s.edge_trim:
        nu, y = nu[s.edge_trim:-s.edge_trim], y[s.edge_trim:-s.edge_trim]
    grid, yi, step = _uniform(nu, y)
    points_per_period = period / step
    window = min(_odd(s.baseline_periods * points_per_period), len(yi) - (1 - len(yi) % 2))
    baseline = savgol_filter(yi, window, 1)
    with np.errstate(divide="ignore", invalid="ignore"):
        norm = yi / baseline
    return grid, norm, points_per_period


def estimate_period(nu: np.ndarray, y: np.ndarray) -> float | None:
    """Fringe period (1/um) from the dominant Fourier component of one segment."""
    if len(nu) < 64:
        return None
    grid, yi, step = _uniform(nu, y)
    r = yi / savgol_filter(yi, min(_odd(len(yi) // 8), len(yi) - (1 - len(yi) % 2)), 2) - 1
    r = np.nan_to_num(r)
    power = np.abs(np.fft.rfft(r * np.hanning(len(r)))) ** 2
    freq = np.fft.rfftfreq(len(r), step)
    lo = np.searchsorted(freq, 5 / (grid[-1] - grid[0]))  # at least five fringes per segment
    hi = len(freq) // 2                                      # at least ~4 points per fringe
    if hi <= lo:
        return None
    return 1 / freq[lo + np.argmax(power[lo:hi])]


def fit_segment(nu, y, period, segment, s: Settings) -> list[dict]:
    """Fit every fringe in one segment (nu in 1/um)."""
    grid, norm, ppp = normalise_segment(nu, y, period, s)
    r = norm - 1
    if s.minima:
        r = -r
    if ppp < 4:
        log.warning("segment %d: only %.1f points per fringe; skipped", segment, ppp)
        return []
    smooth = savgol_filter(r, min(_odd(s.smoothing_periods * ppp), len(r) - (1 - len(r) % 2)), 2)
    peaks, _ = find_peaks(smooth, distance=max(1, int(s.min_peak_separation * ppp)))
    half = int(s.fit_halfwidth * ppp)
    rows = []
    for p in peaks:
        lo, hi = p - half, p + half + 1
        if lo < 0 or hi > len(grid):
            continue
        x, yy = grid[lo:hi], r[lo:hi]
        if not np.all(np.isfinite(yy)):
            continue
        try:
            if s.cosine_check:
                pc, _ = curve_fit(cosine, x, yy, p0=[0, np.ptp(yy) / 2, 0, grid[p], period], maxfev=4000)
                if not (x[0] < pc[3] < x[-1]) or pc[1] <= 0 or abs(pc[4] / period - 1) > s.period_tolerance:
                    continue
            pg, cov = curve_fit(gaussian, x, yy, p0=[np.min(yy), np.ptp(yy), 0, grid[p], period / 4], maxfev=4000)
        except (RuntimeError, ValueError):
            continue
        if not (x[0] < pg[3] < x[-1]):
            continue
        if not s.cosine_check and (pg[1] <= 0 or not 0.02 * period < abs(pg[4]) < period):
            continue
        rows.append(dict(
            segment=segment,
            wavelength_nm=1000.0 / pg[3],
            wavenumber_um_inv=pg[3],
            wavenumber_cm_inv=1.0e4 * pg[3],
            wavenumber_err_um_inv=float(np.sqrt(cov[3, 3])) if np.isfinite(cov[3, 3]) else np.nan,
            gaussian_amplitude=pg[1],
            gaussian_width_um_inv=abs(pg[4]),
            cosine_wavenumber_um_inv=pc[3] if s.cosine_check else np.nan,
            gauss_minus_cosine_um_inv=pg[3] - pc[3] if s.cosine_check else np.nan,
            fit_rms=float(np.std(yy - gaussian(x, *pg))),
        ))
    return rows


def fit_fringes(x, y, x_unit: str = "nm", settings: Settings | None = None,
                period: float | None = None, return_curves: bool = False):
    """Fit all fringes in a spectrum.

    Parameters
    ----------
    x, y : array-like
        x values (wavelength or wavenumber, in any order) and intensity.
    x_unit : str
        Unit of x: "nm", "um", "angstrom", "cm-1" or "um-1".
    settings : Settings, optional
        Analysis settings; the defaults are used if omitted.
    period : float, optional
        Fringe period in 1/um, i.e. 1/(2 n t) for optical path n t in um.
        Estimated from the spectrum if omitted.

    Returns a DataFrame with one row per fringe, sorted by wavenumber
    (and, if ``return_curves``, the normalised segments and the period used).
    """
    s = settings or Settings()
    x, y = np.asarray(x, float), np.asarray(y, float)
    ok = np.isfinite(x) & np.isfinite(y)
    x, y = x[ok], y[ok]
    order = np.argsort(x, kind="stable")
    x, y = x[order], y[order]
    segments = [(to_wavenumber(a, x_unit), b) for a, b in split_segments(x, y, s)
                if len(a) > 2 * s.edge_trim + 10]
    if not segments:
        raise ValueError("no usable segments")
    if period is None and not s.period_per_segment:
        estimates = [p for p in (estimate_period(a, b) for a, b in segments) if p]
        if not estimates:
            raise ValueError("could not estimate the fringe period; give it explicitly")
        period = float(np.median(estimates))
    rows, curves, periods = [], [], []
    for i, (nu, yy) in enumerate(segments):
        p = period if period is not None else estimate_period(nu, yy)
        if not p:
            log.warning("segment %d: no fringe period found; skipped", i)
            continue
        periods.append(p)
        rows += fit_segment(nu, yy, p, i, s)
        grid, norm, _ = normalise_segment(nu, yy, p, s)
        curves.append((i, 1000.0 / grid, norm))
    columns = ["segment", "wavelength_nm", "wavenumber_um_inv", "wavenumber_cm_inv", "wavenumber_err_um_inv",
               "gaussian_amplitude", "gaussian_width_um_inv", "cosine_wavenumber_um_inv",
               "gauss_minus_cosine_um_inv", "fit_rms"]
    table = pd.DataFrame(rows, columns=columns).sort_values("wavenumber_um_inv", ignore_index=True)
    if return_curves:
        return table, curves, float(np.median(periods)) if periods else np.nan
    return table


# ----------------------------------------------------------------------------
# Output
# ----------------------------------------------------------------------------
def plot_overview(name: str, table: pd.DataFrame, curves, path: Path) -> None:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    check = table.gauss_minus_cosine_um_inv.notna().any()
    fig, axes = plt.subplots(2 if check else 1, 1, figsize=(9, 6 if check else 4), sharex=True,
                             gridspec_kw=dict(height_ratios=[2, 1]) if check else None, squeeze=False)
    a = axes[0, 0]
    for i, lam, norm in curves:
        a.plot(lam, norm, lw=0.5, color="0.15")
        sel = table[table.segment == i]
        a.plot(sel.wavelength_nm, np.interp(sel.wavelength_nm, lam[::-1], norm[::-1]), "v", ms=3,
               color="tab:orange", mec="none")
    a.set_ylabel("Normalised intensity")
    a.set_title(f"{name}: {len(table)} fringes", loc="left")
    if check and len(table) > 1:
        b = axes[1, 0]
        spacing = np.median(np.diff(table.wavenumber_um_inv))
        b.plot(table.wavelength_nm, 100 * table.gauss_minus_cosine_um_inv / spacing, "o", ms=2.5, color="tab:blue")
        b.axhline(0, color="0.6", lw=0.6)
        b.set_ylabel("Gaussian − cosine centre\n(% of fringe spacing)")
    axes[-1, 0].set_xlabel("Wavelength (nm)")
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def main() -> None:
    p = argparse.ArgumentParser(description="Fit interference fringes in optical spectra with Gaussians.",
                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("input", type=Path, help="spectrum file (.csv, .txt, .dat, .xlsx, .xls)")
    p.add_argument("--out", type=Path, default=Path("results"), help="output folder")
    p.add_argument("--sheet", help="process only this sheet of a workbook (default: all sheets)")
    p.add_argument("--x-col", type=int, default=0, help="column of the x values, counting from 0")
    p.add_argument("--y-col", default="1", help="intensity column(s): e.g. 1, or 1,2,3, or all")
    p.add_argument("--x-unit", default="nm", choices=X_UNITS, help="unit of the x values")
    group = p.add_mutually_exclusive_group()
    group.add_argument("--period", type=float, help="fringe period in 1/um (default: estimated)")
    group.add_argument("--optical-path", type=float, help="approximate optical path n*t in um; sets the period to 1/(2nt)")
    p.add_argument("--minima", action="store_true", help="fit fringe minima instead of maxima")
    p.add_argument("--no-split", action="store_true", help="treat each spectrum as one continuous segment")
    p.add_argument("--no-cosine-check", action="store_true", help="accept fringes on the Gaussian fit alone")
    p.add_argument("--period-per-segment", action="store_true", help="estimate the period separately for each segment")
    p.add_argument("--baseline-periods", type=float, default=Settings.baseline_periods,
                   help="baseline smoothing window, in fringe periods")
    p.add_argument("--edge-trim", type=int, default=Settings.edge_trim, help="points dropped at each end of a segment")
    p.add_argument("--no-plot", action="store_true", help="do not write overview plots")
    args = p.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    s = Settings(split_segments=not args.no_split, cosine_check=not args.no_cosine_check, minima=args.minima,
                 period_per_segment=args.period_per_segment, baseline_periods=args.baseline_periods,
                 edge_trim=args.edge_trim)
    period = args.period or (1 / (2 * args.optical_path) if args.optical_path else None)
    args.out.mkdir(parents=True, exist_ok=True)
    for name, (x, y) in read_spectra(args.input, args.sheet, args.x_col, args.y_col).items():
        try:
            table, curves, used = fit_fringes(x, y, args.x_unit, s, period, return_curves=True)
        except ValueError as err:
            log.error("%s: %s", name, err)
            continue
        safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
        table.to_csv(args.out / f"{safe}_fringes.csv", index=False)
        if not args.no_plot:
            plot_overview(name, table, curves, args.out / f"{safe}_overview.png")
        msg = (f"{name}: {len(curves)} segment(s), {len(table)} fringes, "
               f"{table.wavelength_nm.min():.0f}-{table.wavelength_nm.max():.0f} nm, period {used:.5f} 1/um")
        if len(table) > 1 and table.gauss_minus_cosine_um_inv.notna().any():
            spacing = np.median(np.diff(table.wavenumber_um_inv))
            msg += f"; median |Gaussian - cosine| = {100 * np.median(np.abs(table.gauss_minus_cosine_um_inv)) / spacing:.2f}% of the spacing"
        log.info(msg)


if __name__ == "__main__":
    main()
