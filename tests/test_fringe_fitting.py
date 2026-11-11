"""Tests on synthetic Fabry-Perot spectra with known fringe positions.

Run with:  python -m pytest tests
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from scipy.optimize import brentq

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fringe_fitting as ff  # noqa: E402


def refractive_index(nu):
    """A dispersive index (Cauchy form), nu in 1/um."""
    return 1.5 + 0.004 * nu ** 2


def airy(nu, t, finesse_coeff=0.25):
    """Fabry-Perot transmission of a layer of thickness t (um)."""
    phase = 2 * np.pi * refractive_index(nu) * t * nu
    return 1 / (1 + finesse_coeff * np.sin(phase) ** 2)


def true_maxima(t, nu_min, nu_max):
    """Wavenumbers where 2 n t nu is an integer (transmission maxima)."""
    order = lambda nu: 2 * refractive_index(nu) * t * nu
    ms = np.arange(np.ceil(order(nu_min)), np.floor(order(nu_max)) + 1)
    return np.array([brentq(lambda nu: order(nu) - m, nu_min, nu_max) for m in ms])


def synthetic(wavelength_nm, t, noise=0.002, finesse_coeff=0.25, seed=0):
    rng = np.random.default_rng(seed)
    nu = 1000 / wavelength_nm
    lamp = 1 + 0.5 * np.exp(-((wavelength_nm - 600) / 250) ** 2)  # smooth lamp spectrum
    return 1e4 * lamp * airy(nu, t, finesse_coeff) * (1 + noise * rng.standard_normal(len(nu)))


def check(table, t, lam_min, lam_max, tol=0.02):
    """Every fitted centre lies within tol fringe spacings of a true maximum."""
    truth = true_maxima(t, 1000 / lam_max, 1000 / lam_min)
    spacing = np.median(np.diff(truth))
    nu = table.wavenumber_um_inv.to_numpy()
    err = np.array([np.min(np.abs(truth - v)) for v in nu]) / spacing
    assert len(nu) > 0.8 * len(truth) - 4, (len(nu), len(truth))
    assert np.max(err) < tol, np.max(err)
    return len(nu), len(truth)


def test_wavelength_csv_with_headers(tmp_path):
    lam = np.linspace(500, 800, 6000)
    f = tmp_path / "spectrum.csv"
    with open(f, "w") as fh:
        fh.write("Wavelength,Counts\nnm,a.u.\n")
        for a, b in zip(lam, synthetic(lam, 40)):
            fh.write(f"{a:.4f},{b:.2f}\n")
    (name, (x, y)), = ff.read_spectra(f).items()
    check(ff.fit_fringes(x, y, "nm"), 40, 500, 800)


def test_wavenumber_text_no_header_descending_minima(tmp_path):
    """Reflection-like spectrum (fringe minima), x in cm-1, whitespace separated, descending order."""
    lam = np.linspace(450, 750, 5000)
    wn = 1e7 / lam
    refl = 2e4 - synthetic(lam, 60)
    f = tmp_path / "spectrum.txt"
    np.savetxt(f, np.column_stack([wn, refl]), fmt="%.5f")
    (name, (x, y)), = ff.read_spectra(f).items()
    table = ff.fit_fringes(x, y, "cm-1", ff.Settings(minima=True))
    check(table, 60, 450, 750)


def test_multiple_columns_and_um_units(tmp_path):
    lam_um = np.linspace(0.45, 0.80, 7000)
    df = pd.DataFrame({"x": lam_um, "a": synthetic(lam_um * 1000, 30, seed=1),
                       "b": synthetic(lam_um * 1000, 80, seed=2)})
    f = tmp_path / "multi.csv"
    with open(f, "w") as fh:
        fh.write("# two samples\n")
        df.to_csv(fh, sep=";", index=False)
    spectra = ff.read_spectra(f, y_cols="all")
    assert len(spectra) == 2
    for (name, (x, y)), t in zip(spectra.items(), (30, 80)):
        check(ff.fit_fringes(x, y, "um"), t, 450, 800)


def test_merged_overlapping_windows():
    """Two detector windows that overlap by about 10 nm, merged and sorted."""
    w1 = np.linspace(500, 620, 1024)
    w2 = np.linspace(610, 730, 1024)
    lam = np.concatenate([w1, w2])
    y = synthetic(lam, 50)
    segments = ff.split_segments(np.sort(lam), y[np.argsort(lam)], ff.Settings())
    assert len(segments) == 2
    assert segments[0][0][-1] < 610 and segments[1][0][0] > 620
    order = np.argsort(lam)
    check(ff.fit_fringes(lam[order], y[order], "nm"), 50, 500, 730)


def test_uneven_sampling_and_higher_finesse():
    rng = np.random.default_rng(3)
    lam = np.sort(rng.uniform(550, 750, 6000))
    y = synthetic(lam, 45, finesse_coeff=3)
    table = ff.fit_fringes(lam, y, "nm", ff.Settings(cosine_check=False, split_segments=False))
    check(table, 45, 550, 750, tol=0.03)


def test_excel_workbook(tmp_path):
    lam = np.linspace(500, 760, 5000)
    f = tmp_path / "book.xlsx"
    with pd.ExcelWriter(f) as w:
        for i, t in enumerate((35, 55)):
            pd.DataFrame({"Wavelength (nm)": lam, "Intensity": synthetic(lam, t, seed=i)}).to_excel(
                w, sheet_name=f"P{i}", index=False)
    spectra = ff.read_spectra(f)
    assert list(spectra) == ["P0", "P1"]
    for (name, (x, y)), t in zip(spectra.items(), (35, 55)):
        check(ff.fit_fringes(x, y), t, 500, 760)


def test_example_spectrum():
    """The KCl example: six detector windows and about 200 fringes."""
    f = Path(__file__).resolve().parents[1] / "examples" / "KCl_0.68GPa_full_range.csv"
    (name, (x, y)), = ff.read_spectra(f).items()
    table, curves, period = ff.fit_fringes(x, y, return_curves=True)
    assert len(curves) == 6
    assert len(table) == 203
    assert abs(1 / (2 * period) - 125.2) < 2
