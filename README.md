# Fringe fitting for optical spectra

Finds and fits the interference fringes in transmission or reflection spectra of thin transparent layers, such as the sample chamber of a diamond anvil cell (DAC), a film or a plate. Each fringe is fitted with a Gaussian to give its centre.

This is the fringe-fitting step used in R. O'Beirne and J. E. Proctor, *Dispersion of the Refractive Index of KCl up to 5 GPa Measured Using Confocal Laser Microscopy and White-Light Interferometry* (submitted). With its default settings it reproduces the 203 fringe centres used for the 0.68 GPa KCl spectrum in the paper exactly. It is written to work on other fringe spectra without changes to the code.

## What it does

For each spectrum, `fringe_fitting.py`:

1. reads the data from a text, CSV or Excel file. The x axis may be wavelength (nm, µm or Å) or wavenumber (cm⁻¹ or µm⁻¹), in either order, and evenly or unevenly spaced;
2. splits the spectrum into continuous segments. A spectrum merged from several overlapping detector windows is split back into its windows. Overlap regions, where points from two windows interleave, are discarded, and nothing is interpolated across gaps. A single continuous spectrum stays as one segment;
3. converts each segment to wavenumber, where the fringes are evenly spaced, and divides it by a smooth Savitzky–Golay baseline;
4. locates each fringe maximum (or minimum) and fits it, over one fringe period, with a Gaussian on a linear background. The Gaussian is used only to locate the centre, not as a model of the Fabry–Perot line shape. By default a cosine fit over the same range is also made; it rejects poor fringes and checks the centres;
5. writes a table of fringe centres and an overview plot.

The fringe period sets the size of the baseline and fit windows. It is estimated automatically from the strongest Fourier component of the spectrum. It can also be given directly, or from the approximate optical path *nt*.

Assigning fringe orders and calculating the refractive index from the centres are separate steps that are not included here.

## Installation

Python 3.9 or later.

```
pip install -r requirements.txt
```

## Command line

```
python fringe_fitting.py examples/KCl_0.68GPa_full_range.csv --out results
```

Some other examples:

```
# several spectra in one file: x in column 0, intensities in all other columns
python fringe_fitting.py data.csv --y-col all

# reflection spectrum against wavenumber in cm-1: fit the fringe minima
python fringe_fitting.py reflection.txt --x-unit cm-1 --minima

# a single continuous spectrum with sharper (higher-finesse) fringes
python fringe_fitting.py film.dat --no-split --no-cosine-check

# one sheet of a workbook, with the period set from an optical path of about 125 um
python fringe_fitting.py spectra.xlsx --sheet Data3 --optical-path 125
```

| Option | Meaning |
| --- | --- |
| `--out FOLDER` | output folder (default `results`) |
| `--sheet NAME` | process only this sheet of a workbook (default: every sheet) |
| `--x-col N` | column of the x values, counting from 0 (default 0) |
| `--y-col LIST` | intensity column(s): `1`, `1,2,3` or `all` (default 1) |
| `--x-unit UNIT` | `nm` (default), `um`, `angstrom`, `cm-1` or `um-1` |
| `--period P` | fringe period in µm⁻¹ |
| `--optical-path UM` | approximate optical path *nt* in µm; sets the period to 1/(2*nt*) |
| `--minima` | fit fringe minima instead of maxima |
| `--no-split` | treat each spectrum as one continuous segment |
| `--no-cosine-check` | accept fringes on the Gaussian fit alone (useful for sharper fringes) |
| `--period-per-segment` | estimate the period separately for each segment |
| `--baseline-periods N` | baseline smoothing window in fringe periods (default 2.5) |
| `--edge-trim N` | points dropped at each end of a segment (default 20) |
| `--no-plot` | do not write overview plots |

Further settings are in the `Settings` class at the top of the script. The defaults are the values used for the paper.

## Python

```python
from fringe_fitting import fit_fringes, read_spectra, Settings

table = fit_fringes(wavelength_nm, intensity)                  # numpy arrays or lists
table = fit_fringes(wavenumber_cm, intensity, x_unit="cm-1",
                    settings=Settings(minima=True))

for name, (x, y) in read_spectra("spectra.xlsx").items():
    print(name, len(fit_fringes(x, y)))
```

## Input files

- **Formats:** `.csv`, `.txt`, `.dat` (comma, semicolon, tab or space separated) and `.xlsx`/`.xls` workbooks. In a workbook each sheet is a separate spectrum.
- **Headers:** header and comment lines are skipped automatically. Any row whose x and intensity are not numbers is ignored.
- **Columns:** by default x is in the first column and intensity in the second; use `--x-col` and `--y-col` otherwise.

## Output

For each spectrum:

- `<name>_fringes.csv`: one row per fringe, sorted by wavenumber.

  | Column | Meaning |
  | --- | --- |
  | `segment` | continuous segment (detector window) the fringe belongs to |
  | `wavelength_nm`, `wavenumber_um_inv`, `wavenumber_cm_inv` | fitted fringe centre |
  | `wavenumber_err_um_inv` | statistical uncertainty of the centre |
  | `gaussian_amplitude`, `gaussian_width_um_inv` | height (normalised units) and width of the Gaussian |
  | `cosine_wavenumber_um_inv` | centre from the cosine check fit |
  | `gauss_minus_cosine_um_inv` | difference between the two centres |
  | `fit_rms` | rms residual of the Gaussian fit |

- `<name>_overview.png`: the normalised spectrum with the fitted centres marked and, when the cosine check is on, the Gaussian − cosine difference for each fringe.

Always check the overview plot for a new dataset. A few fitted fringes may be poor, for example at segment edges or where a fringe is distorted. In the paper such fringes were removed at the next stage, when fringe orders were assigned: fringes more than 0.15 nm from the smooth fringe-order curve were rejected.

## Working range

Tested on synthetic Fabry–Perot spectra with known fringe positions (`tests/`), across 500–800 nm with 6000 points:

| Case | Median centre error (fraction of the fringe spacing) |
| --- | --- |
| 3–50 µm thick layer (7–115 fringes) | below 0.1% |
| 100–200 µm (14–28 points per fringe) | 0.1–0.2% |
| 400 µm (7 points per fringe) | 0.3% |
| 1% noise | 0.3% |
| 3% noise | 1% |
| sharper fringes (finesse ≈ 5) | below 0.1% |

At least about five points per fringe and five fringes per segment are needed.

Run the tests with:

```
python -m pytest tests
```

## Example data

`examples/KCl_0.68GPa_full_range.csv` is the KCl spectrum at 0.68 GPa used in Figs. 4 and 5 of the paper. It contains six detector windows covering 419–772 nm, with a gap between 530 and 550 nm.

## Earlier version

The `legacy` folder contains the notebook and single-window example from the original submission. That version calculated the optical path δ from pairs of fringes. For a dispersive sample this gives the group optical path *n*<sub>g</sub>*t* rather than *nt*, so it has been superseded by the analysis described in the paper.

## Licence

MIT; see `LICENSE`.
