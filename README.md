# Quick guide

This script processes spectra stored in a **multi-sheet Excel workbook** and fits fringe peaks with a **two-stage Gaussian workflow**.

For each sheet, it:
- reads wavelength and intensity data
- converts wavelength (nm) to wavenumber (µm⁻¹)
- subtracts the baseline
- smooths the spectrum with Savitzky-Golay
- detects candidate peaks
- fits Gaussians on the smoothed spectrum (**Stage 1**)
- refits accepted peaks on the raw baseline-subtracted spectrum (**Stage 2**)
- removes overlapping fits using a spacing filter
- calculates delta values from the fitted fringe centres
- calculates a **dispersion table** containing all valid delta values as a function of average wavelength
- saves peak tables, delta summaries, dispersion outputs, and overview plots

Fringe numbers are assigned once after initial peak selection and are **not renumbered later**.

## Input format

Each worksheet must use this fixed layout:
- **column 0** = wavelength in **nm**
- **column 1** = intensity
- data start at **row index 2** (the first two rows are skipped)

## How to run

1. Set the file paths at the top of the script:
   - `INPUT_FILE`
   - `OUTPUT_DIR`

2. Check the analysis settings in `FringeConfig`.

3. Run the script in Python.  
   It will process **all sheets** in the workbook.

## What you may need to change for your own data

### File locations
Edit:
- `INPUT_FILE`
- `OUTPUT_DIR`

### Main fitting and detection settings
These are the most likely parameters to adjust:

- `avg_peak_spacing_um_inv`  
  approximate fringe spacing; affects fitting window size

- `peak_prominence`  
  increase if too many weak peaks are detected; decrease if real peaks are missed

- `min_peak_spacing_um_inv`  
  minimum allowed spacing between neighbouring fitted peaks in spectral units

- `r2_min_processed`  
  Stage 1 fit-quality threshold

- `r2_min_raw`  
  Stage 2 fit-quality threshold

- `width_min_um_inv`, `width_max_um_inv`  
  allowed Gaussian width range

- `sg_window`, `sg_poly`  
  smoothing strength for the Savitzky-Golay filter

### Delta / dispersion settings
These settings control which fringe pairs are used for delta calculations:

- `p_min_index_spacing`  
  minimum allowed **fringe-index spacing** between the two fringes used to calculate a delta value

  Example: if `p_min_index_spacing = 5`, then pairs with fringe-number gaps smaller than 5 are excluded from delta and dispersion calculations.

The scalar delta summaries and the dispersion outputs are generated from the same valid pair set.

### Baseline settings
If the background shape changes significantly between datasets, you may also need to adjust:
- `baseline_lambda`
- `baseline_p`

## Output

For each sheet, the script writes:
- `<sheet>_Peaks.xlsx`  
  with `Raw` and `Processed` peak tables
- `<sheet>_Delta.xlsx`  
  with per-sheet delta summary values
- `<sheet>_Dispersion.xlsx`  
  with `Raw` and `Processed` sheets containing all valid delta values versus average wavelength
- `<sheet>_Overview.png`
- `<sheet>_Overview_Processed.png`

It also writes:
- `Delta_Summary.xlsx`  
  containing one summary row per sheet
- `Dispersion_Summary.xlsx`  
  containing one worksheet per dataset, with that dataset’s dispersion results

## Dispersion output

The dispersion tables contain the individual delta values used for dispersion analysis rather than only a single averaged delta.

Each dispersion row corresponds to one valid fringe pair and includes:
- left fringe number
- right fringe number
- index spacing
- pair order
- left wavelength
- right wavelength
- average wavelength
- delta

The **average wavelength** is used as the x-value for dispersion analysis.

## Notes

- Always check the overview plots when using a new dataset.
- If fitting fails for a sheet, the script still writes empty output files and records the sheet status.
- Fringe numbers are never renumbered after initial assignment, even if some fringes are later rejected.
- The fixed worksheet layout is assumed throughout this version of the script.
