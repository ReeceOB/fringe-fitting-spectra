# fringe-fitting-spectra
Python script for fitting spectral fringes and calculating delta values from fitted fringe centres (see attached publication)

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
- saves peak tables, delta summaries, and overview plots

Fringe numbers are assigned once after initial peak selection and are **not renumbered later**.

## Input format

Each worksheet must use this fixed layout:
- **column 0** = wavelength in **nm**
- **column 1** = intensity
- data start at **row index 2** (the first two rows are skipped)

## How to run

1. Set the file paths at the top of the script:
   - `INPUT_WORKBOOK`
   - `OUTPUT_DIR`

2. Check the analysis settings in `FringeConfig`.

3. Run the script in Python.  
   It will process **all sheets** in the workbook.

## What you may need to change for your own data

### File locations
Edit:
- `INPUT_WORKBOOK`
- `OUTPUT_DIR`

### Main fitting and detection settings
These are the most likely parameters to adjust:

- `avg_peak_spacing_um_inv`  
  approximate fringe spacing; affects fitting window size

- `peak_prominence`  
  increase if too many weak peaks are detected; decrease if real peaks are missed

- `min_peak_spacing_um_inv`  
  minimum allowed spacing between neighbouring peaks

- `r2_min_processed`  
  Stage 1 fit-quality threshold 

- `r2_min_raw`  
  Stage 2 fit-quality threshold

- `width_min_um_inv`, `width_max_um_inv`  
  allowed Gaussian width range

- `sg_window`, `sg_poly`  
  smoothing strength for the Savitzky-Golay filter

### Baseline settings
If the background shape changes significantly between datasets, you may also need to adjust:
- `baseline_lambda`
- `baseline_p`

## Output

For each sheet, the script writes:
- `<sheet>_Peaks.xlsx`  
  with `Raw` and `Processed` peak tables
- `<sheet>_Delta.xlsx`  
  with per-sheet delta values
- `<sheet>_Overview.png`
- `<sheet>_Overview_Processed.png`

It also writes:
- `Delta_Summary.xlsx`  
  containing one summary row per sheet

## Notes

- Always check the overview plots when using a new dataset.
- If fitting fails for a sheet, the script still writes empty output files and records the sheet status.
- The fixed worksheet layout is assumed throughout this version of the script.
