# GelBandPicker

GelBandPicker is a small graphical tool for manual gel band quantification.

It is designed for cases where fully automatic band calling is too fragile. The software keeps the useful automatic part, lane-center detection, while letting users crop the gel, adjust lane centers, and click the actual bands by hand. It then exports background-corrected integrated intensity tables, band fractions, QC images, and a reproducible session file.

## Research Use Notice

GelBandPicker is intended as a research aid. Users should inspect the QC outputs and verify quantification results before using them for reporting or publication.

GelBandPicker performs semi-quantitative densitometry. It is not a substitute for experimental controls, linear-range validation, replicate analysis, or user review of the measured regions of interest.

## Features

- Manual band picking with automatic lane-center detection.
- Crop, pan, zoom, and contrast adjustment inside the GUI.
- Editable layout: drag lane centers and band legends, double-click to add lanes, right-click to delete lanes.
- Undo and redo for manual band and lane edits.
- Configurable ROI size for band quantification.
- Two-line lane labels from `.xlsx`, `.csv`, or `.tsv` annotation tables.
- Local background subtraction from flanking regions.
- CSV/XLSX output plus QC images.
- Cross-platform Python script with Windows and macOS launch helpers.

## Installation

GelBandPicker requires Python 3 and these packages:

- `numpy`
- `pandas`
- `pillow`
- `openpyxl`

Install them manually:

```bash
pip install -r requirements.txt
```

GelBandPicker does not install dependencies automatically. If startup reports missing packages, run the command above from the GelBandPicker folder.

## Start

Windows double-click:

```bash
start_gui.bat
```

macOS:

```bash
start_gui.command
```

macOS security settings may block double-click launchers from downloaded folders. In that case, open Terminal in the GelBandPicker folder and run:

```bash
python gel_gui.py
```

Some systems use:

```bash
python3 gel_gui.py
```

## Workflow

1. Choose an image, annotation table, and output folder in `Files`. GelBandPicker does not include a sample gel image; use your own TIFF/PNG/JPG image.
2. Click the large `LOAD IMAGE` button on the right.
3. Use the `Workflow` bar for mode, crop, auto lanes, and zoom.
4. The default mode is `Move`. Drag the image to pan without adding labels or bands.
5. Use the mouse wheel to scroll vertically. Use `Shift` + mouse wheel to scroll horizontally.
6. To crop, choose `Crop`, hold the left mouse button and drag a yellow rectangle, release, then click `Apply crop`.
7. Choose `Edit layout` mode:
   - drag a vertical lane line to adjust it;
   - double-click to add a lane;
   - right-click a lane line to delete it;
   - drag a left-side `band x` legend to place it manually;
   - right-click or double-click a legend to return it to automatic placement.
8. In `Band picking and ROI`, set `Bands` to the number of band levels to quantify.
9. Choose `Pick bands` mode.
10. Select `Band 1`, then click each matching band on the image.
11. Select `Band 2`, `Band 3`, etc., and click their bands.
12. Use `Undo`/`Redo` for corrections. Right-click also deletes a clicked point where secondary click is available.
13. Click the large `RUN QUANTIFICATION` button to the right of `Style and display`.

`Reset` returns to the original uncropped image, reloads lane annotations, reruns lane detection, restores the default `Label y`, and clears manually clicked band points.

## Annotation Table

Supported formats:

- `.xlsx`
- `.csv`
- `.tsv`

An annotation table is required. A starter CSV template is included:

```text
annotation_template.csv
```

Default columns:

- `lane_index`: physical lane number from left to right.
- `include`: `TRUE` to quantify; `FALSE` for marker or ignored lanes.
- `lane_label`: short lane label.
- `x_center`: optional manual lane center in pixels.
- `top_label`: default first label line, such as sample ID, construct ID, or group name.
- `bottom_label`: default second label line, such as condition, treatment, construct detail, or any other sample note.
- `note`: optional note copied to output tables.

Rows are sorted by `lane_index`. Set `include` to `FALSE` for marker lanes or lanes that should be shown in QC but excluded from intensity totals.

`Polarity` controls how pixel intensity is interpreted for quantification:

- `bright`: use this for bright fluorescent/chemiluminescent bands on a dark background.
- `dark`: use this for dark bands on a light background, such as some stained gels or inverted images.

Choose `Polarity` according to the original image data, not according to the inverted display. For example, if the original image has bright bands on a dark background, keep `Polarity` as `bright` even after turning on `Invert`.

Changing `Polarity` changes the numerical signal used for lane detection and band quantification. It is different from `Invert`, which only changes display/export appearance.

## Display Controls

`Black` and `White` adjust display contrast in real time. They affect visualization and exported figures only.

`Invert` switches the display and exported figures between dark-background and white-background views.

These display controls do not change the numerical intensity calculations.

## ROI Controls

`Band half width` is the horizontal ROI half-width. From the lane center, GelBandPicker measures this many pixels to the left and right.

`Band half height` is the vertical ROI half-height around the clicked band center.

The clicked band box shown directly on the image is the measured ROI. Its middle line is the clicked/refined band center.

`Marker line width` controls the line thickness in the GUI and QC image. It does not change the quantified pixel area.

When an image is loaded, GelBandPicker initializes `Label y` near the loading wells. Use `Label y`, `Top font`, and `Bottom font` to manually adjust lane-label position and size.

Use `Band font` to adjust the size of band labels shown next to clicked band points and in the left-side band legend. By default, each legend is placed at the average y-position of all clicked points for that band. In `Edit layout` mode, drag a legend to set a manual position; right-click or double-click it to return to the automatic average.

## Quantification Method

GelBandPicker quantifies manually clicked bands as local background-subtracted integrated density.

Let `G(x, y)` be the original grayscale image value at pixel `(x, y)`.

For bright bands on a dark background:

```text
S(x, y) = G(x, y)
```

For dark bands on a light background:

```text
S(x, y) = max(P99.8(G) - G(x, y), 0)
```

where `P99.8(G)` is the 99.8th percentile of grayscale values in the image.

For lane `i` and band `b`, the user-clicked ROI is:

```text
R[i,b]
```

The local background is estimated from the left and right flanking background regions next to the same lane and same vertical band window:

```text
B[i,b] = B_left[i,b] union B_right[i,b]
```

The background value is the median signal in the flanking regions:

```text
bg[i,b] = median( S(x, y) for (x, y) in B[i,b] )
```

The background-corrected integrated intensity is:

```text
I[i,b] = sum( max(S(x, y) - bg[i,b], 0) for (x, y) in R[i,b] )
```

The mean corrected signal is:

```text
mean_corrected_signal[i,b] = I[i,b] / area(R[i,b])
```

For within-lane band fractions:

```text
fraction[i,b] = I[i,b] / sum(I[i,k] for all clicked bands k in lane i)
```

Unclicked bands are exported as zero.

## Output Columns

`*_manual_long.csv` contains one row per lane per band.

Important columns:

- `raw_gray_mean`: mean original grayscale value in the ROI.
- `raw_gray_integrated`: sum of original grayscale values in the ROI.
- `background_signal_median`: local background median from flanking regions.
- `background_corrected_integrated`: main corrected integrated intensity, `I[i,b]`.
- `mean_corrected_signal`: corrected integrated intensity divided by ROI area.
- `peak_profile_value`: signal used for peak refinement around the clicked point.

`*_manual_summary.csv` and `*_manual_summary.xlsx` contain one row per lane with:

- `band_1_integrated`, `band_2_integrated`, etc.
- `total_signal`
- `band_1_fraction`, `band_2_fraction`, etc.

## Outputs

For each image, the output folder receives:

- `*_manual_long.csv`
- `*_manual_summary.csv`
- `*_manual_summary.xlsx`
- `*_manual_qc.png`
- `gel_gui_session.json`

The QC image shows the measured ROI boxes and should be inspected before using the values.

The session JSON stores the crop, lane positions, clicked band points, and GUI settings so the analysis can be reviewed or repeated.

## Good Practice

- Prefer original 16-bit TIFF images for quantification.
- Avoid saturated bands; saturated pixels break densitometric linearity.
- Keep ROI settings consistent within a gel.
- Inspect the QC image for every analyzed gel.
- Use biological and technical replicates for robust conclusions.
- For comparing total signal between lanes, use appropriate loading or input controls.
- For comparing band distributions within the same lane, use the exported band fractions.

## Limitations

- GelBandPicker provides semi-quantitative densitometry, not absolute molecular quantification.
- Background subtraction is local and rectangular; unusual smears or uneven backgrounds may require manual review.
- The tool does not determine whether an image is in the linear detection range.
- The user remains responsible for verifying the ROI placement and scientific interpretation.

## License

Apache License 2.0. See [LICENSE](LICENSE).
