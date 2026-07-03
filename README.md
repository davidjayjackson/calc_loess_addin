# Calc LOESS/LOWESS Add-in

A LOESS/LOWESS (locally weighted polynomial regression) smoothing function for
LibreOffice Calc, implemented as a real Python UNO Add-In (`src/loess_impl.py`).
Being a genuine Add-In (`com.sun.star.sheet.AddIn`), rather than a Basic macro
merely callable from a formula, `LOESS` is listed in the Function Wizard and
autocompletes as you type.

## What it does

`LOESS()` fits a local polynomial (constant, linear, or quadratic) around each
requested x-value, weighted by a tricube kernel over the nearest neighbours,
optionally refined with Cleveland-style robustness iterations that down-weight
outliers. This is the same family of algorithm as R's `lowess()`/`loess()`.

## Installation

`LOESS` is a real UNO Add-In (`com.sun.star.sheet.AddIn`), not a Basic macro -
that's what gets it into the Function Wizard and formula autocomplete. A
plain Basic macro used as a cell formula is resolved by name only at
calculation time, so Calc never knows its argument count or types in advance
and it can never appear in either - that registry only enumerates real,
registered Add-Ins. (An earlier `.oxt`-wrapped-Basic attempt confirmed this
directly: the exact same code resolved fine once moved into a genuine Add-In,
and never appeared in the Function Wizard as a Basic macro, no matter how it
was packaged.)

Building the Add-In requires the LibreOffice **SDK** (for `unoidl-write`,
which compiles the IDL interface in `idl/com/example/loess/XLoess.idl` into a
UNO type library - no C++/Java compiler needed, just the SDK tool) and
Python. Check `<LibreOffice install dir>/sdk/bin/unoidl-write` exists; on this
machine that's `/usr/lib64/libreoffice/sdk/bin/unoidl-write`.

```sh
./build_addin.sh                                    # -> build/CalcLoessAddin.oxt
unopkg add --force build/CalcLoessAddin.oxt         # install
```

Restart LibreOffice, then `=LOESS(...)` autocompletes and shows full argument
help in the Function Wizard (category "Add-In").

Remove with `unopkg remove com.example.loess`.

## Usage

```
=LOESS(XRange; YRange; X0; [Span]; [Degree]; [RobustIters])
```

| Argument      | Meaning                                                                                                                                                                                    |
|---------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `XRange`      | Range (or single cell) of sample x-values.                                                                                                                                              |
| `YRange`      | Range (or single cell) of sample y-values, same shape as `XRange`.                                                                                                                      |
| `X0`          | The x-value at which to evaluate the smoothed curve. Drag the formula across/down to build a smoothed series, the same way `TREND()` is used.                                          |
| `Span`        | Fraction of the data used in each local fit. `0 < Span <= 1` selects the nearest `Span*N` points (a k-nearest-neighbour window). `Span > 1` uses all points with a widened bandwidth. Default `0.6667` (2/3, the classic LOWESS default). |
| `Degree`      | Degree of the local polynomial: `0` (local weighted mean), `1` (local linear, default), or `2` (local quadratic).                                                                       |
| `RobustIters` | Number of Cleveland-style robustness (bisquare) iterations that down-weight outliers. Default `0` (fast, single pass). Classic LOWESS uses `3`. See **Performance** below.              |

Example - smoothing a noisy series in column A (x) / B (y), evaluated at the
x-value in D2, dragged down:

```
=LOESS($A$1:$A$100; $B$1:$B$100; D2; 0.3; 1; 0)
```

Blank or non-numeric cells are handled safely: a pair is only used if *both*
the x and y cell at that position are numeric, so a gap on one side can never
shift the two ranges out of alignment.

See [`examples/LOESS_Demo.ods`](examples/LOESS_Demo.ods) for a worked example:
sample noisy data with a deliberate outlier, `LOESS()` dragged down two
columns (degree 1 and 2), a robust-vs-non-robust comparison at the outlier,
and a chart.

![Example workbook: LOESS-smoothed columns next to a chart comparing the raw noisy data against the degree-1 and degree-2 fits, including the spike from the deliberate outlier](docs/LOESS_Demo_screenshot.png)

## Performance

`RobustIters > 0` is O(n²) per formula evaluation, because deriving the
robustness weights requires fitting a local regression at *every* sample
point, not just at `X0`. If you drag a `RobustIters > 0` formula down a
column, each cell redundantly redoes this O(n²) work (Calc has no way to
share it across cells). For a few dozen to a couple hundred points this is
unnoticeable; for large datasets dragged across many cells, prefer
`RobustIters = 0`, or precompute the smoothed series once and paste it as
values.

## Algorithm

- **Neighbourhood weights**: tricube kernel `(1 - u^3)^3` on the distance to
  `X0`, scaled by a bandwidth taken from the `Span*N`-th nearest neighbour
  (or the full data range, scaled by `Span`, when `Span > 1`).
- **Local fit**: the weighted polynomial (degree 0/1/2) is fit in coordinates
  centred on `X0`, so its intercept *is* the smoothed value - solved via
  Gaussian elimination with partial pivoting on the small normal-equations
  system.
- **Robustness**: when `RobustIters > 0`, residuals from an initial fit at
  every sample point feed a bisquare re-weighting (scaled by 6x the median
  absolute residual), repeated `RobustIters` times, following Cleveland
  (1979), *"Robust Locally Weighted Regression and Smoothing Scatterplots"*,
  JASA 74(368): 829-836.

## Repository layout

```
idl/com/example/loess/XLoess.idl
                    UNO interface for the Add-In function.
src/loess_impl.py   Python implementation of the LOESS() function.
registration/       CalcAddIns.xcu / manifest.xml / description.xml for the Add-In package.
build_addin.sh      Compiles the IDL and packages build/CalcLoessAddin.oxt.
tools/test_addin.py End-to-end test of the Add-In against a headless LibreOffice.
examples/LOESS_Demo.ods
                    Sample workbook - see Usage above.
```

## Changelog

Release history is in [CHANGELOG.md](CHANGELOG.md).

## Testing notes

The algorithm was cross-checked against an independent NumPy reference
implementation of the same tricube/robustness algorithm across a range of
`Span`/`Degree`/`RobustIters` combinations, boundary and extrapolation
points, exact linear/quadratic data (degree-matched fits reproduce the data
exactly), misaligned blank cells, single-cell inputs, and invalid-input error
handling - all matching to at least 8 significant digits where an analytic or
independently-computed answer was available.
