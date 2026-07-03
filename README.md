# Calc LOESS/LOWESS Add-in

Objective is to create addin for Calc calculate loess or lowess without using python

A LOESS/LOWESS (locally weighted polynomial regression) smoothing function for
LibreOffice Calc, implemented entirely in LibreOffice Basic (StarBasic). No
Python, no external libraries, no compiler - just a Basic module you can
import directly, or an extension (`.oxt`) you can install through the
Extension Manager.

## What it does

`LOESS()` fits a local polynomial (constant, linear, or quadratic) around each
requested x-value, weighted by a tricube kernel over the nearest neighbours,
optionally refined with Cleveland-style robustness iterations that down-weight
outliers. This is the same family of algorithm as R's `lowess()`/`loess()`.

## Installation

### Option A: Install the extension (recommended)

1. Build (or download) `CalcLoessAddin.oxt`. To build it yourself from
   source: `./build_oxt.sh` (regenerates the `.oxt` from `src/*.bas`).
2. In LibreOffice, go to **Tools > Extension Manager > Add...** and select
   `CalcLoessAddin.oxt`.
3. Restart LibreOffice. `LOESS()` is now available in every Calc document,
   the same way `SUM()` or `TREND()` are.

### Option B: Import the Basic module manually

1. Open **Tools > Macros > Edit Macros...** to open the Basic IDE.
2. Under **My Macros** (or a shared library), insert a new module and paste
   in the contents of `src/LOESS.bas`.
3. `LOESS()` is now usable as a formula in that LibreOffice profile.

Either way, once installed, run `Tools > Macros > Run Macro...` and execute
`RunSelfTest` (from `SelfTest.bas` / the extension's second module) to sanity
check the install - it opens a scratch spreadsheet, runs a few checks with
analytically known answers, and reports PASS/FAIL in a dialog.

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
src/LOESS.bas       The add-in itself: the LOESS() function and its helpers.
src/SelfTest.bas    Interactive self-test (Tools > Macros > Run Macro > RunSelfTest).
oxt/                Generated extension package sources (do not edit directly - see build_oxt.sh).
build_oxt.sh        Regenerates oxt/ and CalcLoessAddin.oxt from src/*.bas.
CalcLoessAddin.oxt  Ready-to-install extension.
```

## Testing notes

The algorithm was cross-checked against an independent NumPy reference
implementation of the same tricube/robustness algorithm across a range of
`Span`/`Degree`/`RobustIters` combinations, boundary and extrapolation
points, exact linear/quadratic data (degree-matched fits reproduce the data
exactly), misaligned blank cells, single-cell inputs, and invalid-input error
handling - all matching to at least 8 significant digits where an analytic or
independently-computed answer was available.
