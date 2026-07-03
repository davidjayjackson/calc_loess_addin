# Calc LOESS/LOWESS Add-in

Objective is to create addin for Calc calculate loess or lowess without using python

A LOESS/LOWESS (locally weighted polynomial regression) smoothing function for
LibreOffice Calc, implemented entirely in LibreOffice Basic (StarBasic). No
Python, no external libraries, no compiler - just a Basic module.

## What it does

`LOESS()` fits a local polynomial (constant, linear, or quadratic) around each
requested x-value, weighted by a tricube kernel over the nearest neighbours,
optionally refined with Cleveland-style robustness iterations that down-weight
outliers. This is the same family of algorithm as R's `lowess()`/`loess()`.

## Installation

Run `./install.sh`. It copies `src/LOESS.bas` and `src/SelfTest.bas` into your
personal LibreOffice **Standard** Basic library
(`~/.config/libreoffice/4/user/basic/Standard`) as two new modules
(`LOESSAddin`, `LOESSSelfTest`), backing up your existing Standard library
first. LibreOffice must be fully closed while it runs. Re-run it any time to
pick up changes to `src/LOESS.bas` - it's additive and idempotent, so it never
touches your own existing macros.

Why not a `.oxt` extension installed through the Extension Manager? That was
the first approach here, but it turned out to be a dead end worth recording:
**Calc's formula compiler only resolves a bare `=FUNCTION(...)` name against
your personal "Standard" library.** A Basic library shipped inside an
extension - regardless of what you name it, even "Standard" itself - gets
installed, shows up in the Extension Manager, and can be run as a macro, but
is never searched when compiling a cell formula, so `=LOESS(...)` fails with
`#NAME?` no matter what. This was verified directly (not assumed): the exact
same code resolves correctly the moment it's placed in the personal Standard
library, and fails every time it's shipped via `.oxt`, independent of the
library's name. A "real" `=FUNCTION(...)` UNO Add-In needs a compiled/scripted
component (Python, Java, C++) implementing `com.sun.star.sheet.AddIn`, which
is off the table given the "no Python" objective - so a Standard-library
install is the correct solution here, not a workaround.

Once installed, run **Tools > Macros > Run Macro... > My Macros > Standard >
LOESSSelfTest > RunSelfTest** to sanity check it - it opens a scratch
spreadsheet, runs a few checks with analytically known answers, and reports
PASS/FAIL in a dialog.

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
install.sh          Installs src/*.bas into your personal Standard library. Run this.
oxt/, build_oxt.sh, CalcLoessAddin.oxt
                    A .oxt extension package, kept for reference. It installs
                    and is runnable as a macro, but =LOESS(...) as a cell
                    formula will NOT work from it - see Installation above.
```

## Testing notes

The algorithm was cross-checked against an independent NumPy reference
implementation of the same tricube/robustness algorithm across a range of
`Span`/`Degree`/`RobustIters` combinations, boundary and extrapolation
points, exact linear/quadratic data (degree-matched fits reproduce the data
exactly), misaligned blank cells, single-cell inputs, and invalid-input error
handling - all matching to at least 8 significant digits where an analytic or
independently-computed answer was available.
