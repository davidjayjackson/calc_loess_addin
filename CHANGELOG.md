# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.1] - 2026-07-03

### Documentation
- Populated `CLAUDE.md` (previously an empty placeholder) with build/test
  commands and the IDL-to-Python-AddIn architecture notes, for future work
  in this repo. No functional changes.

## [2.0.0] - 2026-07-03

### Added
- **Real UNO Add-In implementation** (`com.example.loess`, `src/loess_impl.py`,
  `idl/com/example/loess/XLoess.idl`, `registration/`, `build_addin.sh`).
  Because `LOESS` is now a genuine `com.sun.star.sheet.AddIn` rather than a
  Basic macro merely callable from a formula, it appears in the Function
  Wizard (category "Add-In") and autocompletes as you type - verified
  directly against `com.sun.star.sheet.FunctionDescriptions`.
- `tools/test_addin.py`, an end-to-end test of the Add-In against a headless
  LibreOffice instance.

### Removed
- **Breaking:** the Basic-macro install path (`install.sh`, `src/LOESS.bas`,
  `src/SelfTest.bas`). It worked as a cell formula but could never appear in
  the Function Wizard or autocomplete - a hard architectural limit of Basic
  macros, not a bug. The project is now Python-Add-In-only.
- The original `.oxt`-wrapped-Basic attempt (`oxt/`, `build_oxt.sh`, the
  committed `CalcLoessAddin.oxt`), a dead end that installed but never
  resolved `=LOESS(...)` as a cell formula. Superseded by the real Add-In
  above; the explanation of why it didn't work is kept in the README.

## [1.0.0] - 2026-07-03

First release of **LOESS**, a LibreOffice Calc add-in for LOESS/LOWESS local
polynomial regression smoothing, implemented in pure StarBasic.

### Added
- `LOESS(xrange; yrange; x0; [span]; [degree]; [robust_iters])` worksheet
  function: tricube-weighted local polynomial fit (degree 0/1/2) with
  optional Cleveland-style robustness iterations.
- `install.sh`, installing the function into the personal Standard Basic
  library (no Python, no compiler).
- Interactive self-test (`src/SelfTest.bas`, `Tools > Macros > Run Macro >
  RunSelfTest`).
- `examples/LOESS_Demo.ods`, a worked example workbook with a chart.
- README documenting usage, algorithm, and performance characteristics.

[Unreleased]: https://github.com/davidjayjackson/calc_loess_addin/compare/v2.0.1...HEAD
[2.0.1]: https://github.com/davidjayjackson/calc_loess_addin/compare/v2.0.0...v2.0.1
[2.0.0]: https://github.com/davidjayjackson/calc_loess_addin/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/davidjayjackson/calc_loess_addin/releases/tag/v1.0.0
