# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A LOESS/LOWESS (locally weighted polynomial regression) smoothing function,
`LOESS()`, for LibreOffice Calc. It's implemented as a real UNO Add-In
(`com.sun.star.sheet.AddIn`) in Python (`src/loess_impl.py`), not a Basic
macro — that distinction is load-bearing: a Basic macro can be called from a
formula but never appears in the Function Wizard or formula autocomplete,
because those are driven by a separate registry
(`com.sun.star.sheet.FunctionDescriptions`) that only enumerates real,
registered Add-Ins. Getting `LOESS` into that registry is the entire reason
this is a compiled/scripted Add-In instead of a plain macro.

## Commands

Build the `.oxt` extension package (requires the LibreOffice **SDK**, for
`unoidl-write`; no C++/Java compiler needed — it's a pure IDL-to-type-library
compile step):

```sh
./build_addin.sh                              # -> build/CalcLoessAddin.oxt
```

`LIBREOFFICE` env var overrides the install dir (default
`/usr/lib64/libreoffice`; must contain `sdk/bin/unoidl-write` and
`program/types.rdb`).

Install/reinstall (LibreOffice must be restarted after):

```sh
unopkg add --force build/CalcLoessAddin.oxt
```

Remove:

```sh
unopkg remove com.example.loess
```

Run the end-to-end test against a headless LibreOffice instance (requires
the extension to already be installed):

```sh
soffice --headless --invisible --norestore --accept="socket,host=localhost,port=2002;urp;" &
/usr/lib64/libreoffice/program/python tools/test_addin.py
```

There is no unit test suite — `tools/test_addin.py` is the only test, and it
drives a real Calc instance over UNO rather than testing in isolation. It
checks both that `LOESS` is registered in `FunctionDescriptions` (proves
Function Wizard/autocomplete wiring) and that a formula evaluates to the
correct numeric result, with and without the optional trailing arguments.

## Architecture

**IDL interface (`idl/com/example/loess/XLoess.idl`) defines the contract
Calc calls into.** `unoidl-write` compiles this into `types/XLoess.rdb`
inside the `.oxt`. The directory path `idl/com/example/loess/` must match
the `module com { module example { module loess { ... } } }` declaration —
`unoidl-write`'s source-tree reader derives the UNO module name from the
directory structure, not from the file's own `module {}` block, so moving
the `.idl` file without moving its parent directories breaks the build.

**`src/loess_impl.py` implements that interface** as a `unohelper.Base`
subclass registered as both `com.example.loess.LoessImpl` and the generic
`com.sun.star.sheet.AddIn` service. Required numeric ranges (`xdata`,
`ydata`) are typed `sequence<sequence<any>>` in the IDL so blank/text cells
can be skipped without erroring (a pair is only used if both the x and y
cell at that position are numeric — this keeps the two ranges in lock-step
so a gap on one side never shifts the other out of alignment); the three
optional trailing arguments (`span`, `degree`, `robustIters`) are typed
`any` and arrive as Python `None` when omitted from the formula.

**`registration/CalcAddIns.xcu` is what actually populates the Function
Wizard** — display names, descriptions, category ("Add-In"), and per-argument
help text. The `XAddIn` methods on `LoessAddIn` (`getFunctionDescription`,
`getDisplayArgumentName`, etc.) are a fallback path and largely redundant
with the `.xcu` in practice.

**`build_addin.sh` assembles the `.oxt`** from these pieces:
`idl/` → compiled `.rdb` (via `unoidl-write`), `src/loess_impl.py`,
`registration/*.xcu`/`manifest.xml`/`description.xml` — staged into
`build/oxt/` and zipped. `registration/manifest.xml` maps each staged file
to its final path inside the package (`types/`, `python/`, `config/`).

**The smoothing algorithm itself** (tricube-weighted local polynomial fit,
degree 0/1/2, optional Cleveland-style bisquare robustness iterations) lives
entirely in `loess_impl.py`'s private helper functions
(`_build_xy`, `_local_fit`, `_compute_bandwidth`, `_weighted_poly_fit`,
`_solve_linear_system`, `_compute_robust_weights`) — see README.md's
"Algorithm" section for the math, and its "Performance" section for why
`robust_iters > 0` is O(n²) per cell and shouldn't be dragged across a large
range.

## Versioning

`registration/description.xml`'s `<version>` and `CHANGELOG.md` must be
bumped together on every release, followed by a matching git tag (`vX.Y.Z`)
and a GitHub release with `build/CalcLoessAddin.oxt` attached as the
downloadable asset (see the README's Installation section, which links
`/releases/latest`).
