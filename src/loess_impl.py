# LOESS / LOWESS local polynomial regression smoothing, exposed as a
# Calc spreadsheet function via a real UNO Add-In (so it appears in the
# Function Wizard and autocomplete). See README.md for the full algorithm
# description.
#
# LOESS(xrange; yrange; x0; [span]; [degree]; [robust_iters])
#
# Reference: Cleveland, W.S. (1979). "Robust Locally Weighted Regression
# and Smoothing Scatterplots." JASA 74(368): 829-836.

import unohelper

from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XServiceName
from com.example.loess import XLoess

ADDIN_SERVICE = "com.sun.star.sheet.AddIn"
SERVICE_NAME = "com.example.loess.LoessImpl"
IMPL_NAME = "com.example.loess.LoessImpl.python"


class LoessError(ValueError):
    """Raised for invalid LOESS arguments/data."""


def _is_numeric(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _build_xy(xdata, ydata):
    """Flattens xdata/ydata into matched (x, y) pairs, keeping the two
    ranges in lock-step so a blank/text cell on one side does not shift
    the other side's values out of alignment."""
    rows = len(xdata)
    cols = len(xdata[0]) if rows else 0
    if len(ydata) != rows or any(len(r) != cols for r in ydata):
        raise LoessError("xdata and ydata must have the same shape")

    x, y = [], []
    for i in range(rows):
        for j in range(cols):
            xv, yv = xdata[i][j], ydata[i][j]
            if _is_numeric(xv) and _is_numeric(yv):
                x.append(float(xv))
                y.append(float(yv))

    if not x:
        raise LoessError("no numeric (x, y) pairs found")
    return x, y


def _tricube_weight(u):
    u = abs(u)
    if u >= 1:
        return 0.0
    return (1 - u ** 3) ** 3


def _bisquare_weight(u):
    u = abs(u)
    if u >= 1:
        return 0.0
    return (1 - u ** 2) ** 2


def _median_of(values):
    s = sorted(values)
    n = len(s)
    if n % 2 == 1:
        return s[n // 2]
    return (s[n // 2 - 1] + s[n // 2]) / 2


def _compute_bandwidth(dist, span):
    sorted_dist = sorted(dist)
    n = len(sorted_dist)
    if span >= 1:
        return sorted_dist[-1] * span
    q = int(span * n + 0.0000001)
    q = max(q, 1)
    q = min(q, n)
    h = sorted_dist[q - 1]
    while h <= 0 and q < n:
        q += 1
        h = sorted_dist[q - 1]
    return h


def _solve_linear_system(a, b, m):
    """Gaussian elimination with partial pivoting for the small (<=3x3)
    normal-equations system produced by _weighted_poly_fit."""
    a = [row[:] for row in a]
    b = b[:]

    for k in range(m):
        max_row, max_val = k, abs(a[k][k])
        for i in range(k + 1, m):
            if abs(a[i][k]) > max_val:
                max_val, max_row = abs(a[i][k]), i
        if max_row != k:
            a[k], a[max_row] = a[max_row], a[k]
            b[k], b[max_row] = b[max_row], b[k]

        if abs(a[k][k]) < 0.0000000001:
            a[k][k] = 0.0000000001  # guard: near-singular (e.g. duplicate x values)

        for i in range(k + 1, m):
            factor = a[i][k] / a[k][k]
            for j in range(k, m):
                a[i][j] -= factor * a[k][j]
            b[i] -= factor * b[k]

    beta = [0.0] * m
    for i in range(m - 1, -1, -1):
        tmp = b[i]
        for j in range(i + 1, m):
            tmp -= a[i][j] * beta[j]
        beta[i] = tmp / a[i][i]
    return beta


def _weighted_poly_fit(x, y, x0, degree, w):
    """Weighted least-squares fit of a degree-th order polynomial, centred
    on x0 so that the intercept of the fit IS the smoothed value at x0."""
    m = degree + 1
    a = [[0.0] * m for _ in range(m)]
    b = [0.0] * m

    for xi, yi, wi in zip(x, y, w):
        if wi <= 0:
            continue
        u = xi - x0
        upow = [1.0] * (2 * degree + 1)
        for p in range(1, 2 * degree + 1):
            upow[p] = upow[p - 1] * u
        for p in range(m):
            for q in range(m):
                a[p][q] += wi * upow[p + q]
            b[p] += wi * upow[p] * yi

    beta = _solve_linear_system(a, b, m)
    return beta[0]


def _local_fit(x, y, x0, span, degree, robust_w):
    """Builds the tricube kernel weights around x0, combines them with
    the supplied robustness weights, and solves the weighted
    least-squares polynomial."""
    dist = [abs(xi - x0) for xi in x]
    h = _compute_bandwidth(dist, span)

    if h > 0:
        w = [_tricube_weight(d / h) * rw for d, rw in zip(dist, robust_w)]
    else:
        w = list(robust_w)  # all selected points coincide with x0

    return _weighted_poly_fit(x, y, x0, degree, w)


def _compute_robust_weights(x, y, span, degree, iters):
    """Cleveland-style robustness iterations: fit at every data point,
    derive a scale estimate from the median absolute residual, and
    down-weight points with large residuals via the bisquare function."""
    n = len(x)
    r = [1.0] * n

    for _ in range(iters):
        fitted = [_local_fit(x, y, xi, span, degree, r) for xi in x]
        resid = [abs(yi - fi) for yi, fi in zip(y, fitted)]

        s = _median_of(resid)
        if s <= 0.0000000001:
            # Degenerate case: at least half the points fit exactly, so any
            # point with a nonzero residual is by definition an outlier -
            # give it zero weight rather than dividing by a zero scale.
            r = [1.0 if ri <= 0.0000000001 else 0.0 for ri in resid]
        else:
            r = [_bisquare_weight(ri / (6 * s)) for ri in resid]

    return r


def _loess(xdata, ydata, x0, span, degree, robust_iters):
    if span <= 0:
        raise LoessError("span must be positive")
    if degree not in (0, 1, 2):
        raise LoessError("degree must be 0, 1 or 2")
    if robust_iters < 0:
        raise LoessError("robust_iters must not be negative")

    x, y = _build_xy(xdata, ydata)
    if len(x) < degree + 1:
        raise LoessError("not enough data points for this degree")

    if robust_iters > 0:
        robust_w = _compute_robust_weights(x, y, span, degree, robust_iters)
    else:
        robust_w = [1.0] * len(x)

    return _local_fit(x, y, x0, span, degree, robust_w)


def _as_number(value, default):
    """Coerce an optional 'any' argument (None when omitted) to a number."""
    if value is None:
        return default
    return value


class LoessAddIn(unohelper.Base, XLoess, XAddIn, XServiceName):
    """Implementation of the LOESS spreadsheet function."""

    def __init__(self, ctx):
        self.ctx = ctx
        self._locale = None

    # --- XLoess -------------------------------------------------------
    def loess(self, xdata, ydata, x0, span, degree, robustIters):
        dspan = float(_as_number(span, 0.6667))
        ldegree = int(_as_number(degree, 1))
        literers = int(_as_number(robustIters, 0))
        return _loess(xdata, ydata, x0, dspan, ldegree, literers)

    # --- XAddIn ---------------------------------------------------------
    # Function/argument metadata is supplied by CalcAddIns.xcu, so these
    # return the programmatic names (or empty strings) as a safe fallback.
    def getProgrammaticFuntionName(self, aDisplayName):  # UNO API spelling
        return "loess" if aDisplayName == "LOESS" else ""

    def getDisplayFunctionName(self, aProgrammaticName):
        return "LOESS" if aProgrammaticName == "loess" else ""

    def getFunctionDescription(self, aProgrammaticName):
        if aProgrammaticName == "loess":
            return "Fits a local polynomial (LOESS/LOWESS) smoother and evaluates it at x0."
        return ""

    def getDisplayArgumentName(self, aProgrammaticName, nArgument):
        if aProgrammaticName == "loess":
            names = ("xrange", "yrange", "x0", "span", "degree", "robust_iters")
            return names[nArgument] if 0 <= nArgument < len(names) else ""
        return ""

    def getArgumentDescription(self, aProgrammaticName, nArgument):
        if aProgrammaticName == "loess":
            descs = (
                "Range (or single cell) of sample x-values.",
                "Range (or single cell) of sample y-values, same shape as xrange.",
                "The x-value at which to evaluate the smoothed curve.",
                "Optional. Fraction of data used per local fit (0 < span <= 1 nearest-neighbour, "
                "span > 1 widens the bandwidth). Default 0.6667.",
                "Optional. Degree of the local polynomial: 0, 1 (default) or 2.",
                "Optional. Number of Cleveland-style robustness iterations. Default 0.",
            )
            return descs[nArgument] if 0 <= nArgument < len(descs) else ""
        return ""

    def getProgrammaticCategoryName(self, aProgrammaticName):
        return "Add-In"

    def getDisplayCategoryName(self, aProgrammaticName):
        return "Add-In"

    # --- XLocalizable (base of XAddIn) ----------------------------------
    def setLocale(self, aLocale):
        self._locale = aLocale

    def getLocale(self):
        return self._locale

    # --- XServiceName -----------------------------------------------------
    def getServiceName(self):
        return SERVICE_NAME


# --- component registration ----------------------------------------------
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    LoessAddIn,
    IMPL_NAME,
    (SERVICE_NAME, ADDIN_SERVICE),
)
