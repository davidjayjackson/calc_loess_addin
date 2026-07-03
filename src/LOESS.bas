Option Explicit

'=====================================================================
' LOESS.bas
'
' LOESS / LOWESS local polynomial regression smoothing, exposed as a
' Calc spreadsheet function. Pure StarBasic - no Python, no external
' libraries, no compiled components.
'
' Usage in a cell:
'   =LOESS(XRange; YRange; X0; [Span]; [Degree]; [RobustIters])
'
'   XRange, YRange : ranges (or single cells) holding the sample data.
'                    Must have the same shape. Non-numeric / blank
'                    cells are skipped, but only when BOTH the X and Y
'                    cell at a given position are blank/non-numeric -
'                    if only one side is blank the pair is dropped so
'                    the two ranges never fall out of alignment.
'   X0             : the x-value at which the smoothed y-value is
'                    wanted. Drag the formula across/down to build a
'                    smoothed curve, the same way TREND() is used.
'   Span           : fraction of the data used in each local fit.
'                    0 < Span <= 1 selects the Span*N nearest points
'                    to X0 (a k-nearest-neighbour window, as in
'                    R's loess()). Span > 1 uses all points and widens
'                    the bandwidth by that factor, which is useful to
'                    avoid noisy fits at the ends of the data.
'                    Default 0.6667 (2/3, the classic LOWESS default).
'   Degree         : degree of the local polynomial regression:
'                    0 = local weighted mean, 1 = local linear
'                    (default), 2 = local quadratic.
'   RobustIters    : number of Cleveland-style robustness iterations
'                    (bisquare re-weighting of outliers). Default 0
'                    (fast, single pass). Classic LOWESS uses 3, at
'                    the cost of roughly N times more computation per
'                    call, since residuals across all points are
'                    needed to derive the robustness weights.
'
' Reference: Cleveland, W.S. (1979). "Robust Locally Weighted
' Regression and Smoothing Scatterplots." JASA 74(368): 829-836.
'=====================================================================

Function LOESS(ByVal XRange As Variant, ByVal YRange As Variant, ByVal X0 As Double, _
                Optional ByVal Span As Variant, Optional ByVal Degree As Variant, _
                Optional ByVal RobustIters As Variant) As Double

    Dim dSpan As Double, lDegree As Long, lIters As Long
    Dim x() As Double, y() As Double
    Dim n As Long
    Dim i As Long
    Dim robustW() As Double

    If IsMissing(Span) Then dSpan = 0.6667 Else dSpan = CDbl(Span)
    If IsMissing(Degree) Then lDegree = 1 Else lDegree = CLng(Degree)
    If IsMissing(RobustIters) Then lIters = 0 Else lIters = CLng(RobustIters)

    If dSpan <= 0 Then Error 5          ' Span must be positive
    If lDegree < 0 Or lDegree > 2 Then Error 5   ' Degree must be 0, 1 or 2
    If lIters < 0 Then Error 5          ' RobustIters must not be negative

    Call BuildXY(XRange, YRange, x, y, n)

    If n < lDegree + 1 Then Error 5     ' not enough data points for this degree

    ReDim robustW(1 To n)
    If lIters > 0 Then
        Call ComputeRobustWeights(x, y, n, dSpan, lDegree, lIters, robustW)
    Else
        For i = 1 To n
            robustW(i) = 1
        Next i
    End If

    LOESS = LocalFit(x, y, n, X0, dSpan, lDegree, robustW)

End Function

'---------------------------------------------------------------------
' Flattens XRange/YRange into matched (x,y) pairs, keeping the two
' ranges in lock-step so a blank/text cell on one side does not shift
' the other side's values out of alignment. A lone cell reference
' arrives as a plain scalar rather than an array, so both cases are
' handled.
'---------------------------------------------------------------------
Private Sub BuildXY(ByVal xr As Variant, ByVal yr As Variant, xOut() As Double, yOut() As Double, ByRef n As Long)

    Dim xIsArr As Boolean, yIsArr As Boolean
    Dim xr1 As Long, xr2 As Long, xc1 As Long, xc2 As Long
    Dim yr1 As Long, yr2 As Long, yc1 As Long, yc2 As Long
    Dim rows As Long, cols As Long
    Dim i As Long, j As Long, cnt As Long
    Dim xv As Variant, yv As Variant

    xIsArr = IsArray(xr)
    yIsArr = IsArray(yr)

    If xIsArr Then
        xr1 = LBound(xr, 1) : xr2 = UBound(xr, 1)
        xc1 = LBound(xr, 2) : xc2 = UBound(xr, 2)
    Else
        xr1 = 1 : xr2 = 1 : xc1 = 1 : xc2 = 1
    End If

    If yIsArr Then
        yr1 = LBound(yr, 1) : yr2 = UBound(yr, 1)
        yc1 = LBound(yr, 2) : yc2 = UBound(yr, 2)
    Else
        yr1 = 1 : yr2 = 1 : yc1 = 1 : yc2 = 1
    End If

    rows = xr2 - xr1 + 1
    cols = xc2 - xc1 + 1

    If (yr2 - yr1 + 1) <> rows Or (yc2 - yc1 + 1) <> cols Then
        Error 5   ' XRange and YRange must have the same shape
    End If

    ReDim xOut(1 To rows * cols)
    ReDim yOut(1 To rows * cols)
    cnt = 0

    For i = 0 To rows - 1
        For j = 0 To cols - 1
            If xIsArr Then xv = xr(xr1 + i, xc1 + j) Else xv = xr
            If yIsArr Then yv = yr(yr1 + i, yc1 + j) Else yv = yr
            If IsNumeric(xv) And IsNumeric(yv) Then
                cnt = cnt + 1
                xOut(cnt) = CDbl(xv)
                yOut(cnt) = CDbl(yv)
            End If
        Next j
    Next i

    If cnt = 0 Then Error 5   ' no numeric (x,y) pairs found

    If cnt < rows * cols Then
        ReDim Preserve xOut(1 To cnt)
        ReDim Preserve yOut(1 To cnt)
    End If

    n = cnt

End Sub

'---------------------------------------------------------------------
' Local polynomial fit evaluated at a single point x0: builds the
' tricube kernel weights around x0, combines them with the supplied
' robustness weights, and solves the weighted least-squares polynomial.
'---------------------------------------------------------------------
Private Function LocalFit(x() As Double, y() As Double, ByVal n As Long, ByVal x0 As Double, _
                           ByVal span As Double, ByVal degree As Long, robustW() As Double) As Double

    Dim dist() As Double
    Dim w() As Double
    Dim h As Double
    Dim i As Long

    ReDim dist(1 To n)
    For i = 1 To n
        dist(i) = Abs(x(i) - x0)
    Next i

    h = ComputeBandwidth(dist, n, span)

    ReDim w(1 To n)
    For i = 1 To n
        If h > 0 Then
            w(i) = TricubeWeight(dist(i) / h) * robustW(i)
        Else
            w(i) = robustW(i)   ' all selected points coincide with x0
        End If
    Next i

    LocalFit = WeightedPolyFit(x, y, n, x0, degree, w)

End Function

'---------------------------------------------------------------------
' Bandwidth = distance to the Span*N-th nearest neighbour of x0
' (or, for Span > 1, the maximum distance scaled by Span).
'---------------------------------------------------------------------
Private Function ComputeBandwidth(dist() As Double, ByVal n As Long, ByVal span As Double) As Double

    Dim sorted() As Double
    Dim i As Long, q As Long

    ReDim sorted(1 To n)
    For i = 1 To n
        sorted(i) = dist(i)
    Next i
    Call QuickSort(sorted, 1, n)

    If span >= 1 Then
        ComputeBandwidth = sorted(n) * span
    Else
        q = Int(span * n + 0.0000001)
        If q < 1 Then q = 1
        If q > n Then q = n
        ComputeBandwidth = sorted(q)
        Do While ComputeBandwidth <= 0 And q < n
            q = q + 1
            ComputeBandwidth = sorted(q)
        Loop
    End If

End Function

Private Function TricubeWeight(ByVal u As Double) As Double
    If u < 0 Then u = -u
    If u >= 1 Then
        TricubeWeight = 0
    Else
        TricubeWeight = (1 - u * u * u) ^ 3
    End If
End Function

Private Function BisquareWeight(ByVal u As Double) As Double
    If u < 0 Then u = -u
    If u >= 1 Then
        BisquareWeight = 0
    Else
        BisquareWeight = (1 - u * u) ^ 2
    End If
End Function

'---------------------------------------------------------------------
' Cleveland-style robustness iterations: fit at every data point,
' derive a scale estimate from the median absolute residual, and
' down-weight points with large residuals via the bisquare function.
' O(n^2) per iteration, since a full local fit is needed at every
' point to get residuals - fine for the modest data sizes typical of
' a spreadsheet, but avoid combining a large RobustIters with a very
' large sample dragged across many formula cells.
'---------------------------------------------------------------------
Private Sub ComputeRobustWeights(x() As Double, y() As Double, ByVal n As Long, ByVal span As Double, _
                                  ByVal degree As Long, ByVal iters As Long, r() As Double)

    Dim i As Long, it As Long
    Dim fitted() As Double, resid() As Double
    Dim s As Double

    For i = 1 To n
        r(i) = 1
    Next i

    ReDim fitted(1 To n)
    ReDim resid(1 To n)

    For it = 1 To iters
        For i = 1 To n
            fitted(i) = LocalFit(x, y, n, x(i), span, degree, r)
        Next i
        For i = 1 To n
            resid(i) = Abs(y(i) - fitted(i))
        Next i

        s = MedianOf(resid, n)
        If s <= 0.0000000001 Then
            ' Degenerate case: at least half the points fit exactly, so any
            ' point with a nonzero residual is by definition an outlier -
            ' give it zero weight rather than dividing by a zero scale.
            For i = 1 To n
                If resid(i) <= 0.0000000001 Then
                    r(i) = 1
                Else
                    r(i) = 0
                End If
            Next i
        Else
            For i = 1 To n
                r(i) = BisquareWeight(resid(i) / (6 * s))
            Next i
        End If
    Next it

End Sub

Private Function MedianOf(arr() As Double, ByVal n As Long) As Double

    Dim tmp() As Double
    Dim i As Long

    ReDim tmp(1 To n)
    For i = 1 To n
        tmp(i) = arr(i)
    Next i
    Call QuickSort(tmp, 1, n)

    If n Mod 2 = 1 Then
        MedianOf = tmp((n + 1) \ 2)
    Else
        MedianOf = (tmp(n \ 2) + tmp(n \ 2 + 1)) / 2
    End If

End Function

'---------------------------------------------------------------------
' Weighted least-squares fit of a degree-th order polynomial, centred
' on x0 so that the intercept of the fit IS the smoothed value at x0.
'---------------------------------------------------------------------
Private Function WeightedPolyFit(x() As Double, y() As Double, ByVal n As Long, ByVal x0 As Double, _
                                  ByVal degree As Long, w() As Double) As Double

    Dim m As Long
    Dim A(0 To 2, 0 To 2) As Double
    Dim b(0 To 2) As Double
    Dim beta(0 To 2) As Double
    Dim upow(0 To 4) As Double
    Dim i As Long, p As Long, q As Long
    Dim u As Double, wi As Double

    m = degree + 1

    For i = 1 To n
        wi = w(i)
        If wi > 0 Then
            u = x(i) - x0
            upow(0) = 1
            For p = 1 To 2 * degree
                upow(p) = upow(p - 1) * u
            Next p
            For p = 0 To m - 1
                For q = 0 To m - 1
                    A(p, q) = A(p, q) + wi * upow(p + q)
                Next q
                b(p) = b(p) + wi * upow(p) * y(i)
            Next p
        End If
    Next i

    Call SolveLinearSystem(A, b, m, beta)

    WeightedPolyFit = beta(0)

End Function

'---------------------------------------------------------------------
' Gaussian elimination with partial pivoting for the small (<=3x3)
' normal-equations system produced by WeightedPolyFit.
'---------------------------------------------------------------------
Private Sub SolveLinearSystem(A() As Double, b() As Double, ByVal m As Long, beta() As Double)

    Dim i As Long, j As Long, k As Long
    Dim maxRow As Long, maxVal As Double, tmp As Double, factor As Double

    For k = 0 To m - 1
        maxRow = k
        maxVal = Abs(A(k, k))
        For i = k + 1 To m - 1
            If Abs(A(i, k)) > maxVal Then
                maxVal = Abs(A(i, k))
                maxRow = i
            End If
        Next i
        If maxRow <> k Then
            For j = 0 To m - 1
                tmp = A(k, j) : A(k, j) = A(maxRow, j) : A(maxRow, j) = tmp
            Next j
            tmp = b(k) : b(k) = b(maxRow) : b(maxRow) = tmp
        End If

        If Abs(A(k, k)) < 0.0000000001 Then A(k, k) = 0.0000000001   ' guard: near-singular (e.g. duplicate x values)

        For i = k + 1 To m - 1
            factor = A(i, k) / A(k, k)
            For j = k To m - 1
                A(i, j) = A(i, j) - factor * A(k, j)
            Next j
            b(i) = b(i) - factor * b(k)
        Next i
    Next k

    For i = m - 1 To 0 Step -1
        tmp = b(i)
        For j = i + 1 To m - 1
            tmp = tmp - A(i, j) * beta(j)
        Next j
        beta(i) = tmp / A(i, i)
    Next i

End Sub

'---------------------------------------------------------------------
' In-place recursive quicksort, ascending.
'---------------------------------------------------------------------
Private Sub QuickSort(a() As Double, ByVal lo As Long, ByVal hi As Long)

    Dim i As Long, j As Long
    Dim pivot As Double, temp As Double

    i = lo : j = hi
    pivot = a(Int((lo + hi) / 2))

    Do While i <= j
        Do While a(i) < pivot
            i = i + 1
        Loop
        Do While a(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = a(i) : a(i) = a(j) : a(j) = temp
            i = i + 1 : j = j - 1
        End If
    Loop

    If lo < j Then Call QuickSort(a, lo, j)
    If i < hi Then Call QuickSort(a, i, hi)

End Sub
