Attribute VB_Name = "CBPricing"
Option Explicit

'============================================================
' Convertible Bond Pricing (Binomial Tree)
' Features:
'  - Conversion option
'  - Puttable schedule (holder can put at specific dates)
'  - Resettable conversion price schedule
'
' Notes:
'  - Risk-neutral equity tree, simple reduced-form credit spread discounting.
'  - Coupon assumed fixed rate on face value.
'  - American-style conversion/put exercise at each step where applicable.
'============================================================

Private Type PutEvent
    t As Double      ' time in years
    price As Double  ' put price as % of face (e.g. 1.0 = par)
End Type

Private Type ResetEvent
    t As Double      ' time in years
    floorCP As Double ' minimum conversion price at reset
End Type

Public Function CBPriceWithPutAndReset( _
    ByVal Spot As Double, _
    ByVal Face As Double, _
    ByVal CouponRate As Double, _
    ByVal Maturity As Double, _
    ByVal RiskFree As Double, _
    ByVal CreditSpread As Double, _
    ByVal Vol As Double, _
    ByVal DivYield As Double, _
    ByVal InitConvPrice As Double, _
    ByVal Steps As Long, _
    Optional ByVal PutTimesCSV As String = "", _
    Optional ByVal PutPricesCSV As String = "", _
    Optional ByVal ResetTimesCSV As String = "", _
    Optional ByVal ResetFloorsCSV As String = "" _
) As Double

    Dim dt As Double, u As Double, d As Double, p As Double
    Dim i As Long, j As Long

    If Steps < 2 Then
        CBPriceWithPutAndReset = CVErr(xlErrValue)
        Exit Function
    End If

    If Maturity <= 0# Or Spot <= 0# Or Face <= 0# Or InitConvPrice <= 0# Or Vol < 0# Then
        CBPriceWithPutAndReset = CVErr(xlErrValue)
        Exit Function
    End If

    dt = Maturity / Steps
    u = Exp(Vol * Sqr(dt))
    d = 1# / u

    p = (Exp((RiskFree - DivYield) * dt) - d) / (u - d)
    If p <= 0# Or p >= 1# Then
        CBPriceWithPutAndReset = CVErr(xlErrNum)
        Exit Function
    End If

    Dim puts() As PutEvent
    Dim resets() As ResetEvent
    puts = ParsePutEvents(PutTimesCSV, PutPricesCSV)
    resets = ParseResetEvents(ResetTimesCSV, ResetFloorsCSV)

    Dim bondVals() As Double
    ReDim bondVals(0 To Steps)

    Dim finalCP As Double
    finalCP = EffectiveConvPrice(Maturity, InitConvPrice, resets)

    For j = 0 To Steps
        Dim ST As Double
        ST = Spot * (u ^ j) * (d ^ (Steps - j))
        Dim convShares As Double
        convShares = Face / finalCP

        Dim redeem As Double
        redeem = Face * (1# + CouponRate * dt)

        bondVals(j) = Application.WorksheetFunction.Max(redeem, convShares * ST)
    Next j

    Dim disc As Double
    disc = Exp(-(RiskFree + CreditSpread) * dt)

    For i = Steps - 1 To 0 Step -1
        Dim t As Double
        t = i * dt

        Dim cp As Double
        cp = EffectiveConvPrice(t, InitConvPrice, resets)

        For j = 0 To i
            ST = Spot * (u ^ j) * (d ^ (i - j))
            convShares = Face / cp

            Dim holdVal As Double
            holdVal = disc * (p * bondVals(j + 1) + (1# - p) * bondVals(j)) + Face * CouponRate * dt

            Dim convVal As Double
            convVal = convShares * ST

            Dim putVal As Double
            putVal = PutValueAtTime(t, Face, puts)

            Dim nodeVal As Double
            nodeVal = holdVal
            If convVal > nodeVal Then nodeVal = convVal
            If putVal > nodeVal Then nodeVal = putVal

            bondVals(j) = nodeVal
        Next j
    Next i

    CBPriceWithPutAndReset = bondVals(0)
End Function

Private Function ParsePutEvents(ByVal timesCSV As String, ByVal pricesCSV As String) As PutEvent()
    Dim outArr() As PutEvent
    Dim tArr() As String, pArr() As String
    Dim i As Long, n As Long

    If Len(Trim$(timesCSV)) = 0 Then
        ReDim outArr(0 To -1)
        ParsePutEvents = outArr
        Exit Function
    End If

    tArr = Split(timesCSV, ",")
    pArr = Split(pricesCSV, ",")
    n = UBound(tArr) - LBound(tArr) + 1

    If n <> (UBound(pArr) - LBound(pArr) + 1) Then
        ReDim outArr(0 To -1)
        ParsePutEvents = outArr
        Exit Function
    End If

    ReDim outArr(0 To n - 1)
    For i = 0 To n - 1
        outArr(i).t = CDbl(Trim$(tArr(i)))
        outArr(i).price = CDbl(Trim$(pArr(i)))
    Next i

    ParsePutEvents = outArr
End Function

Private Function ParseResetEvents(ByVal timesCSV As String, ByVal floorsCSV As String) As ResetEvent()
    Dim outArr() As ResetEvent
    Dim tArr() As String, fArr() As String
    Dim i As Long, n As Long

    If Len(Trim$(timesCSV)) = 0 Then
        ReDim outArr(0 To -1)
        ParseResetEvents = outArr
        Exit Function
    End If

    tArr = Split(timesCSV, ",")
    fArr = Split(floorsCSV, ",")
    n = UBound(tArr) - LBound(tArr) + 1


    If n <> (UBound(fArr) - LBound(fArr) + 1) Then
        ReDim outArr(0 To -1)
        ParseResetEvents = outArr
        Exit Function
    End If

    ReDim outArr(0 To n - 1)
    For i = 0 To n - 1
        outArr(i).t = CDbl(Trim$(tArr(i)))
        outArr(i).floorCP = CDbl(Trim$(fArr(i)))
    Next i

    ParseResetEvents = outArr
End Function

Private Function EffectiveConvPrice(ByVal t As Double, ByVal initCP As Double, ByRef resets() As ResetEvent) As Double
    Dim i As Long
    Dim cp As Double
    cp = initCP

    On Error GoTo NoReset
    For i = LBound(resets) To UBound(resets)
        If resets(i).t <= t Then
            If resets(i).floorCP < cp Then cp = resets(i).floorCP
        End If
    Next i

NoReset:
    If cp <= 0# Then cp = initCP
    EffectiveConvPrice = cp
End Function

Private Function PutValueAtTime(ByVal t As Double, ByVal face As Double, ByRef puts() As PutEvent) As Double
    Dim i As Long
    On Error GoTo NoPut

    For i = LBound(puts) To UBound(puts)
        If Abs(puts(i).t - t) < 0.0000001 Then
            PutValueAtTime = face * puts(i).price
            Exit Function
        End If
    Next i

NoPut:
    PutValueAtTime = 0#
End Function

Public Sub Example_CB_Calc()
    Dim px As Double
    px = CBPriceWithPutAndReset( _
        Spot:=100, _
        Face:=1000, _
        CouponRate:=0.02, _
        Maturity:=5, _
        RiskFree:=0.03, _
        CreditSpread:=0.02, _
        Vol:=0.25, _
        DivYield:=0#, _
        InitConvPrice:=120, _
        Steps:=200, _
        PutTimesCSV:="2,3,4", _
        PutPricesCSV:="1.00,0.99,0.98", _
        ResetTimesCSV:="1,2,3", _
        ResetFloorsCSV:="115,110,105")

    Debug.Print "CB price = "; Format(px, "0.0000")
End Sub
