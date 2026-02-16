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
    Dim delimT As String, delimP As String

    If Len(Trim$(timesCSV)) = 0 Or Len(Trim$(pricesCSV)) = 0 Then
        ReDim outArr(0 To 0)
        outArr(0).t = -1#
        outArr(0).price = 0#
        ParsePutEvents = outArr
        Exit Function
    End If

    delimT = DetectDelimiter(timesCSV)
    delimP = DetectDelimiter(pricesCSV)

    tArr = Split(timesCSV, delimT)
    pArr = Split(pricesCSV, delimP)
    n = UBound(tArr) - LBound(tArr) + 1

    If n <> (UBound(pArr) - LBound(pArr) + 1) Then
        ReDim outArr(0 To 0)
        outArr(0).t = -1#
        outArr(0).price = 0#
        ParsePutEvents = outArr
        Exit Function
    End If

    ReDim outArr(0 To n - 1)
    For i = 0 To n - 1
        outArr(i).t = ParseFlexibleNumber(Trim$(tArr(i)))
        outArr(i).price = ParseFlexibleNumber(Trim$(pArr(i)))
    Next i

    ParsePutEvents = outArr
End Function

Private Function ParseResetEvents(ByVal timesCSV As String, ByVal floorsCSV As String) As ResetEvent()
    Dim outArr() As ResetEvent
    Dim tArr() As String, fArr() As String
    Dim i As Long, n As Long
    Dim delimT As String, delimF As String

    If Len(Trim$(timesCSV)) = 0 Or Len(Trim$(floorsCSV)) = 0 Then
        ReDim outArr(0 To 0)
        outArr(0).t = -1#
        outArr(0).floorCP = 0#
        ParseResetEvents = outArr
        Exit Function
    End If

    delimT = DetectDelimiter(timesCSV)
    delimF = DetectDelimiter(floorsCSV)

    tArr = Split(timesCSV, delimT)
    fArr = Split(floorsCSV, delimF)
    n = UBound(tArr) - LBound(tArr) + 1


    If n <> (UBound(fArr) - LBound(fArr) + 1) Then
        ReDim outArr(0 To 0)
        outArr(0).t = -1#
        outArr(0).floorCP = 0#
        ParseResetEvents = outArr
        Exit Function
    End If

    ReDim outArr(0 To n - 1)
    For i = 0 To n - 1
        outArr(i).t = ParseFlexibleNumber(Trim$(tArr(i)))
        outArr(i).floorCP = ParseFlexibleNumber(Trim$(fArr(i)))
    Next i

    ParseResetEvents = outArr
End Function



Private Function DetectDelimiter(ByVal csvText As String) As String
    If InStr(1, csvText, ";", vbTextCompare) > 0 Then
        DetectDelimiter = ";"
    Else
        DetectDelimiter = ","
    End If
End Function

Private Function ParseFlexibleNumber(ByVal txt As String) As Double
    Dim s As String
    s = Trim$(txt)
    s = Replace(s, " ", "")

    If InStr(1, s, ",", vbTextCompare) > 0 And InStr(1, s, ".", vbTextCompare) = 0 Then
        s = Replace(s, ",", ".")
    End If

    ParseFlexibleNumber = Val(s)
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


Public Sub CreateCBInputTemplate()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets("CB_Input")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "CB_Input"
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "Convertible Bond Input Template"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "Spot"
    ws.Range("A4").Value = "Face"
    ws.Range("A5").Value = "CouponRate"
    ws.Range("A6").Value = "Maturity (years)"
    ws.Range("A7").Value = "RiskFree"
    ws.Range("A8").Value = "CreditSpread"
    ws.Range("A9").Value = "Vol"
    ws.Range("A10").Value = "DivYield"
    ws.Range("A11").Value = "InitConvPrice"
    ws.Range("A12").Value = "Steps"
    ws.Range("A13").Value = "PutTimesCSV"
    ws.Range("A14").Value = "PutPricesCSV"
    ws.Range("A15").Value = "ResetTimesCSV"
    ws.Range("A16").Value = "ResetFloorsCSV"

    ws.Range("B3").Value = 100
    ws.Range("B4").Value = 1000
    ws.Range("B5").Value = 0.02
    ws.Range("B6").Value = 5
    ws.Range("B7").Value = 0.03
    ws.Range("B8").Value = 0.02
    ws.Range("B9").Value = 0.25
    ws.Range("B10").Value = 0
    ws.Range("B11").Value = 120
    ws.Range("B12").Value = 200
    ws.Range("B13").Value = "2;3;4"
    ws.Range("B14").Value = "1.00;0.99;0.98"
    ws.Range("B15").Value = "1;2;3"
    ws.Range("B16").Value = "115;110;105"

    ws.Range("A18").Value = "CB Price"
    ws.Range("B18").Formula = "=CBPriceWithPutAndReset(B3,B4,B5,B6,B7,B8,B9,B10,B11,B12,B13,B14,B15,B16)"
    ws.Range("B18").Font.Bold = True

    ws.Range("A20").Value = "How to use:"
    ws.Range("A21").Value = "1) Update input values in column B."
    ws.Range("A22").Value = "2) Use ; as list separator (e.g., 2;3;4) to avoid locale issues."
    ws.Range("A23").Value = "3) Keep put/reset time and value counts matched; price updates in B18."

    ws.Columns("A:B").AutoFit
End Sub

Public Function CBPriceFromTemplateSheet(Optional ByVal SheetName As String = "CB_Input") As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetName)

    CBPriceFromTemplateSheet = CBPriceWithPutAndReset( _
        Spot:=ws.Range("B3").Value, _
        Face:=ws.Range("B4").Value, _
        CouponRate:=ws.Range("B5").Value, _
        Maturity:=ws.Range("B6").Value, _
        RiskFree:=ws.Range("B7").Value, _
        CreditSpread:=ws.Range("B8").Value, _
        Vol:=ws.Range("B9").Value, _
        DivYield:=ws.Range("B10").Value, _
        InitConvPrice:=ws.Range("B11").Value, _
        Steps:=CLng(ws.Range("B12").Value), _
        PutTimesCSV:=CStr(ws.Range("B13").Value), _
        PutPricesCSV:=CStr(ws.Range("B14").Value), _
        ResetTimesCSV:=CStr(ws.Range("B15").Value), _
        ResetFloorsCSV:=CStr(ws.Range("B16").Value))
End Function
