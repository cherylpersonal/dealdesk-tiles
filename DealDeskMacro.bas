Attribute VB_Name = "DealDeskMacro"
Option Explicit

' =====================================================================
'  DAF PROCESSOR  -  DealDeskMacro.bas
'
'  Import this module into DealDesk_Master.xlsm via the VBA editor.
'  Assign ProcessDAF to the button on the Instructions sheet.
'  Run SetupProtection ONCE after import to hide the cost sheet.
' =====================================================================

Private Const COST_SHEET    As String = "CostData"
Private Const TRACKER_SHEET As String = "Deal Desk Tracker"
Private Const SHEET_PW      As String = "DDOnly2026"    ' change as needed

' Approval thresholds
Private Const CEO_SKU_THRESH As Double = 0.05   ' any SKU margin < 5%  -> CEO
Private Const CEO_OVR_THRESH As Double = 0.1    ' overall margin < 10% -> CEO
Private Const PM_SKU_THRESH  As Double = 0.1    ' any SKU margin < 10% -> PM Head
Private Const PM_OVR_THRESH  As Double = 0.2    ' overall margin < 20% -> PM Head

' Colours
Private Const CLR_HDR  As Long = 15132390   ' #E8F0FE  light blue
Private Const CLR_TOT  As Long = 13424076   ' #FFF3CD  amber
Private Const CLR_DD   As Long = 13828828   ' #D4EDDA  green
Private Const CLR_PM   As Long = 13424076   ' #FFF3CD  amber
Private Const CLR_CEO  As Long = 16749784   ' #F8D7DA  red
Private Const CLR_RED  As Long = 16749784   ' margin < 5%
Private Const CLR_AMB  As Long = 13424076   ' margin 5-10%
Private Const CLR_GRN  As Long = 13828828   ' margin >= 10%


' =====================================================================
'  RUN ONCE AFTER IMPORT — hides and protects the CostData sheet
' =====================================================================
Sub SetupProtection()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(COST_SHEET)
    If ws Is Nothing Then
        MsgBox "CostData sheet not found.", vbExclamation
        Exit Sub
    End If
    ws.Unprotect Password:=SHEET_PW
    ws.Protect Password:=SHEET_PW, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
    MsgBox "CostData sheet is now hidden and protected." & vbCrLf & vbCrLf & _
           "Remember to also lock the VBA project:" & vbCrLf & _
           "Tools -> VBAProject Properties -> Protection", vbInformation, "Setup Complete"
End Sub


' =====================================================================
'  MAIN — assign to button
' =====================================================================
Sub ProcessDAF()
    Dim dafPath As String
    dafPath = BrowseForFile()
    If dafPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim dafWB As Workbook
    On Error GoTo ErrHandler
    Set dafWB = Workbooks.Open(dafPath, ReadOnly:=False)

    ' ── Find DAF data sheet & header row ─────────────────────────────
    Dim dataWS As Worksheet
    Dim hRow   As Long
    Set dataWS = FindDataSheet(dafWB, hRow)

    If dataWS Is Nothing Then
        MsgBox "Cannot find DAF data." & vbCrLf & _
               "The file must contain columns: sku_code (or Material), boxes, nef", vbExclamation
        dafWB.Close False
        GoTo Cleanup
    End If

    ' ── Map DAF columns ───────────────────────────────────────────────
    Dim dCols As Object
    Set dCols = MapCols(dataWS, hRow)

    Dim lastRow As Long
    lastRow = dataWS.Cells(dataWS.Rows.Count, dCols("sku_code")).End(xlUp).Row
    Dim n As Long
    n = lastRow - hRow

    If n <= 0 Then
        MsgBox "No data rows found in DAF.", vbExclamation
        dafWB.Close False
        GoTo Cleanup
    End If

    ' ── Temporarily reveal cost sheet ────────────────────────────────
    Dim cWS As Worksheet
    ThisWorkbook.Sheets(COST_SHEET).Visible = xlSheetVisible
    Set cWS = ThisWorkbook.Sheets(COST_SHEET)
    cWS.Unprotect Password:=SHEET_PW

    Dim cHRow  As Long
    cHRow = FindHeaderRowInSheet(cWS, 10)
    Dim cCols As Object
    Set cCols = MapCols(cWS, cHRow)

    ' Cost sheet column aliases
    Dim skuCI   As Long
    Dim priceCI As Long
    Dim areaCI  As Long
    skuCI   = FirstExisting(cCols, Array("material", "sku_code", "sku"))
    priceCI = FirstExisting(cCols, Array("purchase_price/-sft", "buying_price", "purchase_price"))
    areaCI  = FirstExisting(cCols, Array("sft/_box", "area_per_box", "sft/box"))

    If skuCI = 0 Or priceCI = 0 Or areaCI = 0 Then
        MsgBox "Cost sheet columns not recognised. Check CostData headers.", vbExclamation
        cWS.Protect Password:=SHEET_PW, Contents:=True
        ThisWorkbook.Sheets(COST_SHEET).Visible = xlSheetVeryHidden
        dafWB.Close False
        GoTo Cleanup
    End If

    ' ── Build cost lookup dict: UPPER(sku) -> {area, price} ──────────
    Dim cLookup As Object
    Set cLookup = BuildCostLookup(cWS, cHRow, skuCI, priceCI, areaCI)

    ' Re-protect and hide cost sheet
    cWS.Protect Password:=SHEET_PW, Contents:=True
    ThisWorkbook.Sheets(COST_SHEET).Visible = xlSheetVeryHidden

    ' ── Read DAF rows ─────────────────────────────────────────────────
    Dim sku()  As String
    Dim bx()   As Double
    Dim nf()   As Double
    Dim lv()   As Double
    Dim apb()  As Double
    Dim bp()   As Double
    Dim ta()   As Double
    Dim rv()   As Double
    Dim ct()   As Double
    Dim mv()   As Double
    Dim mp()   As Double
    Dim ok()   As Boolean

    ReDim sku(1 To n):  ReDim bx(1 To n):  ReDim nf(1 To n)
    ReDim lv(1 To n):   ReDim apb(1 To n): ReDim bp(1 To n)
    ReDim ta(1 To n):   ReDim rv(1 To n):  ReDim ct(1 To n)
    ReDim mv(1 To n):   ReDim mp(1 To n):  ReDim ok(1 To n)

    Dim meta(8) As String

    Dim i As Long
    For i = 1 To n
        Dim dr As Long
        dr = hRow + i

        sku(i) = UCase(Trim(dataWS.Cells(dr, dCols("sku_code")).Value))
        bx(i)  = SafeD(dataWS.Cells(dr, dCols("boxes")).Value)
        nf(i)  = SafeD(dataWS.Cells(dr, dCols("nef")).Value)
        If dCols.Exists("list_value") Then
            lv(i) = SafeD(dataWS.Cells(dr, dCols("list_value")).Value)
        End If

        If i = 1 Then
            meta(0) = CellVal(dataWS, dr, dCols, Array("daf_ref_no", "daf_ref_nometa", "daf_ref"))
            meta(1) = CellVal(dataWS, dr, dCols, Array("lob"))
            meta(2) = CellVal(dataWS, dr, dCols, Array("channel"))
            meta(3) = CellVal(dataWS, dr, dCols, Array("state"))
            meta(4) = CellVal(dataWS, dr, dCols, Array("project_name"))
            meta(5) = CellVal(dataWS, dr, dCols, Array("developer_name"))
            meta(6) = CellVal(dataWS, dr, dCols, Array("dealer_name"))
            meta(7) = CellVal(dataWS, dr, dCols, Array("zonal_coordinator"))
            meta(8) = CellVal(dataWS, dr, dCols, Array("submitted_date"))
        End If

        If cLookup.Exists(sku(i)) Then
            Dim info As Variant
            info   = cLookup(sku(i))
            apb(i) = info(0)
            bp(i)  = info(1)
            If bx(i) > 0 And apb(i) > 0 Then
                ta(i)  = bx(i) * apb(i)
                rv(i)  = nf(i) * ta(i)
                ct(i)  = bp(i) * ta(i)
                mv(i)  = rv(i) - ct(i)
                If rv(i) > 0 Then mp(i) = mv(i) / rv(i)
                ok(i) = True
            End If
        End If
    Next i

    ' ── Totals ───────────────────────────────────────────────────────
    Dim tRev As Double, tCst As Double, tMV As Double
    Dim tBx  As Double, tTA  As Double, tLV As Double

    For i = 1 To n
        If ok(i) Then
            tRev = tRev + rv(i): tCst = tCst + ct(i)
            tMV  = tMV  + mv(i): tBx  = tBx  + bx(i)
            tTA  = tTA  + ta(i)
        End If
        tLV = tLV + lv(i)
    Next i

    Dim ovMP As Double
    If tRev > 0 Then ovMP = tMV / tRev

    ' ── Approval ─────────────────────────────────────────────────────
    Dim appLvl    As String
    Dim appReason As String
    appLvl = CalcApproval(n, sku, mp, ok, ovMP, appReason)

    ' ── Write output ─────────────────────────────────────────────────
    WriteCalcSheet dafWB, n, sku, bx, nf, apb, bp, ta, rv, ct, mv, mp, ok, _
                   tBx, tTA, tRev, tCst, tMV, ovMP, appLvl, appReason

    AppendTracker meta, tLV, tRev, tMV, ovMP, appLvl

    dafWB.Save
    ThisWorkbook.Save

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Done!" & vbCrLf & vbCrLf & _
           "  Calculations sheet added to: " & dafWB.Name & vbCrLf & _
           "  Deal Desk Tracker updated.", vbInformation, "DAF Processor"
    Exit Sub

ErrHandler:
    On Error Resume Next
    cWS.Protect Password:=SHEET_PW, Contents:=True
    ThisWorkbook.Sheets(COST_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DAF Processor"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


' =====================================================================
'  HELPERS
' =====================================================================

Private Function BrowseForFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the filled DAF Excel file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then BrowseForFile = .SelectedItems(1)
    End With
End Function

' Scans all sheets in a workbook for one that looks like a DAF data sheet
Private Function FindDataSheet(wb As Workbook, ByRef hRow As Long) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim r As Long
        r = FindHeaderRowInSheet(ws, 10)
        If r > 0 Then
            Dim cols As Object
            Set cols = MapCols(ws, r)
            If cols.Exists("sku_code") And cols.Exists("boxes") And cols.Exists("nef") Then
                Set FindDataSheet = ws
                hRow = r
                Exit Function
            End If
        End If
    Next ws
End Function

' Returns the first row (1-based) where 2+ known DAF/cost keywords appear
Private Function FindHeaderRowInSheet(ws As Worksheet, maxScan As Long) As Long
    Dim hints As Variant
    hints = Array("sku_code", "sku", "material", "boxes", "nef", "buying_price", _
                  "purchase_price", "area_per_box", "sft", "channel", "lob")
    Dim r As Long, c As Long
    Dim lastCol As Long
    lastCol = ws.UsedRange.Columns.Count
    For r = 1 To maxScan
        Dim score As Long
        score = 0
        For c = 1 To lastCol
            Dim v As String
            v = NormKey(CStr(ws.Cells(r, c).Value))
            Dim h As Variant
            For Each h In hints
                If InStr(1, v, CStr(h), vbTextCompare) > 0 Then
                    score = score + 1
                    Exit For
                End If
            Next h
            If score >= 2 Then
                FindHeaderRowInSheet = r
                Exit Function
            End If
        Next c
    Next r
End Function

' Returns a Scripting.Dictionary: normalised_col_name -> column index
Private Function MapCols(ws As Worksheet, hRow As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' case-insensitive
    Dim lastCol As Long
    lastCol = ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim k As String
        k = NormKey(CStr(ws.Cells(hRow, c).Value))
        If Len(k) > 0 And Not d.Exists(k) Then d(k) = c
    Next c
    Set MapCols = d
End Function

' lowercase + trim + spaces->underscores
Private Function NormKey(s As String) As String
    NormKey = LCase(Trim(Replace(s, " ", "_")))
End Function

' Returns column index of first matching alias; 0 if none found
Private Function FirstExisting(cols As Object, aliases As Variant) As Long
    Dim a As Variant
    For Each a In aliases
        If cols.Exists(CStr(a)) Then
            FirstExisting = cols(CStr(a))
            Exit Function
        End If
    Next a
End Function

' Returns cell value string for first alias found in cols dict
Private Function CellVal(ws As Worksheet, r As Long, cols As Object, aliases As Variant) As String
    Dim a As Variant
    For Each a In aliases
        If cols.Exists(CStr(a)) Then
            CellVal = CStr(ws.Cells(r, cols(CStr(a))).Value)
            Exit Function
        End If
    Next a
End Function

' Build lookup dict: UPPER(sku) -> Array(area_per_box, buying_price)
Private Function BuildCostLookup(ws As Worksheet, hRow As Long, _
                                  skuCI As Long, priceCI As Long, areaCI As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, skuCI).End(xlUp).Row
    Dim r As Long
    For r = hRow + 1 To lastRow
        Dim k As String
        k = UCase(Trim(ws.Cells(r, skuCI).Value))
        If Len(k) > 0 And Not d.Exists(k) Then
            d(k) = Array(SafeD(ws.Cells(r, areaCI).Value), SafeD(ws.Cells(r, priceCI).Value))
        End If
    Next r
    Set BuildCostLookup = d
End Function

' Safe numeric conversion
Private Function SafeD(v As Variant) As Double
    If IsNumeric(v) Then SafeD = CDbl(v) Else SafeD = 0
End Function

' Apply approval matrix; returns level and populates reason string
Private Function CalcApproval(n As Long, sku() As String, mp() As Double, _
                               ok() As Boolean, ovMP As Double, ByRef reason As String) As String
    Dim below5  As String
    Dim below10 As String
    Dim i As Long

    For i = 1 To n
        If ok(i) Then
            If mp(i) < CEO_SKU_THRESH Then
                If Len(below5) > 0 Then below5 = below5 & ", "
                below5 = below5 & sku(i)
            ElseIf mp(i) < PM_SKU_THRESH Then
                If Len(below10) > 0 Then below10 = below10 & ", "
                below10 = below10 & sku(i)
            End If
        End If
    Next i

    reason = ""
    If Len(below5) > 0 Or ovMP < CEO_OVR_THRESH Then
        CalcApproval = "CEO"
        If Len(below5) > 0 Then
            reason = "SKU margin < 5%: " & below5
        End If
        If ovMP < CEO_OVR_THRESH Then
            If Len(reason) > 0 Then reason = reason & "  |  "
            reason = reason & "Overall margin " & Format(ovMP * 100, "0.0") & "% < 10%"
        End If
    ElseIf Len(below10) > 0 Or ovMP < PM_OVR_THRESH Then
        CalcApproval = "PM Head"
        If Len(below10) > 0 Then
            reason = "SKU margin < 10%: " & below10
        End If
        If ovMP < PM_OVR_THRESH Then
            If Len(reason) > 0 Then reason = reason & "  |  "
            reason = reason & "Overall margin " & Format(ovMP * 100, "0.0") & "% < 20%"
        End If
    Else
        CalcApproval = "Deal Desk"
        reason = "All SKU margins >= 10%  |  Overall margin " & Format(ovMP * 100, "0.0") & "% >= 20%"
    End If
End Function

' Returns the approval colour constant for a given level
Private Function ApprovalColour(lvl As String) As Long
    Select Case lvl
        Case "Deal Desk": ApprovalColour = CLR_DD
        Case "PM Head":   ApprovalColour = CLR_PM
        Case "CEO":       ApprovalColour = CLR_CEO
        Case Else:        ApprovalColour = RGB(255, 255, 255)
    End Select
End Function


' =====================================================================
'  WRITE CALCULATIONS SHEET  (inside the DAF workbook)
' =====================================================================
Private Sub WriteCalcSheet(dafWB As Workbook, n As Long, _
    sku() As String, bx() As Double, nf() As Double, apb() As Double, bp() As Double, _
    ta() As Double, rv() As Double, ct() As Double, mv() As Double, mp() As Double, ok() As Boolean, _
    tBx As Double, tTA As Double, tRev As Double, tCst As Double, tMV As Double, ovMP As Double, _
    appLvl As String, appReason As String)

    ' Remove existing Calculations sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    dafWB.Sheets("Calculations").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = dafWB.Sheets.Add(After:=dafWB.Sheets(dafWB.Sheets.Count))
    ws.Name = "Calculations"

    ' ── Column headers ─────────────────────────────────────────────────
    Dim hdrs As Variant
    hdrs = Array("SKU Code", "Boxes", "NEF (per sqft)", "Area per Box (sqft)", _
                 "Buying Price (per sqft)", "Total Area (sqft)", "Revenue", _
                 "Cost", "Margin Value", "Margin %")
    Dim c As Long
    For c = 1 To 10
        With ws.Cells(1, c)
            .Value = hdrs(c - 1)
            .Font.Bold = True
            .Interior.Color = CLR_HDR
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next c
    ws.Rows(1).RowHeight = 30

    ' ── Data rows ──────────────────────────────────────────────────────
    Dim i As Long
    For i = 1 To n
        Dim r As Long
        r = i + 1
        ws.Cells(r, 1).Value = sku(i)
        ws.Cells(r, 2).Value = bx(i)
        ws.Cells(r, 3).Value = nf(i):   ws.Cells(r, 3).NumberFormat = "#,##0.00"
        ws.Cells(r, 4).Value = apb(i):  ws.Cells(r, 4).NumberFormat = "#,##0.00"
        ws.Cells(r, 5).Value = bp(i):   ws.Cells(r, 5).NumberFormat = "#,##0.00"

        If ok(i) Then
            ws.Cells(r, 6).Value = ta(i):  ws.Cells(r, 6).NumberFormat = "#,##0.00"
            ws.Cells(r, 7).Value = rv(i):  ws.Cells(r, 7).NumberFormat = "#,##0.00"
            ws.Cells(r, 8).Value = ct(i):  ws.Cells(r, 8).NumberFormat = "#,##0.00"
            ws.Cells(r, 9).Value = mv(i):  ws.Cells(r, 9).NumberFormat = "#,##0.00"
            ws.Cells(r, 10).Value = mp(i): ws.Cells(r, 10).NumberFormat = "0.00%"

            Dim mc As Long
            If mp(i) < 0.05 Then
                mc = CLR_RED
            ElseIf mp(i) < 0.1 Then
                mc = CLR_AMB
            Else
                mc = CLR_GRN
            End If
            ws.Cells(r, 10).Interior.Color = mc
        Else
            ws.Cells(r, 6).Value = "Not in cost sheet"
            ws.Cells(r, 6).Font.Italic = True
            ws.Cells(r, 6).Font.Color = RGB(100, 100, 100)
        End If
    Next i

    ' ── Total row ──────────────────────────────────────────────────────
    Dim tRow As Long
    tRow = n + 2
    ws.Cells(tRow, 1).Value = "TOTAL"
    ws.Cells(tRow, 2).Value = tBx
    ws.Cells(tRow, 6).Value = tTA:  ws.Cells(tRow, 6).NumberFormat = "#,##0.00"
    ws.Cells(tRow, 7).Value = tRev: ws.Cells(tRow, 7).NumberFormat = "#,##0.00"
    ws.Cells(tRow, 8).Value = tCst: ws.Cells(tRow, 8).NumberFormat = "#,##0.00"
    ws.Cells(tRow, 9).Value = tMV:  ws.Cells(tRow, 9).NumberFormat = "#,##0.00"
    ws.Cells(tRow, 10).Value = ovMP: ws.Cells(tRow, 10).NumberFormat = "0.00%"

    For c = 1 To 10
        With ws.Cells(tRow, c)
            .Font.Bold = True
            .Interior.Color = CLR_TOT
            .Borders.LineStyle = xlContinuous
        End With
    Next c

    ' ── Approval block ─────────────────────────────────────────────────
    Dim aRow As Long
    aRow = tRow + 2

    With ws.Cells(aRow, 1)
        .Value = "Recommended Approval Level"
        .Font.Bold = True
        .Interior.Color = CLR_HDR
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Cells(aRow, 2)
        .Value = appLvl
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = ApprovalColour(appLvl)
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Cells(aRow + 1, 1)
        .Value = "Reason(s)"
        .Font.Bold = True
        .Interior.Color = CLR_HDR
        .Borders.LineStyle = xlContinuous
    End With
    ws.Cells(aRow + 1, 2).Value = appReason
    ws.Cells(aRow + 1, 2).WrapText = True

    ' ── Column widths ──────────────────────────────────────────────────
    Dim widths As Variant
    widths = Array(20, 8, 15, 18, 20, 16, 14, 14, 15, 12)
    For c = 1 To 10
        ws.Columns(c).ColumnWidth = widths(c - 1)
    Next c

    ' Freeze header row
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
End Sub


' =====================================================================
'  APPEND TRACKER ROW  (into Deal Desk Tracker sheet of this workbook)
' =====================================================================
Private Sub AppendTracker(meta() As String, tLV As Double, tRev As Double, _
                           tMV As Double, ovMP As Double, appLvl As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TRACKER_SHEET)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' 6-hr due date from submitted_date
    Dim sixHrStr As String
    If Len(Trim(meta(8))) > 0 Then
        On Error Resume Next
        Dim subDt As Date
        subDt = CDate(meta(8))
        If Err.Number = 0 Then
            sixHrStr = Format(subDt + TimeSerial(6, 0, 0), "YYYY-MM-DD HH:MM")
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' Avg discount %
    Dim avgDiscStr As String
    If tLV > 0 Then
        avgDiscStr = Format((tLV - tRev) / tLV * 100, "0.00") & "%"
    End If

    ' Write values in tracker column order (must match TRACKER_HEADERS in build_master.py)
    Dim vals(22) As Variant
    vals(0)  = meta(0)         ' DAF Reference No.
    vals(1)  = meta(1)         ' LOB
    vals(2)  = meta(2)         ' Channel
    vals(3)  = meta(3)         ' State
    vals(4)  = meta(4)         ' Project Name
    vals(5)  = meta(5)         ' Developer Name
    vals(6)  = meta(6)         ' Dealer Name
    vals(7)  = meta(7)         ' Zonal Coordinator
    vals(8)  = meta(8)         ' Submitted Date
    vals(9)  = ""              ' DD Received Date    (manual)
    vals(10) = sixHrStr        ' 6-Hr Due Date
    vals(11) = ""              ' Response Date       (manual)
    vals(12) = ""              ' TAT                 (manual)
    vals(13) = ""              ' SLA Met             (manual)
    vals(14) = ""              ' Quote Version       (manual)
    vals(15) = appLvl          ' Approval Level
    vals(16) = ""              ' Approver Name       (manual)
    vals(17) = ""              ' Approval Status     (manual)
    vals(18) = ""              ' Approval Date       (manual)
    vals(19) = tLV             ' List Value
    vals(20) = Round(tRev, 2)  ' Deal Value
    vals(21) = avgDiscStr      ' Avg Discount %
    vals(22) = Format(ovMP * 100, "0.00") & "%"  ' Deal Margin %

    Dim c As Long
    For c = 0 To 22
        ws.Cells(nextRow, c + 1).Value = vals(c)
    Next c

    ' Colour-code the Approval Level cell (column 16)
    With ws.Cells(nextRow, 16)
        .Interior.Color = ApprovalColour(appLvl)
        .Font.Bold = True
    End With

    ' Thin border on new row
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 23)).Borders.LineStyle = xlContinuous
End Sub
