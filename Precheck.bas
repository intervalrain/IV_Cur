Public WaferSet As Boolean
Option Explicit

Public Sub LoadPrecheck()

    Dim i As Long, j As Long
    Dim nowSheet As Worksheet
    
    Speed
    
    Set nowSheet = Worksheets("Precheck")
    '==========get Precheck List==========
    Dim DS As String
    Dim DE As String
    Dim GS As String
    Dim GE As String
    Dim dAry()
    Dim gAry()
    Dim dCheck As Boolean
    Dim gCheck As Boolean
    Dim nowRow As Long
    Dim nowCol As Long
    
    dCheck = False
    gCheck = False
    WaferSet = False
    nowRow = 1
    nowCol = 1
    Do
        If UCase(nowSheet.Cells(nowRow, nowCol).Value) = "IDVD" Then
            DS = nowSheet.Cells(nowRow, nowCol).Address
            Do Until nowSheet.Cells(nowRow, nowCol).Value = "": nowCol = nowCol + 1: Loop
            nowCol = nowCol - 1
            Do Until nowSheet.Cells(nowRow, nowCol).Value = "": nowRow = nowRow + 1: Loop
            nowRow = nowRow - 1
            DE = nowSheet.Cells(nowRow, nowCol).Address
        End If
        nowCol = 1
        If UCase(nowSheet.Cells(nowRow, nowCol).Value) = "IDVG" Then
            GS = nowSheet.Cells(nowRow, nowCol).Address
            Do Until nowSheet.Cells(nowRow, nowCol).Value = "": nowCol = nowCol + 1: Loop
            nowCol = nowCol - 1
            Do Until nowSheet.Cells(nowRow, nowCol).Value = "": nowRow = nowRow + 1: Loop
            nowRow = nowRow - 1
            GE = nowSheet.Cells(nowRow, nowCol).Address
        End If
        nowRow = nowRow + 1
    Loop Until nowRow >= nowSheet.UsedRange.Rows.Count
    
    If Not DS = "" And Not DE = "" Then dAry = nowSheet.Range(DS & ":" & DE): dCheck = True
    If Not GS = "" And Not GE = "" Then gAry = nowSheet.Range(GS & ":" & GE): gCheck = True
    
    '==========get file path==========
    Dim tempFile As String
    Dim Path As String
    Dim AryFileName() As String
    
    tempFile = Application.GetOpenFilename("txt File, *.txt", 1, "Load text file", "Open", False)
    AryFileName = Split(tempFile, "\")
    Path = Replace(tempFile, AryFileName(UBound(AryFileName)), "") & getCOL(AryFileName(UBound(AryFileName)), "_", 1) & "_" & getCOL(AryFileName(UBound(AryFileName)), "_", 2) & "_"
    If tempFile = "False" Then Exit Sub
    
    '==========get file name==========
    For i = 3 To UBound(dAry, 1)
        If Not dCheck Then Exit For
        If IsExistSheet(CStr(dAry(i, 2))) Then Worksheets(CStr(dAry(i, 2))).Delete
    Next i
    For i = 3 To UBound(gAry, 1)
        If Not gCheck Then Exit For
        If IsExistSheet(CStr(gAry(i, 2))) Then Worksheets(CStr(gAry(i, 2))).Delete
    Next i
    
    If dCheck Then
        For i = 3 To UBound(dAry, 1)
            For j = 3 To UBound(dAry, 2)
                Call LoadFilePlus(Path & CStr(dAry(i, j)) & ".txt", CStr(dAry(i, 2)), CInt(dAry(1, j)))
            Next j
            If WaferSet = False Then Call Select_Wafer
            Call ArrangeIDVDsheet(i, dAry)
            Call PreCheckPlot(CInt(dAry(1, UBound(dAry, 2))), 1)
            ActiveSheet.Cells(1, Range("BoundRange").Columns.Count * 3 - 7).Activate
            ActiveWindow.SmallScroll ToRight:=Range("BoundRange").Columns.Count * 3 - 7
        Next i
    End If
        
    If gCheck Then
        For i = 3 To UBound(gAry, 1)
            For j = 3 To UBound(gAry, 2)
                Call LoadFilePlus(Path & CStr(gAry(i, j)) & ".txt", CStr(gAry(i, 2)), CInt(gAry(1, j)))
            Next j
            If WaferSet = False Then Call Select_Wafer
            Call ArrangeIDVGsheet(i, gAry)
            Call PreCheckPlot(CInt(gAry(1, UBound(gAry, 2))), 2)
            Call ArrangeGmVGSheet
            Call PreCheckPlot(CInt(gAry(1, UBound(gAry, 2))), 3)
            Call ArrangeSSVGSheet
            Call PreCheckPlot(CInt(gAry(1, UBound(gAry, 2))), 4)
            ActiveSheet.Cells(1, Range("BoundRange").Columns.Count * 3 - 7).Activate
            ActiveWindow.SmallScroll ToRight:=Range("BoundRange").Columns.Count * 3 - 7
        Next i
    End If
    PreCheckSummary
    Unspeed
    On Error GoTo 0

End Sub

Private Function LoadFilePlus(FileAddress As Variant, NewWorkSheetName As String, mLevel As Integer)

    Dim OpenFile As String
    Dim nowSheet As Worksheet
    Dim mRange As Range
    Dim ColBound As Long
    Dim RowBound As Long
    Dim setBound As Boolean
    Dim nowRow As Long
    Dim nowCol As Long
    
    OpenFile = "TEXT;" & FileAddress
    
    If Len(NewWorkSheetName) > 31 Then NewWorkSheetName = Right(NewWorkSheetName, 31)
    
    setBound = False
    If Not IsExistSheet(NewWorkSheetName) Then
        Set nowSheet = AddSheet(NewWorkSheetName)
        Set mRange = nowSheet.Cells(1, 1)
        setBound = True
    Else
        Set nowSheet = Worksheets(NewWorkSheetName)
        ColBound = nowSheet.Range("BoundRange").Columns.Count
        RowBound = nowSheet.Range("BoundRange").Rows.Count
        
        nowCol = (mLevel - 1) * ColBound + 1
        nowRow = 1
        
        Do Until nowSheet.Cells(nowRow, nowCol).Value = ""
            nowRow = nowRow + 1
        Loop
        Set mRange = nowSheet.Cells(nowRow, nowCol)
    End If
    
    With nowSheet.QueryTables.Add(Connection:=OpenFile, Destination:=mRange)
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1251
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=True
    End With
    If setBound = True Then nowSheet.Names.Add "BoundRange", mRange.CurrentRegion

End Function

Private Sub ArrangeIDVDsheet(mRow, dAry)
    Dim mSheet As String
    Dim nowSheet As Worksheet
    Dim nowRow As Long, nowCol As Long
    Dim srcRow As Long, srcCol As Long
    Dim nowRange As Range
    Dim srcRange As Range
    Dim ColBound As Long
    Dim RowBound As Long
    Dim waferNum As Integer
    Dim Col_ID As Integer
    Dim Col_VD As Integer
    Dim Col_VG As Integer
    Dim StressTimes As Integer
    Dim Condition
    Dim i As Integer, j As Integer, k As Integer
    Dim iSrc As Integer, iNow As Integer
    Dim n0 As Long, nv1 As Long, nv2 As Long
    Dim VCC As Double
  
    mSheet = dAry(mRow, 2)
    Set nowSheet = Worksheets(mSheet)
    ColBound = nowSheet.Range("BoundRange").Columns.Count
    RowBound = nowSheet.Range("BoundRange").Rows.Count
    
    waferNum = WorksheetFunction.CountA(Range(Cells(1, ColBound - 2), Cells(RowBound, ColBound - 2)))
    StressTimes = dAry(1, UBound(dAry, 2))
    ReDim Condition(1 To (UBound(dAry, 2) - 2) / StressTimes)
    For i = 1 To UBound(Condition)
        Condition(i) = dAry(2, 3 + (i - 1) * StressTimes)
        If dAry(mRow, 1) = -1 And InStr(Condition(i), "[Sign]") > 0 Then
            Condition(i) = getCOL(Condition(i), "=", 1) & "=" & CDbl(Replace(getCOL(Condition(i), "=", 2), "[Sign]", "")) * -1
        ElseIf InStr(Condition(i), "[Sign]") > 0 Then
            Condition(i) = Replace(Condition(i), "[Sign]", "")
        End If
    Next i
    
    For nowCol = 1 To ColBound
        If nowSheet.Cells(1, nowCol).Value = "ID" Then Col_ID = nowCol
        If nowSheet.Cells(1, nowCol).Value = "VD" Then Col_VD = nowCol
        If nowSheet.Cells(1, nowCol).Value = "VG" Then Col_VG = nowCol
    Next nowCol
    
    nowRow = 3: nowCol = Col_VD
    Do Until nowSheet.Cells(nowRow, nowCol) = nowSheet.Cells(2, Col_VD)
        nowRow = nowRow + 1
    Loop
    
    n0 = RowBound / waferNum - 2
    nv1 = nowRow - 2
    nv2 = n0 / nv1
        
    'Filter ID-VD
    nowRow = 1
    nowCol = 3
    srcRow = 2
    srcCol = Col_ID
    i = 1       'Condition Control
    j = 1       'Stress time
                'Source Range Control
    For k = 0 To waferNum
        If Not WaferArray(1, k) = "NO" Then iSrc = k: Exit For
        If k = waferNum Then Exit For
    Next k
    iNow = 0    'Print Range Control
    Set srcRange = Range(Cells((n0 + 2) * iSrc + 1, 1), Cells((n0 + 2) * (iSrc + 1), ColBound * 3))
    Set nowRange = Range(Cells((nv1 + 1) * iNow + 1, ColBound * 3 + 6), Cells((nv1 + 1) * (iNow + 1) - 1, ColBound * 3 + 8 + nv2 * StressTimes))
    nowRange.Cells(nowRow, nowCol - 2) = Condition(i)
    nowRange.Cells(nowRow + 1, nowCol - 2) = srcRange.Cells(srcRow - 1, ColBound - 2) & srcRange.Cells(srcRow - 1, ColBound - 1)
    nowRange.Range(Cells(1, nowCol - 1), Cells((nv1 + 1), nowCol - 1)).Value = srcRange.Range(Cells(1, Col_VD), Cells((nv1 + 1), Col_VD)).Value
    Do
        'Change Source & Print Range
        If j > StressTimes Then
            iNow = iNow + 1

            If i >= UBound(Condition) Then
                i = 1
                iSrc = iSrc - (waferNum * (UBound(Condition) - 1) - 1)
                For k = iSrc To waferNum - 1
                    If Not WaferArray(1, k) = "NO" Then iSrc = k: Exit For
                    If k = waferNum - 1 Then
                        Exit Do
                    End If
                    iSrc = iSrc + 1
                Next k
            Else
                i = i + 1
                iSrc = iSrc + waferNum
            End If
            
            j = 1
            nowRow = 1
            nowCol = 3
            srcRow = 2
            srcCol = Col_ID
            Set srcRange = Range(Cells((n0 + 2) * iSrc + 1, 1), Cells((n0 + 2) * (iSrc + 1), ColBound * 3))
            Set nowRange = Range(Cells((nv1 + 2) * iNow + 1, ColBound * 3 + 6), Cells((nv1 + 2) * (iNow + 1) - 1, ColBound * 3 + 8 + nv2 * StressTimes))
            nowRange.Cells(nowRow, nowCol - 2) = Condition(i)
            nowRange.Cells(nowRow + 1, nowCol - 2) = srcRange.Cells(srcRow - 1, ColBound - 2) & srcRange.Cells(srcRow - 1, ColBound - 1)
            nowRange.Range(Cells(1, nowCol - 1), Cells((nv1 + 1), nowCol - 1)).Value = srcRange.Range(Cells(1, Col_VD), Cells((nv1 + 1), Col_VD)).Value
            'Terminate of do-loop
            If iSrc = waferNum * UBound(Condition) Then Exit Do
        End If
        'Print
        If nowRow = 1 Then
            nowRange.Cells(nowRow, nowCol) = "ID" & j & "(Vg=" & srcRange.Cells(srcRow + 1, Col_VG) & ")"
            nowRow = nowRow + 1
        Else
            nowRange.Cells(nowRow, nowCol) = "=ABS(" & srcRange.Cells(srcRow, srcCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            nowRow = nowRow + 1
            srcRow = srcRow + 1
        End If
        'Split Vg
        If nowRow > nv1 + 1 Then
            nowRow = 1
            nowCol = nowCol + 1
        End If
        If srcRow > n0 + 1 Then
            srcRow = 2
            srcCol = srcCol + ColBound
            j = j + 1
        End If
    Loop
    VCC = CDbl(getCOL(getCOL(nowSheet.Cells(1, nowSheet.UsedRange.Columns.Count), "=", 2), ")", 1))
    Call PreCheckSummaryIDVD(VCC)
    
End Sub

Private Sub ArrangeIDVGsheet(mRow, gAry)
    Dim mSheet As String
    Dim nowSheet As Worksheet
    Dim nowRow As Long, nowCol As Long
    Dim srcRow As Long, srcCol As Long
    Dim nowRange As Range
    Dim srcRange As Range
    Dim ColBound As Long
    Dim RowBound As Long
    Dim waferNum As Integer
    Dim Col_ID As Integer
    Dim Col_VG As Integer
    Dim Col_VB As Integer
    Dim StressTimes As Integer
    Dim Condition
    Dim i As Integer, j As Integer, k As Integer
    Dim iSrc As Integer, iNow As Integer
    Dim n0 As Long, nv1 As Long, nv2 As Long
    Dim VCC As Double
    
    mSheet = gAry(mRow, 2)
    Set nowSheet = Worksheets(mSheet)
    ColBound = nowSheet.Range("BoundRange").Columns.Count
    RowBound = nowSheet.Range("BoundRange").Rows.Count
    
    waferNum = WorksheetFunction.CountA(Range(Cells(1, ColBound - 2), Cells(RowBound, ColBound - 2)))
    StressTimes = gAry(1, UBound(gAry, 2))
    ReDim Condition(1 To (UBound(gAry, 2) - 2) / StressTimes)
    For i = 1 To UBound(Condition)
        Condition(i) = gAry(2, 3 + (i - 1) * StressTimes)
        If gAry(mRow, 1) = -1 And InStr(Condition(i), "[Sign]") > 0 Then
            Condition(i) = getCOL(Condition(i), "=", 1) & "=" & CDbl(Replace(getCOL(Condition(i), "=", 2), "[Sign]", "")) * -1
        ElseIf InStr(Condition(i), "[Sign]") > 0 Then
            Condition(i) = Replace(Condition(i), "[Sign]", "")
        End If
    Next i
    
    For nowCol = 1 To ColBound
        If nowSheet.Cells(1, nowCol).Value = "ID" Then Col_ID = nowCol
        If nowSheet.Cells(1, nowCol).Value = "VG" Then Col_VG = nowCol
        If nowSheet.Cells(1, nowCol).Value = "VB" Then Col_VB = nowCol
    Next nowCol
    
    nowRow = 3: nowCol = Col_VG
    Do Until nowSheet.Cells(nowRow, nowCol) = nowSheet.Cells(2, Col_VG)
        nowRow = nowRow + 1
    Loop
    
    n0 = RowBound / waferNum - 2
    nv1 = nowRow - 2
    nv2 = n0 / nv1
        
    'Filter ID-VD
    nowRow = 1
    nowCol = 3
    srcRow = 2
    srcCol = Col_ID
    i = 1       'Condition Control
    j = 1       'Stress time
                'Source Range Control
    For k = 0 To waferNum
        If Not WaferArray(1, k) = "NO" Then iSrc = k: Exit For
        If k = waferNum Then Exit For
    Next k
    iNow = 0    'Print Range Control
    Set srcRange = Range(Cells((n0 + 2) * iSrc + 1, 1), Cells((n0 + 2) * (iSrc + 1), ColBound * 3))
    Set nowRange = Range(Cells((nv1 + 1) * iNow + 1, ColBound * 3 + 6), Cells((nv1 + 1) * (iNow + 1) - 1, ColBound * 3 + 8 + nv2 * StressTimes))
    nowRange.Cells(nowRow, nowCol - 2) = Condition(i)
    nowRange.Cells(nowRow + 1, nowCol - 2) = srcRange.Cells(srcRow - 1, ColBound - 2) & srcRange.Cells(srcRow - 1, ColBound - 1)
    nowRange.Range(Cells(1, nowCol - 1), Cells((nv1 + 1), nowCol - 1)).Value = srcRange.Range(Cells(1, Col_VG), Cells((nv1 + 1), Col_VG)).Value
    Do
        'Change Source & Print Range
        If j > StressTimes Then
            iNow = iNow + 1

            If i >= UBound(Condition) Then
                i = 1
                iSrc = iSrc - (waferNum * (UBound(Condition) - 1) - 1)
                For k = iSrc To waferNum - 1
                    If Not WaferArray(1, k) = "NO" Then iSrc = k: Exit For
                    If k = waferNum - 1 Then Exit Do
                    iSrc = iSrc + 1
                Next k
            Else
                i = i + 1
                iSrc = iSrc + waferNum
            End If
            
            j = 1
            nowRow = 1
            nowCol = 3
            srcRow = 2
            srcCol = Col_ID
            Set srcRange = Range(Cells((n0 + 2) * iSrc + 1, 1), Cells((n0 + 2) * (iSrc + 1), ColBound * 3))
            Set nowRange = Range(Cells((nv1 + 2) * iNow + 1, ColBound * 3 + 6), Cells((nv1 + 2) * (iNow + 1) - 1, ColBound * 3 + 8 + nv2 * StressTimes))
            nowRange.Cells(nowRow, nowCol - 2) = Condition(i)
            nowRange.Cells(nowRow + 1, nowCol - 2) = srcRange.Cells(srcRow - 1, ColBound - 2) & srcRange.Cells(srcRow - 1, ColBound - 1)
            nowRange.Range(Cells(1, nowCol - 1), Cells((nv1 + 1), nowCol - 1)).Value = srcRange.Range(Cells(1, Col_VG), Cells((nv1 + 1), Col_VG)).Value
            'Terminate of do-loop
            If iSrc = waferNum * UBound(Condition) Then Exit Do
        End If
        'Print
        If nowRow = 1 Then
            nowRange.Cells(nowRow, nowCol) = "ID" & j & "(Vb=" & srcRange.Cells(srcRow + 1, Col_VB) & ")"
            nowRow = nowRow + 1
        Else
            nowRange.Cells(nowRow, nowCol) = "=ABS(" & srcRange.Cells(srcRow, srcCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            nowRow = nowRow + 1
            srcRow = srcRow + 1
        End If
        'Split Vg
        If nowRow > nv1 + 1 Then
            nowRow = 1
            nowCol = nowCol + 1
        End If
        If srcRow > n0 + 1 Then
            srcRow = 2
            srcCol = srcCol + ColBound
            j = j + 1
        End If
    Loop
'    Dim mRange As Range
'    Set mRange = nowSheet.Cells(1, Range("BoundRange").Columns.Count * 3 + 6).CurrentRegion
'    Set mRange = Range(mRange.Row + mRange.Rows.Count + 1, mRange.Column).CurrentRegion
'    Vcc = CDbl(getCOL(mRange.Cells(1, 1).Value, "=", 2))
    
End Sub
Private Sub ArrangeGmVGSheet()
    Dim nowSheet As Worksheet
    Dim srcRange As Range
    Dim nowRange As Range
    
    Dim nowRow As Long
    Dim nowCol As Long
    Dim i As Long, j As Long
    Dim VCC As Double
    Dim VBP As Double
    
    Set nowSheet = ActiveSheet
    
    Set srcRange = nowSheet.Cells(1, nowSheet.UsedRange.Columns.Count).CurrentRegion
    srcRange.Copy
    nowSheet.Cells(1, srcRange.Column + srcRange.Columns.Count + 1).Select
    Selection.PasteSpecial
    Set nowRange = Selection

    Do Until srcRange.Cells(1, 1) = ""
        For j = 3 To nowRange.Columns.Count
            nowRange.Cells(1, j).Value = Replace(srcRange.Cells(1, j).Value, "ID", "GM")
            For i = 2 To nowRange.Rows.Count - 1
                nowRange.Cells(i, j).Value = "=(" & srcRange.Cells(i + 1, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "-" & srcRange.Cells(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/ABS(" & srcRange.Cells(i + 1, 2).Address(RowAbsolute:=False) & "-" & srcRange.Cells(i, 2).Address(RowAbsolute:=False) & ")"
            Next i
        Next j
        nowRange.Rows(i).ClearContents
        For j = 3 To nowRange.Columns.Count Step ((nowRange.Columns.Count - 2) / 3)
            nowRange.Cells(i, j).Value = "=MAX(" & Range(nowRange.Cells(1, j), nowRange.Cells(nowRange.Rows.Count - 1, j)).Address & ")"
        Next j
        VCC = CDbl(getCOL(nowRange.Cells(1, 1).Value, "=", 2))
        Set srcRange = nowSheet.Cells(srcRange.Row + srcRange.Rows.Count + 1, srcRange.Column).CurrentRegion
        srcRange.Copy
        nowSheet.Cells(srcRange.Row, srcRange.Column + srcRange.Columns.Count + 1).Select
        Selection.PasteSpecial
        Set nowRange = Selection
    Loop
    Call PreCheckSummaryIDVG(VCC)
    
End Sub

Private Sub ArrangeSSVGSheet()
    Dim nowSheet As Worksheet
    Dim srcRange As Range
    Dim nowRange As Range
    
    Dim nowRow As Long
    Dim nowCol As Long
    Dim i As Long, j As Long
    Dim VCC As Double
    Dim VBP As Double
    
    Set nowSheet = ActiveSheet
    
    Set srcRange = nowSheet.Cells(1, nowSheet.Cells(1, nowSheet.UsedRange.Columns.Count).CurrentRegion.Column - 2).CurrentRegion
    srcRange.Copy
    nowSheet.Cells(1, srcRange.Column + srcRange.Columns.Count * 2 + 2).Select
    Selection.PasteSpecial
    Set nowRange = Selection

    For j = 3 To nowRange.Columns.Count
        nowRange.Cells(1, j).Value = Replace(srcRange.Cells(1, j).Value, "ID", "SS")
        For i = 2 To nowRange.Rows.Count - 1
            nowRange.Cells(i, j).Value = "=(LN(ABS(" & srcRange.Cells(i + 1, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "))-LN(ABS(" & srcRange.Cells(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")))/ABS(" & srcRange.Cells(i + 1, 2).Address(RowAbsolute:=False) & "-" & srcRange.Cells(i, 2).Address(RowAbsolute:=False) & ")"
        Next i
    Next j
    nowRange.Rows(i).ClearContents
    VCC = CDbl(getCOL(nowRange.Cells(1, 1).Value, "=", 2))
    Set srcRange = nowRange
    srcRange.Copy
    nowSheet.Cells(srcRange.Row + srcRange.Rows.Count + 1, srcRange.Column).Select
    Selection.PasteSpecial
    Set nowRange = Selection
    
    For j = 3 To nowRange.Columns.Count
        nowRange.Cells(1, j).Value = Replace(srcRange.Cells(1, j).Value, "SS", "drev(1/SS)")
        For i = 2 To nowRange.Rows.Count - 1
            nowRange.Cells(i, j).Value = "=(" & srcRange.Cells(i + 1, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "-" & srcRange.Cells(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/ABS(" & srcRange.Cells(i + 1, 2).Address(RowAbsolute:=False) & "-" & srcRange.Cells(i, 2).Address(RowAbsolute:=False) & ")"
        Next i
    Next j
    
End Sub

Private Sub PreCheckPlot(StressTimes As Integer, mType As Integer)
    ''''''''''''
    ''Type 1: ID-VD
    ''Type 2: ID-VG
    ''Type 3: GM-VG
    ''Type 4: 1/SS-VG / drev(1/SS)
    ''''''''''''

    Dim i As Long, j As Long, k As Long
    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Set nowSheet = ActiveSheet
    Set setSheet = Worksheets("Setting")
    
    
    '==========Auto-Naming==========
    Dim StrDef() As String
    Dim mStrDef
    Dim StrW As String
    Dim StrL As String
    Dim StrBias As String
    
    If Not InStr(LCase(setSheet.Cells(1, 2).Value), "default") > 1 Then
        StrDef = WorksheetSetting(setSheet.Range("B1"))
        ReDim mStrDef(UBound(StrDef)) As String
        For i = 0 To UBound(StrDef)
            mStrDef(i) = SetStr(StrDef(i))
        Next i
        StrW = SetStr("[width]")
        StrL = SetStr("[length]")
        StrBias = SetStr("[bias]")
        
        If mStrDef(0) = "IDVD" Then mStrDef(0) = "ID-VD"
        If mStrDef(0) = "IDVG" Then mStrDef(0) = "ID-VG"
    End If
    '==========Main==========
    Dim nowRange As Range
    Dim nowChartObj As ChartObject
    Dim nowChart As Chart
    
    Dim iStep As Integer
    Dim jStep As Integer
    
    Set nowRange = nowSheet.Cells(1, nowSheet.UsedRange.Columns.Count).CurrentRegion
    
    Do Until nowRange.Cells(1, 1).Value = ""
        Set nowChartObj = nowSheet.ChartObjects.Add(Range(Columns(1), Columns(Range("BoundRange").Columns.Count * 3 - 7)).Width, (nowSheet.ChartObjects.Count) * 300, 432, 300)
        Set nowChart = nowChartObj.Chart
        '==========Chart Plot==========
        With nowChart
            .ChartType = xlXYScatterLinesNoMarkers
            '==========Data Range==========
            If mType = 1 Then
                For i = 1 To nowRange.Columns.Count - 2
                    .SeriesCollection.NewSeries
                    .SeriesCollection(i).XValues = nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))
                    .SeriesCollection(i).Values = nowRange.Range(Cells(2, i + 2), Cells(nowRange.Rows.Count, i + 2))
                    .SeriesCollection(i).Name = nowRange.Cells(1, i + 2)
                Next i
            ElseIf mType = 2 Or mType = 3 Then
                iStep = CInt((nowRange.Columns.Count - 2) / StressTimes)
                jStep = CInt((iStep - 1) / 2)
                k = 1
                For i = 0 To StressTimes - 1
                    For j = 1 To 3
                        .SeriesCollection.NewSeries
                        .SeriesCollection(k).XValues = nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))
                        .SeriesCollection(k).Values = nowRange.Range(Cells(2, iStep * i + jStep * j), Cells(nowRange.Rows.Count, iStep * i + jStep * j))
                        .SeriesCollection(k).Name = nowRange.Cells(1, iStep * i + jStep * j)
                    k = k + 1
                    Next j
                Next i
            ElseIf mType = 4 Then
                iStep = CInt((nowRange.Columns.Count - 2) / StressTimes)
                
                '1/SS-Vg
                .SeriesCollection.NewSeries
                .SeriesCollection(1).XValues = nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))
                .SeriesCollection(1).Values = nowRange.Range(Cells(2, iStep + 2), Cells(nowRange.Rows.Count, iStep + 2))
                .SeriesCollection(1).Name = nowRange.Cells(1, iStep + 2)
                
                'derv(1/SS)-Vg
                Set nowRange = nowSheet.Cells(nowRange.Row + nowRange.Rows.Count + 2, nowRange.Column).CurrentRegion
                .SeriesCollection.NewSeries
                .SeriesCollection(2).XValues = nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))
                .SeriesCollection(2).Values = nowRange.Range(Cells(2, iStep + 2), Cells(nowRange.Rows.Count, iStep + 2))
                .SeriesCollection(2).Name = nowRange.Cells(1, iStep + 2)
                
                .SeriesCollection(2).AxisGroup = 2
                .HasAxis(xlValue, xlPrimary) = True
                .HasAxis(xlValue, xlSecondary) = True
            
            End If
            '==========Line Color==========
            Dim LineColor
            ReDim LineColor(1 To StressTimes)
            Dim GroupNum As Integer
            LineColor(1) = RGB(0, 112, 192)
            LineColor(2) = RGB(0, 176, 80)
            LineColor(3) = RGB(228, 108, 10)
            GroupNum = .SeriesCollection.Count / StressTimes
            If mType = 4 Then
                .SeriesCollection(1).Border.Color = RGB(255, 0, 0)
                .SeriesCollection(2).Border.Color = LineColor(2)
            Else
                For i = 1 To StressTimes
                    For j = 1 To GroupNum
                        .SeriesCollection(j + GroupNum * (i - 1)).Border.Color = LineColor(i)
                    Next j
                Next i
            End If
            '==========Chart Title==========
            Dim oTitle As String
            Dim waferInfo As String
            oTitle = setSheet.Range("B2")
            waferInfo = trim(getCOL(nowRange.Cells(2, 1).Value, ",", 2))
            For j = 0 To UBound(mStrDef)
                oTitle = Replace(oTitle, StrDef(j), mStrDef(j))
            Next j
            oTitle = Replace(oTitle, "[width]", StrW)
            oTitle = Replace(oTitle, "[length]", StrL)
            oTitle = Replace(oTitle, "[bias]", StrBias)
            oTitle = Replace(oTitle, "[wf]", waferInfo)
            oTitle = Replace(oTitle, "[Condition]", nowRange.Cells(1, 1))
            If mType = 3 Then oTitle = Replace(oTitle, "ID", "GM")
            If mType = 4 Then oTitle = Replace(oTitle, "ID", "1/SS")
            .HasTitle = True
            .ChartTitle.Caption = oTitle
            '==========Axit Setting==========
            .Axes(xlCategory).HasTitle = True
            If mType = 1 Then
                .Axes(xlCategory).AxisTitle.Caption = "VD(V)"
            Else
                .Axes(xlCategory).AxisTitle.Caption = "VG(V)"
            End If
            On Error Resume Next
            .Axes(xlCategory).ScaleType = xlLinear
            If nowRange.Cells(2, 2).Value > nowRange.Cells(3, 2).Value Then
                .Axes(xlCategory).MinimumScale = WorksheetFunction.RoundUp(WorksheetFunction.Min(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                .Axes(xlCategory).MaximumScale = WorksheetFunction.RoundDown(WorksheetFunction.Max(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                If mType = 4 Then
                    .HasAxis(xlCategory, xlSecondary) = True
                    .Axes(xlCategory).MaximumScale = -0.5
                    .Axes(xlCategory, xlSecondary).MaximumScale = -0.5
                    .Axes(xlCategory, xlSecondary).MinimumScale = WorksheetFunction.RoundUp(WorksheetFunction.Min(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                End If
            Else
                .Axes(xlCategory).MinimumScale = WorksheetFunction.RoundDown(WorksheetFunction.Min(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                .Axes(xlCategory).MaximumScale = WorksheetFunction.RoundUp(WorksheetFunction.Max(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                If mType = 4 Then
                    .HasAxis(xlCategory, xlSecondary) = True
                    .Axes(xlCategory).MinimumScale = 0.5
                    .Axes(xlCategory, xlSecondary).MinimumScale = 0.5
                    .Axes(xlCategory, xlSecondary).MaximumScale = WorksheetFunction.RoundUp(WorksheetFunction.Max(nowRange.Range(Cells(2, 2), Cells(nowRange.Rows.Count, 2))) * 10, 0) / 10
                End If
            End If
            .Axes(xlCategory).CrossesAt = .Axes(xlCategory).MinimumScale
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "0.0"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlSecondary).HasTitle = True
            
            '.Axes(xlValue).ReversePlotOrder = True
            If nowRange.Cells(2, 2).Value > nowRange.Cells(3, 2).Value Then
                .Axes(xlCategory).ReversePlotOrder = True: .Axes(xlCategory).CrossesAt = .Axes(xlCategory).MaximumScale
                If mType = 4 Then
                    .Axes(xlCategory, xlPrimary).ReversePlotOrder = True: .Axes(xlCategory, xlPrimary).CrossesAt = .Axes(xlCategory).MaximumScale
                    .Axes(xlCategory, xlSecondary).ReversePlotOrder = True: .Axes(xlCategory, xlSecondary).CrossesAt = .Axes(xlCategory).MinimumScale
                End If
                .Axes(xlValue).TickLabelPosition = xlHigh
                .Axes(xlValue).AxisTitle.Left = 7
                .PlotArea.Width = 310.969842519685
                .PlotArea.Left = 22.2050393700787
            End If
            
            
            
            If mType = 3 Then
                .Axes(xlValue).AxisTitle.Caption = "GM(A/V)"
            ElseIf mType = 4 Then
                .Axes(xlValue, xlPrimary).AxisTitle.Caption = "1/SS(dec./V)"
                .Axes(xlValue, xlSecondary).AxisTitle.Caption = "derv(1/SS)"
            Else
                .Axes(xlValue).AxisTitle.Caption = "ID(A)"
            End If
            If mType = 2 Then
                .Axes(xlValue).ScaleType = xlLogarithmic
            Else
                .Axes(xlValue).ScaleType = xlLinear
            End If
            '.Axes(xlValue).MinimumScale
            '.Axes(xlValue).MaximumScale
            .Axes(xlValue).CrossesAt = .Axes(xlValue).MinimumScale
            If mType = 4 Then
                .Axes(xlValue, xlPrimary).TickLabels.NumberFormatLocal = "00"
                .Axes(xlValue, xlSecondary).TickLabels.NumberFormatLocal = "00"
                .Axes(xlValue, xlPrimary).MinimumScale = 0
                .Axes(xlValue, xlPrimary).MaximumScale = 20
                .Axes(xlValue, xlSecondary).MinimumScale = -40
                .Axes(xlValue, xlSecondary).MaximumScale = 40
                
            Else
                .Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
            End If

            '==========Legend==========
            .Legend.Height = .Legend.Height * 1.35
            .Legend.Top = .Legend.Top - 20
            '==========Chart Style==========
            .SetElement (msoElementPrimaryValueGridLinesNone)
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            .SetElement (msoElementChartTitleAboveChart)
        End With
        If mType = 3 Then
            Dim CondNum As Integer
            CondNum = nowSheet.UsedRange.Rows.Count / Range("BoundRange").Rows.Count
            Set nowRange = nowSheet.Cells(nowRange.Row + (nowRange.CurrentRegion.Rows.Count + 2) * CondNum, nowRange.Column).CurrentRegion
        Else
            Set nowRange = nowSheet.Cells(nowRange.Row + nowRange.CurrentRegion.Rows.Count + 1, nowRange.Column).CurrentRegion
        End If
    Loop
        
End Sub

Private Sub PreCheckSummaryIDVD(VCC As Double)

    Dim mRange As Range
    Dim srcRange1 As Range
    Dim srcRange2 As Range
    Dim nowSheet As Worksheet
    Dim i As Integer
    Dim nv2 As Integer
    
    Set nowSheet = ActiveSheet
    Set mRange = nowSheet.Range(Cells(1, Range("BoundRange").Columns.Count * 3 + 2), Cells(28, Range("BoundRange").Columns.Count * 3 + 4))
    Set srcRange1 = nowSheet.Cells(1, Range("BoundRange").Columns.Count * 3 + 6).CurrentRegion
    Set srcRange2 = nowSheet.Cells(srcRange1.Row + srcRange1.Rows.Count + 1, srcRange1.Column).CurrentRegion
    
    For i = 0 To UBound(WaferArray, 2)
        If Not WaferArray(1, i) = "NO" Then
            
            nv2 = (srcRange1.Columns.Count - 2) / 3
            
            'Style
            With mRange
                .Font.Name = "Arial"
                .Font.Size = 12
                .HorizontalAlignment = xlLeft 'xlleft 'xlright 'xlcenter
            End With
            With Union(mRange.Cells(1, 2), _
                       mRange.Cells(3, 1), mRange.Cells(3, 2), mRange.Cells(3, 3), mRange.Cells(4, 1), mRange.Cells(5, 1), mRange.Cells(6, 1), _
                       mRange.Cells(12, 1), mRange.Cells(12, 2), mRange.Cells(12, 3), mRange.Cells(13, 1), mRange.Cells(14, 1), mRange.Cells(15, 1), _
                       mRange.Cells(21, 1), mRange.Cells(21, 2), mRange.Cells(21, 3), mRange.Cells(22, 1), mRange(23, 1), mRange.Cells(24, 1), _
                       mRange.Cells(26, 1), mRange.Cells(26, 2), mRange.Cells(27, 1), mRange.Cells(28, 1))
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
            With Union(mRange.Cells(7, 1), mRange.Cells(7, 2), mRange.Cells(7, 3), mRange.Cells(8, 1), mRange.Cells(9, 1), mRange.Cells(10, 1), _
                       mRange.Cells(16, 1), mRange.Cells(16, 2), mRange.Cells(16, 3), mRange.Cells(17, 1), mRange.Cells(18, 1), mRange.Cells(19, 1))
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            End With
            
            
            'Value
            mRange.Cells(1, 1).Value = WaferArray(0, i)
            mRange.Cells(1, 2).Value = "Vcc"
            mRange.Cells(1, 3).Value = VCC
            
            mRange.Cells(3, 1).Value = srcRange1.Cells(1, 1)
            mRange.Cells(4, 1).Value = "IDL-1"
            mRange.Cells(5, 1).Value = "IDL-2"
            mRange.Cells(6, 1).Value = "IDL-3"
            mRange.Cells(3, 2).Value = "Value"
            mRange.Cells(4, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 1).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(5, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 2).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(6, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 3).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(3, 3).Value = "Shift%"
            mRange.Cells(4, 3).Value = ""
            mRange.Cells(5, 3).Value = "=" & mRange.Cells(5, 2).Address & "/" & mRange.Cells(4, 2).Address & "-1"
            mRange.Cells(6, 3).Value = "=" & mRange.Cells(6, 2).Address & "/" & mRange.Cells(4, 2).Address & "-1"
            Range(mRange.Cells(4, 2), mRange.Cells(6, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(4, 3), mRange.Cells(6, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(7, 1).Value = srcRange2.Cells(1, 1)
            mRange.Cells(8, 1).Value = "IDL-1"
            mRange.Cells(9, 1).Value = "IDL-2"
            mRange.Cells(10, 1).Value = "IDL-3"
            mRange.Cells(7, 2).Value = "Value"
            mRange.Cells(8, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 1).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(9, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 2).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(10, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 3).Address & ",MATCH(" & CDbl(VCC / Abs(VCC) * 0.05) & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(7, 3).Value = "Shift%"
            mRange.Cells(8, 3).Value = ""
            mRange.Cells(9, 3).Value = "=" & mRange.Cells(9, 2).Address & "/" & mRange.Cells(8, 2).Address & "-1"
            mRange.Cells(10, 3).Value = "=" & mRange.Cells(10, 2).Address & "/" & mRange.Cells(8, 2).Address & "-1"
            Range(mRange.Cells(8, 2), mRange.Cells(10, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(8, 3), mRange.Cells(10, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(12, 1).Value = srcRange1.Cells(1, 1)
            mRange.Cells(13, 1).Value = "IDS-1"
            mRange.Cells(14, 1).Value = "IDS-2"
            mRange.Cells(15, 1).Value = "IDS-3"
            mRange.Cells(12, 2).Value = "Value"
            mRange.Cells(13, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(14, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(15, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 3).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(12, 3).Value = "Shift%"
            mRange.Cells(13, 3).Value = ""
            mRange.Cells(14, 3).Value = "=" & mRange.Cells(14, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            mRange.Cells(15, 3).Value = "=" & mRange.Cells(15, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            Range(mRange.Cells(13, 2), mRange.Cells(15, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(13, 3), mRange.Cells(15, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(16, 1).Value = srcRange2.Cells(1, 1)
            mRange.Cells(17, 1).Value = "IDS-1"
            mRange.Cells(18, 1).Value = "IDS-2"
            mRange.Cells(19, 1).Value = "IDS-3"
            mRange.Cells(16, 2).Value = "Value"
            mRange.Cells(17, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(18, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(19, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 3).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(16, 3).Value = "Shift%"
            mRange.Cells(17, 3).Value = ""
            mRange.Cells(18, 3).Value = "=" & mRange.Cells(18, 2).Address & "/" & mRange.Cells(17, 2).Address & "-1"
            mRange.Cells(19, 3).Value = "=" & mRange.Cells(19, 2).Address & "/" & mRange.Cells(17, 2).Address & "-1"
            Range(mRange.Cells(17, 2), mRange.Cells(19, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(17, 3), mRange.Cells(19, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(21, 1).Value = srcRange1.Cells(1, 1)
            mRange.Cells(22, 1).Value = "ID1max"
            mRange.Cells(23, 1).Value = "ID2max"
            mRange.Cells(24, 1).Value = "ID3max"
            mRange.Cells(21, 2).Value = "Value"
            mRange.Cells(22, 2).Value = "=MAX(" & srcRange1.Columns(2 + nv2 * 1).Offset(1, 0).Address & ")"
            mRange.Cells(23, 2).Value = "=MAX(" & srcRange1.Columns(2 + nv2 * 2).Offset(1, 0).Address & ")"
            mRange.Cells(24, 2).Value = "=MAX(" & srcRange1.Columns(2 + nv2 * 3).Offset(1, 0).Address & ")"
            mRange.Cells(21, 3).Value = "Shift%"
            mRange.Cells(22, 3).Value = ""
            mRange.Cells(23, 3).Value = "=" & mRange.Cells(14, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            mRange.Cells(24, 3).Value = "=" & mRange.Cells(15, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            Range(mRange.Cells(22, 2), mRange.Cells(24, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(22, 3), mRange.Cells(24, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(26, 1).Value = srcRange1.Cells(1, 1)
            mRange.Cells(27, 1).Value = "AÂI"
            mRange.Cells(28, 1).Value = "BÂI"
            mRange.Cells(26, 2).Value = "Value"
            mRange.Cells(27, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 3).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(28, 2).Value = "=INDEX(" & srcRange1.Columns(4 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            Range(mRange.Cells(27, 2), mRange.Cells(28, 2)).NumberFormat = "0.00E+00"
            
            Set mRange = nowSheet.Range(Cells(mRange.Row + mRange.Rows.Count + 1, Range("BoundRange").Columns.Count * 3 + 2), Cells(mRange.Row + 2 * mRange.Rows.Count, Range("BoundRange").Columns.Count * 3 + 4))
            Set srcRange1 = nowSheet.Cells(srcRange2.Row + srcRange2.Rows.Count + 1, srcRange2.Column).CurrentRegion
            Set srcRange2 = nowSheet.Cells(srcRange1.Row + srcRange1.Rows.Count + 1, srcRange1.Column).CurrentRegion
        End If
    Next i
    
End Sub

Private Sub PreCheckSummaryIDVG(VCC As Double)

    Dim mRange As Range
    Dim srcRange1 As Range
    Dim srcRange2 As Range
    Dim srcRange3 As Range
    Dim nowSheet As Worksheet
    Dim i As Integer
    Dim nv2 As Integer
    
    Set nowSheet = ActiveSheet
    Set mRange = nowSheet.Range(Cells(1, Range("BoundRange").Columns.Count * 3 + 2), Cells(28, Range("BoundRange").Columns.Count * 3 + 4))
    Set srcRange1 = nowSheet.Cells(1, Range("BoundRange").Columns.Count * 3 + 6).CurrentRegion
    Set srcRange2 = nowSheet.Cells(srcRange1.Row + srcRange1.Rows.Count + 1, srcRange1.Column).CurrentRegion
    Set srcRange3 = nowSheet.Cells(srcRange1.Row, srcRange1.Column + srcRange1.Columns.Count + 1).CurrentRegion
    
    For i = 0 To UBound(WaferArray, 2)
        If Not WaferArray(1, i) = "NO" Then
            
            nv2 = (srcRange1.Columns.Count - 2) / 3
            
            'Style
            With mRange
                .Font.Name = "Arial"
                .Font.Size = 12
                .HorizontalAlignment = xlLeft 'xlleft 'xlright 'xlcenter
            End With
            With Union(mRange.Cells(1, 2), _
                       mRange.Cells(3, 1), mRange.Cells(3, 2), mRange.Cells(3, 3), mRange.Cells(4, 1), mRange.Cells(5, 1), mRange.Cells(6, 1), _
                       mRange.Cells(12, 1), mRange.Cells(12, 2), mRange.Cells(12, 3), mRange.Cells(13, 1), mRange.Cells(14, 1), mRange.Cells(15, 1), _
                       mRange.Cells(21, 1), mRange.Cells(21, 2), mRange.Cells(21, 3), mRange.Cells(22, 1), mRange.Cells(23, 1), mRange.Cells(24, 1), _
                       mRange.Cells(26, 1), mRange.Cells(26, 2), mRange.Cells(27, 1), mRange.Cells(28, 1))
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
            With Union(mRange.Cells(7, 1), mRange.Cells(7, 2), mRange.Cells(7, 3), mRange.Cells(8, 1), mRange.Cells(9, 1), mRange.Cells(10, 1), _
                       mRange.Cells(16, 1), mRange.Cells(16, 2), mRange.Cells(16, 3), mRange.Cells(17, 1), mRange.Cells(18, 1), mRange.Cells(19, 1))
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            End With
            
            
            'Value
            mRange.Cells(1, 1).Value = WaferArray(0, i)
            mRange.Cells(1, 2).Value = "Vcc"
            mRange.Cells(1, 3).Value = VCC
            
            mRange.Cells(3, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 3), "(", 2), ")", 1))
            mRange.Cells(4, 1).Value = "IDL-1"
            mRange.Cells(5, 1).Value = "IDL-2"
            mRange.Cells(6, 1).Value = "IDL-3"
            mRange.Cells(3, 2).Value = "Value"
            mRange.Cells(4, 2).Value = "=INDEX(" & srcRange1.Columns(3 + nv2 * 0).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(5, 2).Value = "=INDEX(" & srcRange1.Columns(3 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(6, 2).Value = "=INDEX(" & srcRange1.Columns(3 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(3, 3).Value = "Shift%"
            mRange.Cells(4, 3).Value = ""
            mRange.Cells(5, 3).Value = "=" & mRange.Cells(5, 2).Address & "/" & mRange.Cells(4, 2).Address & "-1"
            mRange.Cells(6, 3).Value = "=" & mRange.Cells(6, 2).Address & "/" & mRange.Cells(4, 2).Address & "-1"
            Range(mRange.Cells(4, 2), mRange.Cells(6, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(4, 3), mRange.Cells(6, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(7, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 2 + nv2), "(", 2), ")", 1))
            mRange.Cells(8, 1).Value = "IDL-1"
            mRange.Cells(9, 1).Value = "IDL-2"
            mRange.Cells(10, 1).Value = "IDL-3"
            mRange.Cells(7, 2).Value = "Value"
            mRange.Cells(8, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(9, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(10, 2).Value = "=INDEX(" & srcRange1.Columns(2 + nv2 * 3).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange1.Columns(2).Address & ",0))"
            mRange.Cells(7, 3).Value = "Shift%"
            mRange.Cells(8, 3).Value = ""
            mRange.Cells(9, 3).Value = "=" & mRange.Cells(9, 2).Address & "/" & mRange.Cells(8, 2).Address & "-1"
            mRange.Cells(10, 3).Value = "=" & mRange.Cells(10, 2).Address & "/" & mRange.Cells(8, 2).Address & "-1"
            Range(mRange.Cells(8, 2), mRange.Cells(10, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(8, 3), mRange.Cells(10, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(12, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 3), "(", 2), ")", 1))
            mRange.Cells(13, 1).Value = "IDS-1"
            mRange.Cells(14, 1).Value = "IDS-2"
            mRange.Cells(15, 1).Value = "IDS-3"
            mRange.Cells(12, 2).Value = "Value"
            mRange.Cells(13, 2).Value = "=INDEX(" & srcRange2.Columns(3 + nv2 * 0).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(14, 2).Value = "=INDEX(" & srcRange2.Columns(3 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(15, 2).Value = "=INDEX(" & srcRange2.Columns(3 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(12, 3).Value = "Shift%"
            mRange.Cells(13, 3).Value = ""
            mRange.Cells(14, 3).Value = "=" & mRange.Cells(14, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            mRange.Cells(15, 3).Value = "=" & mRange.Cells(15, 2).Address & "/" & mRange.Cells(13, 2).Address & "-1"
            Range(mRange.Cells(13, 2), mRange.Cells(15, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(13, 3), mRange.Cells(15, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(16, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 2 + nv2), "(", 2), ")", 1))
            mRange.Cells(17, 1).Value = "IDS-1"
            mRange.Cells(18, 1).Value = "IDS-2"
            mRange.Cells(19, 1).Value = "IDS-3"
            mRange.Cells(16, 2).Value = "Value"
            mRange.Cells(17, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 1).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(18, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(19, 2).Value = "=INDEX(" & srcRange2.Columns(2 + nv2 * 3).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            mRange.Cells(16, 3).Value = "Shift%"
            mRange.Cells(17, 3).Value = ""
            mRange.Cells(18, 3).Value = "=" & mRange.Cells(18, 2).Address & "/" & mRange.Cells(17, 2).Address & "-1"
            mRange.Cells(19, 3).Value = "=" & mRange.Cells(19, 2).Address & "/" & mRange.Cells(17, 2).Address & "-1"
            Range(mRange.Cells(17, 2), mRange.Cells(19, 2)).NumberFormat = "0.00E+00"
            Range(mRange.Cells(17, 3), mRange.Cells(19, 3)).NumberFormat = "0.00%"
            
            mRange.Cells(21, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 3), "(", 2), ")", 1))
            mRange.Cells(22, 1).Value = "Vtgm1"
            mRange.Cells(23, 1).Value = "Vtgm2"
            mRange.Cells(24, 1).Value = "Vtgm3"
            mRange.Cells(21, 2).Value = "Value"
            mRange.Cells(22, 2).Value = "=(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 0).Address & "*INDEX(" & srcRange3.Columns(2).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 0).Address & "," & srcRange3.Columns(3 + nv2 * 0).Address & ",0))-" & "INDEX(" & srcRange1.Columns(3 + nv2 * 0).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 0).Address & "," & srcRange3.Columns(3 + nv2 * 0).Address & ",0))" & ")/" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 0).Address & "-0.5*0.05"
            mRange.Cells(23, 2).Value = "=(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 1).Address & "*INDEX(" & srcRange3.Columns(2).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 1).Address & "," & srcRange3.Columns(3 + nv2 * 1).Address & ",0))-" & "INDEX(" & srcRange1.Columns(3 + nv2 * 1).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 1).Address & "," & srcRange3.Columns(3 + nv2 * 1).Address & ",0))" & ")/" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 1).Address & "-0.5*0.05"
            mRange.Cells(24, 2).Value = "=(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 2).Address & "*INDEX(" & srcRange3.Columns(2).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 2).Address & "," & srcRange3.Columns(3 + nv2 * 2).Address & ",0))-" & "INDEX(" & srcRange1.Columns(3 + nv2 * 2).Address & ",MATCH(" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 2).Address & "," & srcRange3.Columns(3 + nv2 * 2).Address & ",0))" & ")/" & srcRange3.Cells(srcRange3.Rows.Count, 3 + nv2 * 2).Address & "-0.5*0.05"
            mRange.Cells(21, 3).Value = "Shift(mV)"
            mRange.Cells(22, 3).Value = ""
            mRange.Cells(23, 3).Value = "=(" & mRange.Cells(23, 2).Address & "-" & mRange.Cells(22, 2).Address & ")*1000"
            mRange.Cells(24, 3).Value = "=(" & mRange.Cells(24, 2).Address & "-" & mRange.Cells(22, 2).Address & ")*1000"
            Range(mRange.Cells(22, 2), mRange.Cells(24, 3)).NumberFormat = "0.000"
            
            mRange.Cells(26, 1).Value = UCase(getCOL(getCOL(srcRange1.Cells(1, 3), "(", 2), ")", 1))
            mRange.Cells(27, 1).Value = "AÂI"
            mRange.Cells(28, 1).Value = "BÂI"
            mRange.Cells(26, 2).Value = "Value"
            mRange.Cells(27, 2).Value = "=INDEX(" & srcRange2.Columns(3 + nv2 * 2).Address & ",MATCH(" & mRange.Cells(1, 3).Address & "," & srcRange2.Columns(2).Address & ",0))"
            Dim VBP As Double
            VBP = CDbl(getCOL(getCOL(Worksheets(Replace(nowSheet.Name, "VG", "VD")).Cells(1, Worksheets(Replace(nowSheet.Name, "VG", "VD")).UsedRange.Columns.Count).CurrentRegion.Cells(1, 4).Value, "=", 2), ")", 1))
            mRange.Cells(28, 2).Value = "=INDEX(" & srcRange2.Columns(3 + nv2 * 2).Address & ",MATCH(" & VBP & "," & srcRange2.Columns(2).Address & ",0))"
            Range(mRange.Cells(27, 2), mRange.Cells(28, 2)).NumberFormat = "0.00E+00"
        
            Set mRange = nowSheet.Range(Cells(mRange.Row + mRange.Rows.Count + 1, Range("BoundRange").Columns.Count * 3 + 2), Cells(mRange.Row + 2 * mRange.Rows.Count, Range("BoundRange").Columns.Count * 3 + 4))
            Set srcRange1 = nowSheet.Cells(srcRange2.Row + srcRange2.Rows.Count + 1, srcRange2.Column).CurrentRegion
            Set srcRange2 = nowSheet.Cells(srcRange1.Row + srcRange1.Rows.Count + 1, srcRange1.Column).CurrentRegion
            Set srcRange3 = nowSheet.Cells(srcRange1.Row, srcRange1.Column + srcRange1.Columns.Count + 1).CurrentRegion
        End If
    Next i

End Sub

Public Sub PreCheckSummary()
    
    Dim nowSheet As Worksheet
    Dim srcSheet As Worksheet
    Dim nowRow As Long
    Dim nowCol As Long
    Dim i As Integer, j As Integer
    Dim waferNum As Integer
    Dim DS
    Dim DE
    Dim itemAry
    Dim sColl As New Collection
    Dim sAry
    Dim nomiAry()
    Dim Item
    
    On Error Resume Next
    
    If Not IsExistSheet("Precheck") Then Exit Sub
    
    Set srcSheet = Worksheets("Precheck")
    nowRow = 1: nowCol = 1
    Do
        If UCase(srcSheet.Cells(nowRow, nowCol).Value) = "IDVD" Then
            nowRow = nowRow + 2
            nowCol = nowCol + 1
            DS = srcSheet.Cells(nowRow, nowCol).Address
            Do Until srcSheet.Cells(nowRow, nowCol).Value = ""
                nowRow = nowRow + 1
            Loop
            DE = srcSheet.Cells(nowRow - 1, nowCol).Address
            Exit Do
        End If
    Loop Until nowRow >= srcSheet.UsedRange.Rows.Count
    ReDim itemAry(1 To srcSheet.Range(DS & ":" & DE).Count)
    '==========Load Data List==========
    For Each Item In srcSheet.Range(DS & ":" & DE)
        i = i + 1
        itemAry(i) = Item
        sColl.Add getCOL(Item, "_", 4), getCOL(Item, "_", 4)
    Next Item
    Err.Clear
    On Error GoTo 0
    '==========Decide device no.==========
    ReDim nomiAry(1 To sColl.Count, 1 To 3)
    For i = 1 To sColl.Count
        nomiAry(i, 1) = "PARAM" 'Params
        nomiAry(i, 2) = 0       'W
        nomiAry(i, 3) = 100     'L
    Next i
    '==========Decide nominal device==========
    Item = itemAry((UBound(itemAry) / 2 + LBound(itemAry)) / 2)
    nomiAry(1, 1) = Item
    nomiAry(1, 2) = CDbl(Replace(getCOL(Item, "_", 2), "p", "."))
    nomiAry(1, 3) = CDbl(Replace(getCOL(Item, "_", 3), "p", "."))
    
    Item = itemAry((UBound(itemAry) / 2 + LBound(itemAry)) / 2 + UBound(itemAry) / 2)
    nomiAry(2, 1) = Item
    nomiAry(2, 2) = CDbl(Replace(getCOL(Item, "_", 2), "p", "."))
    nomiAry(2, 3) = CDbl(Replace(getCOL(Item, "_", 3), "p", "."))
    
    'For Each item In itemAry
    '    For i = 1 To sColl.Count
    '        If InStr(item, sColl.item(i)) Then
    '            If CDbl(Replace(getCOL(item, "_", 2), "p", ".")) >= nomiAry(i, 2) And CDbl(Replace(getCOL(item, "_", 3), "p", ".")) <= nomiAry(i, 3) Then
    '                nomiAry(i, 1) = item
    '                nomiAry(i, 2) = CDbl(Replace(getCOL(item, "_", 2), "p", "."))
    '                nomiAry(i, 3) = CDbl(Replace(getCOL(item, "_", 3), "p", "."))
    '            End If
    '        End If
    '    Next i
    'Next item
    '==========Print Summary=========='
    For i = 1 To sColl.Count
        sAry = Filter(itemAry, sColl.Item(i))
        If i > 1 Then
            Set nowSheet = AddSheet("PrecheckSummary_" & sColl.Item(i), True, "PrecheckSummary_" & sColl.Item(i - 1))
        Else
            Set nowSheet = AddSheet("PrecheckSummary_" & sColl.Item(i), True, "Precheck")
        End If
        nowRow = 2
        nowCol = 2
    For waferNum = 0 To UBound(WaferArray, 2)
    If Not WaferArray(1, waferNum) = "NO" Then
        With nowSheet
            .Cells(nowRow, nowCol - 1) = WaferArray(0, waferNum)
            .Cells.Font.Name = "Arial"
            .Cells.HorizontalAlignment = xlCenter
            .Cells.VerticalAlignment = xlCenter
            '==========DC Instability Check==========
            .Cells(nowRow, nowCol) = "DC Instability Check"
            .Cells(nowRow, nowCol + 1) = "Vt shift < 30mV; Idlin shift < 10% and Idsat < 5%"
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).Merge
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).HorizontalAlignment = xlLeft
            .Cells(nowRow, nowCol).Interior.Color = RGB(255, 255, 0)
            .Cells(nowRow, nowCol).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol).Font.Bold = True
            .Cells(nowRow, nowCol + 1).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol + 1).Font.Bold = True
            With Range(.Cells(nowRow, nowCol), .Cells(nowRow, nowCol + 11))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                '.Borders(xlInsideVertical).LineStyle = xlContinuous
                '.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                '.Borders(xlInsideVertical).Weight = xlThin
                '.Borders(xlInsideHorizontal).Weight = xlMedium
            End With
            For j = 0 To 2
                .Cells(nowRow + 1, nowCol + 3 + 3 * j) = "1st"
                .Cells(nowRow + 1, nowCol + 4 + 3 * j) = "2nd"
                .Cells(nowRow + 1, nowCol + 5 + 3 * j) = "3rd"
                With Range(.Cells(nowRow + 1, nowCol + 3 + 3 * j), .Cells(nowRow + 1, nowCol + 5 + 3 * j))
                    .Interior.Color = RGB(250, 191, 143)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
            Next j
            .Cells(nowRow + 2, nowCol) = "Item"
            .Cells(nowRow + 2, nowCol + 1) = "W"
            .Cells(nowRow + 2, nowCol + 2) = "L"
            .Cells(nowRow + 2, nowCol + 3) = "Vth (V)"
            .Cells(nowRow + 2, nowCol + 6) = "IDL (uA/um)"
            .Cells(nowRow + 2, nowCol + 9) = "IDS (uA/um)"
            With Range(.Cells(nowRow + 2, nowCol), .Cells(nowRow + 2, nowCol + 2))
                .Interior.Color = RGB(196, 215, 155)
                
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideHorizontal).Weight = xlThin
            End With
            With Range(.Cells(nowRow + 2, nowCol + 3), .Cells(nowRow + 2, nowCol + 11))
                .Interior.Color = RGB(253, 233, 217)
                
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlMedium
                .Borders(xlInsideHorizontal).Weight = xlMedium
            End With
            Range(.Cells(nowRow + 2, nowCol + 3), .Cells(nowRow + 2, nowCol + 5)).Merge
            Range(.Cells(nowRow + 2, nowCol + 6), .Cells(nowRow + 2, nowCol + 8)).Merge
            Range(.Cells(nowRow + 2, nowCol + 9), .Cells(nowRow + 2, nowCol + 11)).Merge
            nowRow = nowRow + 3
            For Each Item In sAry
                .Cells(nowRow, nowCol) = Item
                .Cells(nowRow, nowCol + 1) = CDbl(Replace(getCOL(Item, "_", 2), "p", "."))
                .Cells(nowRow, nowCol + 2) = CDbl(Replace(getCOL(Item, "_", 3), "p", "."))
                With Range(.Cells(nowRow, nowCol), .Cells(nowRow + 1, nowCol + 2))
                    .Interior.Color = RGB(216, 228, 188)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                .Cells(nowRow + 1, nowCol) = "Diff.(mV; %; %)"
                Range(.Cells(nowRow + 1, nowCol), .Cells(nowRow + 1, nowCol + 2)).Interior.Color = RGB(235, 241, 222)
                .Rows(nowRow + 1).Font.Size = 10
                For j = 0 To 2
                    With Range(.Cells(nowRow, nowCol + 3 + 3 * j), .Cells(nowRow + 1, nowCol + 5 + 3 * j))
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                
                        .Borders(xlEdgeTop).Weight = xlMedium
                        .Borders(xlEdgeBottom).Weight = xlMedium
                        .Borders(xlEdgeLeft).Weight = xlMedium
                        .Borders(xlEdgeRight).Weight = xlMedium
                        .Borders(xlInsideVertical).Weight = xlThin
                        .Borders(xlInsideHorizontal).Weight = xlThin
                    End With
                    Range(.Cells(nowRow + 1, nowCol + 3 + 3 * j), .Cells(nowRow + 1, nowCol + 5 + 3 * j)).Interior.Color = RGB(255, 255, 153)
                Next j
                '==========Derive Data From Other Sheets==========
                Dim dSheet As Worksheet
                Dim gSheet As Worksheet
                Dim tmpAds As String
                Set dSheet = Worksheets(.Cells(nowRow, nowCol).Value)
                Set gSheet = Worksheets(Replace(.Cells(nowRow, nowCol).Value, "VD", "VG"))

                .Cells(nowRow, nowCol + 3).Value = "=" & getKA(Range(getKA(gSheet.Cells(1, gSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "Vtgm1", 1)
                .Cells(nowRow, nowCol + 4).Value = "=" & getKA(Range(getKA(gSheet.Cells(1, gSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "Vtgm2", 1)
                .Cells(nowRow, nowCol + 5).Value = "=" & getKA(Range(getKA(gSheet.Cells(1, gSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "Vtgm3", 1)
                Range(.Cells(nowRow, nowCol + 3), .Cells(nowRow, nowCol + 5)).NumberFormat = "0.000"
                .Cells(nowRow, nowCol + 6).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDL-1", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                .Cells(nowRow, nowCol + 7).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDL-2", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                .Cells(nowRow, nowCol + 8).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDL-3", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                Range(.Cells(nowRow, nowCol + 6), .Cells(nowRow, nowCol + 8)).NumberFormat = "0.0"
                .Cells(nowRow, nowCol + 9).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-1", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                .Cells(nowRow, nowCol + 10).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-2", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                .Cells(nowRow, nowCol + 11).Value = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-3", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
                Range(.Cells(nowRow, nowCol + 9), .Cells(nowRow, nowCol + 11)).NumberFormat = "0.0"
                .Cells(nowRow + 1, nowCol + 4).Value = "=(" & .Cells(nowRow, nowCol + 4).Address & "-" & .Cells(nowRow, nowCol + 3).Address & ")*1000"
                .Cells(nowRow + 1, nowCol + 5).Value = "=(" & .Cells(nowRow, nowCol + 5).Address & "-" & .Cells(nowRow, nowCol + 3).Address & ")*1000"
                Range(.Cells(nowRow + 1, nowCol + 3), .Cells(nowRow + 1, nowCol + 5)).NumberFormat = "0.0"
                .Cells(nowRow + 1, nowCol + 7).Value = "=" & .Cells(nowRow, nowCol + 7).Address & "/" & .Cells(nowRow, nowCol + 6).Address & "-1"
                .Cells(nowRow + 1, nowCol + 8).Value = "=" & .Cells(nowRow, nowCol + 8).Address & "/" & .Cells(nowRow, nowCol + 6).Address & "-1"
                Range(.Cells(nowRow + 1, nowCol + 6), .Cells(nowRow + 1, nowCol + 8)).NumberFormat = "0.0%"
                .Cells(nowRow + 1, nowCol + 10).Value = "=" & .Cells(nowRow, nowCol + 10).Address & "/" & .Cells(nowRow, nowCol + 9).Address & "-1"
                .Cells(nowRow + 1, nowCol + 11).Value = "=" & .Cells(nowRow, nowCol + 11).Address & "/" & .Cells(nowRow, nowCol + 9).Address & "-1"
                Range(.Cells(nowRow + 1, nowCol + 9), .Cells(nowRow + 1, nowCol + 11)).NumberFormat = "0.0%"
                '========================================
                nowRow = nowRow + 2
            Next Item
            nowRow = nowRow + 1
            '==========Vb Stress Precheck==========
            .Cells(nowRow, nowCol) = "Vb Stress Precheck"
            .Cells(nowRow, nowCol + 3) = "Check Id differences of Id_Vg & Id_Vd @ Vg=Vth+0.1V and Vd=Vcc->Vmax (after Id_Vg sweep with Vb=-Vmax) for both N/P nominal device. Diff% < 10%"
            .Rows(nowRow).RowHeight = 33
            .Cells(nowRow, nowCol + 3).WrapText = True
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).Merge
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).HorizontalAlignment = xlLeft
            .Cells(nowRow, nowCol).Interior.Color = RGB(255, 255, 0)
            .Cells(nowRow, nowCol).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol).Font.Bold = True
            .Cells(nowRow, nowCol + 1).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol + 1).Font.Bold = True
            With Range(.Cells(nowRow, nowCol), .Cells(nowRow, nowCol + 11))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                '.Borders(xlInsideVertical).LineStyle = xlContinuous
                '.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                '.Borders(xlInsideVertical).Weight = xlThin
                '.Borders(xlInsideHorizontal).Weight = xlMedium
            End With
            .Cells(nowRow + 1, nowCol + 4) = "(uA/um)"
            .Cells(nowRow + 1, nowCol + 6) = "(pA/um)"
            For j = 0 To 1
                With Range(.Cells(nowRow + 1, nowCol + 3 + 2 * j), .Cells(nowRow + 1, nowCol + 4 + 2 * j))
                    .Interior.Color = RGB(250, 191, 143)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                With Range(.Cells(nowRow + 2, nowCol + 3 + 2 * j), .Cells(nowRow + 3, nowCol + 4 + 2 * j))
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                With Range(.Cells(nowRow + 4, nowCol + 3 + 2 * j), .Cells(nowRow + 4, nowCol + 4 + 2 * j))
                    .Interior.Color = RGB(255, 255, 153)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
            Next j
            nowRow = nowRow + 2
            .Cells(nowRow, nowCol) = nomiAry(i, 1)
            Range(.Cells(nowRow, nowCol), .Cells(nowRow + 1, nowCol)).Merge
            .Cells(nowRow, nowCol + 1) = nomiAry(i, 2)
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow + 1, nowCol + 1)).Merge
            .Cells(nowRow, nowCol + 2) = nomiAry(i, 3)
            Range(.Cells(nowRow, nowCol + 2), .Cells(nowRow + 1, nowCol + 2)).Merge
            With Range(.Cells(nowRow, nowCol), .Cells(nowRow + 2, nowCol + 2))
                .Interior.Color = RGB(216, 228, 188)
                
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideHorizontal).Weight = xlThin
            End With
            .Cells(nowRow + 2, nowCol) = "Diff.%)"
            Range(.Cells(nowRow + 2, nowCol), .Cells(nowRow + 2, nowCol + 2)).Interior.Color = RGB(235, 241, 222)
            .Cells(nowRow + 0, nowCol + 3) = "A(Id-Vg)"
            .Cells(nowRow + 1, nowCol + 3) = "A'(Id-Vd)"
            .Cells(nowRow + 0, nowCol + 5) = "B(Id-Vg)"
            .Cells(nowRow + 1, nowCol + 5) = "B'(Id-Vd)"
            
            
            Set dSheet = Worksheets(Cells(nowRow, 2).Value)
            Set gSheet = Worksheets(Replace(Cells(nowRow, 2).Value, "IDVD", "IDVG"))
            
            .Cells(nowRow + 0, nowCol + 4) = "=" & getKA(Range(getKA(gSheet.Cells(1, gSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "AÂI", 1) & "*1e6/" & .Cells(nowRow, nowCol + 1)
            .Cells(nowRow + 1, nowCol + 4) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "AÂI", 1) & "*1e6/" & .Cells(nowRow, nowCol + 1)
            .Cells(nowRow + 2, nowCol + 4) = "=" & .Cells(nowRow + 1, nowCol + 4).Address & "/" & .Cells(nowRow + 0, nowCol + 4).Address & "-1"
            .Cells(nowRow + 0, nowCol + 6) = "=" & getKA(Range(getKA(gSheet.Cells(1, gSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "BÂI", 1) & "*1e12/" & .Cells(nowRow, nowCol + 1)
            .Cells(nowRow + 1, nowCol + 6) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "BÂI", 1) & "*1e12/" & .Cells(nowRow, nowCol + 1)
            .Cells(nowRow + 2, nowCol + 6) = "=" & .Cells(nowRow + 1, nowCol + 6).Address & "/" & .Cells(nowRow + 0, nowCol + 6).Address & "-1"
            
            Union(Range(.Cells(nowRow, nowCol + 4), .Cells(nowRow + 1, nowCol + 4)), Range(.Cells(nowRow, nowCol + 6), .Cells(nowRow + 1, nowCol + 6))).NumberFormat = "0.0"
            Range(.Cells(nowRow + 2, nowCol + 4), .Cells(nowRow + 2, nowCol + 6)).NumberFormat = "0.0%"
            
            With Union(Range(.Cells(nowRow, nowCol + 3), Cells(nowRow + 1, nowCol + 3)), Range(.Cells(nowRow, nowCol + 5), Cells(nowRow + 1, nowCol + 5)))
                .Font.Size = 10
                .Interior.Color = RGB(253, 233, 217)
            End With
            .Cells(nowRow + 2, nowCol).Font.Size = 10
            
            nowRow = nowRow + 4
            '==========Kirk Effect Check==========
            .Cells(nowRow, nowCol) = "Kirk Effect Check"
            .Cells(nowRow, nowCol + 1) = "Nominal device: < 2%"
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).Merge
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow, nowCol + 11)).HorizontalAlignment = xlLeft
                        .Cells(nowRow, nowCol).Interior.Color = RGB(255, 255, 0)
            .Cells(nowRow, nowCol).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol).Font.Bold = True
            .Cells(nowRow, nowCol + 1).Font.Color = RGB(0, 0, 255)
            .Cells(nowRow, nowCol + 1).Font.Bold = True
            With Range(.Cells(nowRow, nowCol), .Cells(nowRow, nowCol + 11))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                '.Borders(xlInsideVertical).LineStyle = xlContinuous
                '.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                '.Borders(xlInsideVertical).Weight = xlThin
                '.Borders(xlInsideHorizontal).Weight = xlMedium
            End With
            For j = 0 To 0
                .Cells(nowRow + 1, nowCol + 4 + 4 * j) = "1st"
                .Cells(nowRow + 1, nowCol + 5 + 4 * j) = "2nd"
                .Cells(nowRow + 1, nowCol + 6 + 4 * j) = "3rd"
                With Range(.Cells(nowRow + 1, nowCol + 3 + 4 * j), .Cells(nowRow + 1, nowCol + 6 + 4 * j))
                    .Interior.Color = RGB(250, 191, 143)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                With Range(.Cells(nowRow + 2, nowCol + 3 + 4 * j), .Cells(nowRow + 3, nowCol + 6 + 4 * j))
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                With Range(.Cells(nowRow + 4, nowCol + 3 + 4 * j), .Cells(nowRow + 4, nowCol + 6 + 4 * j))
                    .Interior.Color = RGB(255, 255, 153)
                    
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeLeft).Weight = xlMedium
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlInsideVertical).Weight = xlThin
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
                
            Next j
            nowRow = nowRow + 2
            .Cells(nowRow, nowCol) = nomiAry(i, 1)
            Range(.Cells(nowRow, nowCol), .Cells(nowRow + 1, nowCol)).Merge
            .Cells(nowRow, nowCol + 1) = nomiAry(i, 2)
            Range(.Cells(nowRow, nowCol + 1), .Cells(nowRow + 1, nowCol + 1)).Merge
            .Cells(nowRow, nowCol + 2) = nomiAry(i, 3)
            Range(.Cells(nowRow, nowCol + 2), .Cells(nowRow + 1, nowCol + 2)).Merge
            With Range(.Cells(nowRow, nowCol), .Cells(nowRow + 2, nowCol + 2))
                .Interior.Color = RGB(216, 228, 188)
                
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlInsideHorizontal).Weight = xlThin
            End With
            .Cells(nowRow + 2, nowCol) = "Diff.(%)"
            Range(.Cells(nowRow + 2, nowCol), .Cells(nowRow + 2, nowCol + 2)).Interior.Color = RGB(235, 241, 222)
            .Cells(nowRow + 0, nowCol + 3) = "IDS"
            .Cells(nowRow + 0, nowCol + 4) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-1", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 0, nowCol + 5) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-2", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 0, nowCol + 6) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "IDS-3", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 1, nowCol + 3) = "IDmax"
            .Cells(nowRow + 1, nowCol + 4) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "ID1max", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 1, nowCol + 5) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "ID2max", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 1, nowCol + 6) = "=" & getKA(Range(getKA(dSheet.Cells(1, dSheet.Range("BoundRange").Columns.Count * 3 + 2), WaferArray(0, waferNum), 0)), "ID3max", 1) & "*1E6/" & .Cells(nowRow, nowCol + 1).Address
            .Cells(nowRow + 2, nowCol + 4) = "=" & .Cells(nowRow + 1, nowCol + 4).Address & "/" & .Cells(nowRow, nowCol + 4).Address & "-1"
            .Cells(nowRow + 2, nowCol + 5) = "=" & .Cells(nowRow + 1, nowCol + 5).Address & "/" & .Cells(nowRow, nowCol + 5).Address & "-1"
            .Cells(nowRow + 2, nowCol + 6) = "=" & .Cells(nowRow + 1, nowCol + 6).Address & "/" & .Cells(nowRow, nowCol + 6).Address & "-1"
            Range(.Cells(nowRow, nowCol + 4), .Cells(nowRow + 1, nowCol + 6)).NumberFormat = "0.0"
            Range(.Cells(nowRow + 2, nowCol + 4), .Cells(nowRow + 2, nowCol + 6)).NumberFormat = "0.0%"
            
            With Range(.Cells(nowRow, nowCol + 3), .Cells(nowRow + 1, nowCol + 3))
                .Interior.Color = RGB(253, 233, 217)
                .Font.Size = 10
            End With
            .Cells(nowRow + 2, nowCol).Font.Size = 10
            .Columns(2).ColumnWidth = 21.3
        End With
        nowRow = nowRow + 11
    End If
    Next waferNum
    Next i

End Sub
    
    

