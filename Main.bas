Const sSheet As String = "Setting"
Const pSheet As String = "PPT"
    
Public AryTrimStr
Public trimOrNot() As Boolean

Public WaferArray() As String
Public Params() As String

Public goon As Boolean
Public once As Boolean
Public setwafer As Boolean
Public setLog As Boolean
Public setMerge As Boolean
Public setBVD As Boolean
Public setName As Boolean
Public Icom As Double
Public setABS1 As Boolean
Public setABS2 As Boolean

Public y As String
Public var1 As String
Public var2 As String

Public Type LotInfo
    Lot As String
    Process As String
    sheetName As String
    Index As String
End Type


Option Explicit

Public Sub LoadTxt()

    Dim i As Long, j As Long, k As Long
    
    Dim Filename
    Dim AryFileName() As String
    Dim tempStr As String
    Dim ArytempStr
    Dim nSheet As String
    Dim SkipTrimming As Boolean
    Dim mLot() As LotInfo
    
    SkipTrimming = True
    '==========¨ú±oÀÉ®×¸ô®|==========
    If Worksheets(sSheet).Cells(3, 2).Value = "[on]" Then
        Dim Path As String
        Dim tempFile As String
        
        tempFile = Application.GetOpenFilename("txt File, *.txt", 1, "Load text file", "Open", False)
        AryFileName = Split(tempFile, "\")
        tempStr = AryFileName(UBound(AryFileName))
        Path = Replace(tempFile, tempStr, "")
        ReDim Filename(Worksheets(sSheet).Cells(3, 1).CurrentRegion.Rows.Count - 3) As String
        For i = 4 To Worksheets(sSheet).Cells(3, 1).CurrentRegion.Rows.Count
            If Worksheets(sSheet).Cells(i, 1).Value = "" Then ReDim Preserve Filename(i - 4): Exit For
            On Error GoTo 0
            Filename(i - 3) = Path & getCOL(tempStr, "_", 1) & "_" & getCOL(tempStr, "_", 2) & "_" & trim(Worksheets(sSheet).Cells(i, 1).Value) & ".txt"
        Next
    Else
        Filename = Application.GetOpenFilename("txt File, *.txt", 1, "Load text file", "Open", True)
    End If
    '==========¿ï¾Ü¨ú®ø©ÎÃö³¬µøµ¡==========
    If VarType(Filename) = vbBoolean Then Exit Sub
    ReDim mLot(1 To UBound(Filename))
    
    '==========¥u¨úÀÉ¦W==========
    For i = 1 To UBound(Filename)
        AryFileName = Split(Filename(i), "\")
        tempStr = Replace(AryFileName(UBound(AryFileName)), ".txt", "")
        ArytempStr = Split(tempStr, "_")
        mLot(i).Lot = ArytempStr(0)
        mLot(i).Process = ArytempStr(1)
                
        '==========¦Û°Ê³B²zÀÉ¦W==========
        If i = 1 Or SkipTrimming = False Then
            AryTrimStr = ArytempStr
            ReDim trimOrNot(UBound(AryTrimStr) + 1) As Boolean
            Setting2.Show
            SkipTrimming = True
        End If
        
        If UBound(AryTrimStr) > UBound(ArytempStr) Then
                Dim temp As Integer
                temp = UBound(ArytempStr)
                ReDim ArytempStr(UBound(AryTrimStr))
            For j = 0 To temp
                ArytempStr(j) = getCOL(tempStr, "_", j + 1)
            Next
        End If
        
        Dim tail As Integer
            If trimOrNot(UBound(trimOrNot)) = True Then
                tail = UBound(ArytempStr)
            Else
                tail = UBound(AryTrimStr)
            End If
        
        For j = 0 To UBound(AryTrimStr)
            If trimOrNot(j) = False And AryTrimStr(j) <> ArytempStr(j) Then
                SkipTrimming = False
                i = i - 1
                Exit For
            End If
            If trimOrNot(j) = False And AryTrimStr(j) = ArytempStr(j) Then ArytempStr(j) = ""
        Next j
        If SkipTrimming = True Then

        '==========§¹¦¨ÀÉ¦W==========
            For k = 0 To tail
                If ArytempStr(k) <> "" Then nSheet = nSheet & "_" & ArytempStr(k)
            Next
            If nSheet = "" Then
                MsgBox ("Worksheet name cannot be empty.")
                Exit Sub
            Else
                nSheet = Mid(nSheet, 2)
            End If

        '==========Load Files=========
            Call LoadFile(Filename(i), nSheet)
            
        '==========¬ö¿ýlotInfo=========
            Dim lotSheet As Worksheet
            Dim mRow As Integer
            Set lotSheet = AddSheet("LotSheet", False, "Setting")
            If lotSheet.Cells(1, 1) <> "LOT" Then
                lotSheet.Cells(1, 1) = "LOT"
                lotSheet.Cells(1, 2) = "PROCESS"
                lotSheet.Cells(1, 3) = "SHEETNAME"
                lotSheet.Cells(1, 4) = "INDEX"
            End If
            mLot(i).sheetName = nSheet
            mLot(i).Index = Worksheets(nSheet).codeName
            
            Worksheets(nSheet).Activate
            Call Initial
            nSheet = ""
        End If
    Next i


    'lotSheet.Visible = xlSheetHidden
    For i = 1 To UBound(Filename)
        mRow = lotSheet.UsedRange.Rows.Count
        lotSheet.Cells(mRow + 1, 1).Value = mLot(i).Lot
        lotSheet.Cells(mRow + 1, 2).Value = mLot(i).Process
        lotSheet.Cells(mRow + 1, 3).Value = mLot(i).sheetName
        lotSheet.Cells(mRow + 1, 4).Value = mLot(i).Index
    Next i
    On Error GoTo 0

End Sub
    
Public Sub Initial()

    Dim i As Long, j As Long
    Dim Col_wf As Integer
    Dim n0 As Long
    Dim ns As Integer
    Dim new_ns As Integer
    Dim n_wafer As Integer
    Dim mRow As Integer
    Dim sysArray
    Dim nowSheet As Worksheet
    Set nowSheet = ActiveSheet
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(nowSheet.Name, sysArray) Then Exit Sub
    
        Columns(1).Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1)), TrailingMinusNumbers:=True
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Left(Cells(i, 1).Value, 2) = "NO" Then
            For j = 1 To i - 1
                Rows(1).Delete
            Next j
            Exit For
        End If
    Next i
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If i > 1 And mRow = 0 Then
            If Cells(i, 1).Value = Cells(1, 1).Value Then mRow = i: Exit For
        End If
    Next i
    
    For i = 1 To nowSheet.UsedRange.Columns.Count
        If Cells(1, i).Value = "" Or Left(Cells(1, i).Value, 3) = "Wfr" Then
            Col_wf = i
            If Left(Cells(1, i).Value, 3) = "Wfr" Then
            Else
                Cells(1, i).Value = "Wfr #"
                Cells(1, i + 1).Value = "#"
            End If
            Exit For
        End If
    Next i
            
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Cells(i, 1).Value = "" Then Rows(i).Delete: Exit For
    Next i
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Cells(i, 1).Value = Cells(i + 1, 1).Value And Left(Cells(i, 1).Value, 3) = "===" Then Rows(i).Delete
    Next i
           
    If Left(Cells(1, 1).Value, 2) = "NO" Then
            Columns(1).Delete
    End If
    
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Left(Cells(i, 1).Value, 2) <> "" Then
            If Left(Cells(i, 1).Value, 2) = "--" Then Rows(i).Delete: i = i - 1
        Else
            Exit For
        End If
    Next
    
    n_wafer = Application.WorksheetFunction.CountA(Range(Cells(1, Col_wf), Cells(ActiveSheet.Rows.Count, Col_wf)))
    n0 = nowSheet.UsedRange.Rows.Count
    If mRow <> 0 Then
        ns = mRow - 1
    Else
        ns = n0
    End If
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Cells(i, Col_wf).Value = "" And Not IsNumeric(Cells(i, 1).Value) And Not Left(Cells(i, 1).Value, 1) = "=" Then
            Rows(i).Delete
        End If
    Next i
    
    If n0 / n_wafer <> ns Then
        If Cells(mRow, Col_wf).Value = Cells(1, Col_wf).Value And Cells(mRow, Col_wf + 1).Value = Cells(1, Col_wf + 1).Value Then
            new_ns = n0 / (n_wafer / 2)
            For j = new_ns * (n_wafer / 2 - 1) + mRow To 1 Step -new_ns
                Rows(j).Delete: Rows(j - 1).Delete
            Next j
        End If
    End If
    
    Call SetWaferRange
    Call getDataInfo
    Call updateLotInfo(ActiveSheet.codeName)
    
    Cells(1, 1).Select
    once = False
    
End Sub

Public Sub run()

    Dim i As Integer
    Dim Col_y As Integer
    Dim Col_v1 As Integer
    Dim Col_v2 As Integer
    Dim sysArray
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(ActiveSheet.Name, sysArray) Then Exit Sub
                          
    Call Speed
                          
    If setwafer = False Then Call SetWaferRange
        Setting.Show
        If goon = True Then
            For i = 1 To UBound(Params, 2) + 1
                If Params(0, i - 1) = y Then Col_y = i
                If Params(0, i - 1) = var1 Then Col_v1 = i
                If Params(0, i - 1) = var2 Then Col_v2 = i
            Next i
            Call RevValue(Col_y, Col_v1, Col_v2, setABS1, setABS2)
            If setMerge Then
                Call PlotByWafer(Col_y, Col_v1, Col_v2)
            Else
                Call PlotByVar2(Col_y, Col_v1, Col_v2)
            End If
        Else
            Call Unspeed
            Exit Sub
        End If
        
    Call Unspeed
    
    once = True

End Sub


Public Sub SplitSheet()

    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Set nowSheet = ActiveSheet
    Set setSheet = Worksheets(sSheet)
    
    Dim nowRange As Range
    
    Dim i As Long, j As Long
    Dim nowWafer As Integer
    
    Dim nt As Long
    Dim nv1 As Long
    Dim vn2 As Long
    
    Dim sysArray
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(ActiveSheet.Name, sysArray) Then Exit Sub
    
    Call Speed
    If setwafer = False Then Call SetWaferRange

    Setting.Show
    If goon = False Then Call Unspeed: Exit Sub

    '******************derive parameter column******************
    Dim Col_y As Integer
    Dim Col_v1 As Integer
    Dim Col_v2 As Integer

    For i = 1 To UBound(Params, 2) + 1
        If Params(0, i - 1) = y Then Col_y = i
        If Params(0, i - 1) = var1 Then Col_v1 = i
        If Params(0, i - 1) = var2 Then Col_v2 = i
    Next i
        
    '******************derive parameter no.******************
    For nowWafer = 0 To UBound(WaferArray, 2)
        If WaferArray(1, nowWafer) <> "NO" Then
            Set nowRange = nowSheet.Range("wafer_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 1) & "_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 2))
            nt = nowRange.Rows.Count - 1
            For i = 3 To nt + 1
                If nowRange.Cells(2, Col_v1) = nowRange.Cells(i, Col_v1) Then
                    nv1 = i - 2: Exit For
                Else
                    nv1 = nt
                End If
            Next i
            nv2 = nt / nv1
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''20201209
            
            
        End If
    Next nowWafer
            
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''20201209
            
            
    For i = 3 To n0 + 1
        If nowSheet.Cells(2, Col_v1) = nowSheet.Cells(i, Col_v1) Then
            Exit For
        Else
            nv1 = nv1 + 1
        End If
    Next
    
    nv2 = n0 / nv1

        
    '******************Unit of parameter setting******************
    Dim Unit_y As String
    Dim Unit_v1 As String
    Dim Unit_v2 As String
    
    Unit_y = Unit(y)
    Unit_v1 = Unit(var1)
    Unit_v2 = Unit(var2)
  
    '******************Chart Title setting******************
        '******************Str List(by user)******************
        Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String
        Dim StrW, StrL
        Dim strWp As String, strLp As String
        Dim StrBias As Double
        Dim str_wf As String
        Dim str_site As String
        Dim strInfo As String
         
        If setSheet.Cells(1, 2).Value <> "[default]" Then
            str1 = SetStr("[str1]")
            str2 = SetStr("[str2]")
            str3 = SetStr("[str3]")
            str4 = SetStr("[str4]")
            str5 = SetStr("[str5]")
            StrW = SetStr("[width]")
            StrL = SetStr("[length]")
            strWp = SetStr("[widthp]")
            strLp = SetStr("[lengthp]")
            StrBias = SetStr("[bias]")
        End If

    '******************Data Transform******************
    Dim newsheet As Worksheet
        
For z = 1 To ns
    If WaferArray(1, z - 1) <> "NO" Then
        AddSheet ("Data" & z)
        Set newsheet = ActiveSheet
        
        newsheet.Cells(1, 1).Value = var1
        
        For k = 1 To nv2
            If setMerge = True Then
                str_wf = Replace(Replace(nowSheet.Cells((z - 1) * (n0 + 2) + 1, Col_wf).Value, "Wfr, # ", ""), ",Site", "")
                str_site = Replace(nowSheet.Cells((z - 1) * (n0 + 2) + 1, Col_wf + 1).Value, " # ", "")
                strInfo = "#" & str_wf & "-" & str_site
                newsheet.Cells(1, 1 + k).Value = strInfo
            Else
                newsheet.Cells(1, 1 + k).Value = var2 & "=" & nowSheet.Cells(2 + (k - 1) * nv1, Col_v2).Value
            End If
            For i = 1 To nv1
                If setABS1 = True Then
                    newsheet.Cells(1 + i, 1 + k).Value = Abs(nowSheet.Cells((k - 1) * nv1 + 1 + i + (z - 1) * (n0 + 2), Col_y).Value)
                Else
                    newsheet.Cells(1 + i, 1 + k).Value = nowSheet.Cells((k - 1) * nv1 + 1 + i + (z - 1) * (n0 + 2), Col_y).Value
                End If
            Next
        Next
    
            For i = 1 To nv1
                If setABS2 = True Then
                    newsheet.Cells(1 + i, 1).Value = Abs(nowSheet.Cells(i + 1, Col_v1).Value)
                Else
                    newsheet.Cells(1 + i, 1).Value = nowSheet.Cells(i + 1, Col_v1).Value
                End If
            Next
        
        '******************Plot Chart******************
        ActiveSheet.ChartObjects.Add(0 + 54 * (nv2 + 1), 0, 324, 200).Select
        ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
        
            '******************Derive System Str******************
            
                'null
            
            '******************Define Data Source******************
            Dim ChartRange, sec1, sec2, sec3, sec4 As Range
    
            For j = 1 To nv2
                Set sec1 = newsheet.Cells(1, 1)
                Set sec2 = Range(newsheet.Cells(1, 2), newsheet.Cells(1, 1 + nv2))
                Set sec3 = Range(newsheet.Cells(2, 1), newsheet.Cells(1 + nv1, 1))
                Set sec4 = Range(newsheet.Cells(2, 2), newsheet.Cells(1 + nv1, 1 + nv2))
                
                Set ChartRange = Union(sec1, sec2, sec3, sec4)
    
                ActiveChart.SetSourceData Source:=ChartRange
                
                '******************Set Chart Scale Format******************
                If setLog = True Then
                    ActiveChart.Axes(xlValue).ScaleType = xlLogarithmic
                    ActiveChart.Axes(xlCategory).CrossesAt = 0
                Else
                    ActiveChart.Axes(xlValue).ScaleType = xlLinear
                    ActiveChart.Axes(xlValue).CrossesAt = 0
                    ActiveChart.Axes(xlCategory).CrossesAt = 0
                End If
                    ActiveChart.Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
            Next
    
            ActiveChart.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(sec3)
            ActiveChart.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(sec3)
            ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
            ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
            ActiveChart.SetElement (msoElementChartTitleAboveChart)
    
            '******************Title******************
            Dim title As String
            If setSheet.Cells(2, 2).Value = "[default]" Then
                title = nowSheet.Name
            Else
                title = setSheet.Cells(2, 2).Value
                title = Replace(title, "[str1]", str1)
                title = Replace(title, "[str2]", str2)
                title = Replace(title, "[str3]", str3)
                title = Replace(title, "[str4]", str4)
                title = Replace(title, "[str5]", str5)
                title = Replace(title, "[width]", StrW)
                title = Replace(title, "[length]", StrL)
                title = Replace(title, "[bias]", StrBias)
                title = Replace(title, "[widthp]", strWp)
                title = Replace(title, "[lengthp]", strLp)
                title = Replace(title, "[wf]", strInfo)
                title = Replace(title, "[BVD]", "")
            End If
        
            '******************Setting Controller******************
            If setName = True Then
                ActiveChart.ChartTitle.Caption = title
            Else
                ActiveChart.ChartTitle.Caption = y & "-" & var1 & " " & strInfo
            End If
            
            '******************Axis******************
                ActiveChart.Axes(xlValue).AxisTitle.Caption = y & "(" & Unit_y & ")"
                ActiveChart.Axes(xlCategory).AxisTitle.Caption = var1 & "(" & Unit_v1 & ")"
            If nowSheet.Cells(2, Col_v1).Value > nowSheet.Cells(3, Col_v1).Value Then
                If setABS2 = False Then ActiveChart.Axes(xlValue).ReversePlotOrder = True
                If setABS1 = False Then ActiveChart.Axes(xlCategory).ReversePlotOrder = True
            End If
    End If
Next
    Set nowSheet = Nothing
    Set setSheet = Nothing
    Set newsheet = Nothing
    
    Call Unspeed
    
    once = True

End Sub

Public Sub getSheet()

    Dim i As Long, j As Long, k As Long
    Dim xR As Integer, xC As Integer
    Dim nowSheet As Worksheet
    Dim mCount As Long
    Dim mType As String
    Dim sysArray
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    
    Set nowSheet = Worksheets(pSheet)
    j = 12
    k = 1
    mType = nowSheet.Cells(7, 1).Text
    nowSheet.Range(Rows(13), Rows(nowSheet.UsedRange.Rows.Count)).ClearContents
    Range("E2").Value = "=TODAY()"
    mCount = 9
    
    For i = 1 To Worksheets.Count
        If isInArray(Worksheets(i).Name, sysArray) Then
        Else
            Select Case UCase(nowSheet.Cells(7, 1).Text)
                Case "A"
                    mCount = 1
                Case "B"
                    mCount = 2
                Case "C"
                    mCount = 4
                Case "D"
                    mCount = 6
                Case "E"
                    mCount = 9
            End Select
            
            Dim cNum As Integer
            Dim cRow As Integer
            
            cNum = Worksheets(i).ChartObjects.Count
            If Not cNum Mod mCount = 0 Then
                cRow = Int(cNum / mCount) + 1
            Else
                cRow = cNum / mCount
            End If
            
            For xR = 1 To cRow
                nowSheet.Cells(j + xR, 1).Value = "YES"
                nowSheet.Cells(j + xR, 2).Value = Worksheets(i).Name
                nowSheet.Cells(j + xR, 3).Value = "=B" & j + xR
                nowSheet.Cells(j + xR, 4).Value = UCase(nowSheet.Cells(7, 1).Text)
                nowSheet.Cells(j + xR, 5).Value = j + xR - 11
                For xC = 1 To mCount
                    nowSheet.Cells(j + xR, 6 + xC).Value = Worksheets(i).ChartObjects(k).Chart.ChartTitle.Caption
                    On Error Resume Next
                    k = k + 1
                Next
            Next
            j = j + cRow
            k = 1
        End If
    Next
End Sub
Public Sub GenPPT()

    Dim i As Long, j As Long, k As Long
    Dim nowSheet As Worksheet
    Dim SourceSheet As Worksheet
    Dim mPPT As New PowerPoint.Application
    Dim nowPPT As PowerPoint.Presentation
    Dim nowSlide As PowerPoint.Slide
    Dim nowShape As PowerPoint.Shape
    Dim x As Design
    Dim CopyPage As Long, mCount As Long
    Dim nChart As New Collection
    Dim LType As String
    Dim mFile
    '******************set parameter******************
    Const pTitle As Integer = 1     'row 1
    Const pBlank As Integer = 12    'row 12
    Const pOrder As Integer = 5     'col 5
    Const ppTitle As Integer = 3    'col 3
    Const pContent As Integer = 6   'col 6
    Const pType As Integer = 4      'col 4
    '******************pre-check******************
    If Not IsExistSheet(pSheet) Then
        MsgBox "Cannot access worksheet ""PPT"""
        Exit Sub
    End If
    '*********************************************
    Set nowSheet = Worksheets(pSheet)
    CopyPage = nowSheet.Cells(9, 1).Value
'    mPPT.Visible = True
    '******************pre-check2******************
    On Error GoTo ErrHandler
    Set nowPPT = mPPT.Presentations.Open(Application.ThisWorkbook.Path & "\" & "PPT File.pptx")
    mPPT.Visible = True
    On Error GoTo 0
    '******************set ppLayoutTitle******************
    With nowPPT.Slides(1)
        For i = 1 To nowSheet.Cells(pTitle, 1).CurrentRegion.Columns.Count
            If nowSheet.Cells(pTitle, i) = "Date" Then
                .Shapes(i).TextFrame.TextRange.Text = Date
            ElseIf Not nowSheet.Cells(pTitle + 1, i) = "" Then
                .Shapes(i).TextFrame.TextRange.Text = nowSheet.Cells(pTitle + 1, i)
            End If
        Next i
    End With
    On Error Resume Next
    '******************by case******************
    For j = 1 To nowSheet.Cells(pBlank, 1).CurrentRegion.Rows.Count - 1
        If Not UCase(nowSheet.Cells(j + pBlank, 1)) = "" Then
            LType = nowSheet.Cells(j + pBlank, pType).Text
            Set SourceSheet = Worksheets(nowSheet.Cells(j + pBlank, ppTitle - 1).Text)
            '******************CopyPage******************
            nowPPT.Slides(CopyPage).Copy
            Set x = nowPPT.Slides(CopyPage).Design
            nowPPT.Slides.Paste.Design = x
            Set nowSlide = nowPPT.Slides(CopyPage)
            '******************set text******************
            With nowPPT.Slides(nowPPT.Slides.Count)
                If nowSheet.Cells(j + pBlank, ppTitle) = "" Then
                    .Shapes(1).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, ppTitle - 1)
                Else
                    .Shapes(1).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, ppTitle)
                End If
                .Shapes(2).TextFrame.TextRange.Text = nowSheet.Cells(j + pBlank, pContent)
                '******************paste chart******************
                Dim W As Integer
                Dim H As Integer
                Dim cNum As Integer

                cNum = nowSlide.Shapes.Count
                
                For i = 1 To SourceSheet.ChartObjects.Count         'object
                    For k = 0 To 8                                  'site
                        If SourceSheet.ChartObjects(i).Chart.ChartTitle.Caption = nowSheet.Cells(j + pBlank, pContent + k + 1).Value Then
                            SourceSheet.ChartObjects(i).CopyPicture
                            .Shapes.PasteSpecial
                            cNum = cNum + 1
                            '******************Set position******************
                            W = .Shapes(cNum).Parent.Master.Width    '1024
                            H = .Shapes(cNum).Parent.Master.Height   '768
    
                            Select Case LType
                                Case "A" '1
                                   .Shapes(cNum).Width = W * 0.74
                                   .Shapes(cNum).Top = 168
                                   .Shapes(cNum).Left = 88
                                     If k = 1 Then Exit For
                                Case "B" '2
                                   .Shapes(cNum).Width = W * 0.48
                                   .Shapes(cNum).Top = 283.2861
                                   .Shapes(cNum).Left = 8.5 + W * 0.48 * k + 10 * k
                                     If k = 2 Then Exit For
                                Case "C" '4
                                   .Shapes(cNum).Width = W * 0.48
                                   .Shapes(cNum).Top = 102 + H * 0.4 * Int(k / 2)
                                   .Shapes(cNum).Left = 8.5 + W * 0.48 * (k Mod 2) + 10 * (k Mod 2)
                                     If k = 4 Then Exit For
                                Case "D" '6
                                   .Shapes(cNum).Width = W * 0.333
                                   .Shapes(cNum).Top = 198 + H * 0.27 * Int(k / 3)
                                   .Shapes(cNum).Left = 0.5 + W * 0.48 * 0.68 * (k Mod 3) + 5 * (k Mod 3)
                                     If k = 6 Then Exit For
                                Case "E" '9
                                   .Shapes(cNum).Width = W * 0.333
                                   .Shapes(cNum).Top = 100 + H * 0.27 * Int(k / 3)
                                   .Shapes(cNum).Left = 0.5 + W * 0.48 * 0.68 * (k Mod 3) + 5 * (k Mod 3)
                                Case "F" '6
                                     If k = 0 Or k = 3 Then
                                     End If
                                     If k = 1 Then
                                       .Shapes(cNum).Width = W * 0.456323
                                       .Shapes(cNum).Top = 99.744
                                       .Shapes(cNum).Left = 61.53403
                                     End If
                                     If k = 4 Then
                                       .Shapes(cNum).Width = W * 0.456323
                                       .Shapes(cNum).Top = 318.7844
                                       .Shapes(cNum).Left = 61.53403
                                     End If
                                     If k = 2 Or k = 5 Then
                                       .Shapes(cNum).Width = W * 0.415
                                       .Shapes(cNum).Top = 124.68 + 184.47 * Int(k / 3)
                                       .Shapes(cNum).Left = -190 + 297.532 * (k Mod 3) + 5 * (k Mod 3)
                                     End If
                                     If k = 6 Then Exit For
                             End Select
                        End If
                        'Exit For
                    Next k
                Next i
            End With
        End If
    Next j
    nowPPT.Slides(CopyPage).Delete
    
    '******************save file & finish******************
    On Error GoTo 0
    If nowSheet.Cells(4, 2).Value <> "" Then nowPPT.SaveAs (Application.ThisWorkbook.Path & "\" & nowSheet.Cells(4, 2))
    
    Set nowSheet = Nothing
    Set SourceSheet = Nothing
    Set nowSlide = Nothing
    Set nowPPT = Nothing
    Set nowShape = Nothing
    Set mPPT = Nothing
    Exit Sub
    '******************************************************
ErrHandler:
    MsgBox "Cannot access ""PPT file.pptx. Please select PPT sample file manually."""
    mFile = Application.GetOpenFilename("pptx File, *.pptx", 1, "Load PPT sample file", "Open", False)
    If mFile = False Then Exit Sub
    Set nowPPT = mPPT.Presentations.Open(mFile)
    Resume Next

End Sub


Public Sub Version()

    Ver.Show

End Sub

Public Sub Select_Wafer()
    Dim sysArray
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(ActiveSheet.Name, sysArray) Then Exit Sub
    
    If ActiveSheet.UsedRange.Columns.Count >= 30 Then
        FrmSelectWaferPrecheck.Show
    Else
        FrmSelectWafer.Show
    End If
    

End Sub
Public Sub PlotWR()
    
    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Dim nowRange As Range
    Dim nowChart As ChartObject
    Dim i As Long, j As Long, k As Long
    
    Dim nowWafer As Integer
    Dim nt As Long
    Dim Col_v1 As Integer
    
    Dim Unit_y As String
    Dim Unit_v1 As String
    
    Dim StrW As Double, StrL As Double, StrBias As Double
    Dim StrDef() As String
    Dim mStrDef() As String
    Dim ChartTitle As String
    Dim title As String
    
    Dim sec() As Range
    Dim sysArray
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(ActiveSheet.Name, sysArray) Then Exit Sub
      
    On Error GoTo ErrHandler
      
    Call Speed
    Set nowSheet = ActiveSheet
    Set setSheet = Worksheets("Setting")
                         
    If setwafer = False Then Call SetWaferRange
    FrmSelectPin.Show
    If Not goon Then Call Unspeed: Exit Sub
    For i = 0 To UBound(Params, 2)
        If Params(0, i) = FrmSelectPin.ComboX Then Col_v1 = i + 1: Exit For
    Next i
    '==========¨úµ´¹ï­È==========
    For j = 1 To UBound(Params, 2) + 1
        If j = Col_v1 And nowSheet.Cells(3, j).Value < nowSheet.Cells(2, j).Value Then
            For i = 1 To nowSheet.UsedRange.Rows.Count
                If IsNumeric(nowSheet.Cells(i, j).Value) And Not (nowSheet.Cells(i, j).Value) = "" Then nowSheet.Cells(i, j).Value = -1 * (nowSheet.Cells(i, j).Value)
            Next i
        End If
        If Not Params(1, j - 1) = "NO" And UCase(Left(Params(0, j - 1), 1)) = "I" Then
            For i = 1 To nowSheet.UsedRange.Rows.Count
                If IsNumeric(nowSheet.Cells(i, j).Value) And Not (nowSheet.Cells(i, j).Value) = "" Then nowSheet.Cells(i, j).Value = Abs(nowSheet.Cells(i, j).Value)
            Next i
        End If
    Next j
    '==========Auto-Naming¦r¦ê¨ú¼Ë==========
    If Not InStr(LCase(setSheet.Cells(1, 2).Value), "default") > 1 Then
        StrDef = WorksheetSetting(setSheet.Range("B1"))
        ReDim mStrDef(UBound(StrDef)) As String
        For i = 0 To UBound(StrDef)
            mStrDef(i) = SetStr(StrDef(i))
        Next i
        StrW = SetStr("[width]")
        StrL = SetStr("[length]")
        StrBias = SetStr("[bias]")
    End If
    '==========¨ú³æ¦ì==========
        Unit_y = "A"
        Unit_v1 = "V"
    '==========¥Dµ{¦¡==========
    For nowWafer = 0 To UBound(WaferArray, 2)
        Set nowRange = nowSheet.Range("wafer_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 1) & "_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 2))
        nt = nowRange.Rows.Count - 1
        '==========«Ø¥ßChartObject==========
        If WaferArray(1, nowWafer) <> "NO" Then
            Set nowChart = nowSheet.ChartObjects.Add(Range(Columns(1), Columns(UBound(Params, 2) + 1)).Width, (nowSheet.ChartObjects.Count) * 200, 324, 200)
            '==========¨ú±oData½d³ò==========
            ReDim sec(1, 1) As Range
            Set sec(0, 0) = nowRange.Cells(1, Col_v1)
            Set sec(1, 0) = Range(nowRange.Cells(2, Col_v1), nowRange.Cells(1 + nt, Col_v1))
            For i = 0 To UBound(Params, 2)
                If Not Params(1, i) = "NO" And i + 1 <> Col_v1 Then
                    Set sec(0, UBound(sec, 2)) = nowRange.Cells(1, i + 1)
                    Set sec(1, UBound(sec, 2)) = Range(nowRange.Cells(2, i + 1), nowRange.Cells(1 + nt, i + 1))
                    ReDim Preserve sec(1, UBound(sec, 2) + 1) As Range
                End If
            Next i
            ReDim Preserve sec(1, UBound(sec, 2) - 1) As Range
            '==========¼ÐÃD¦WºÙ==========
            If InStr(setSheet.Cells(2, 2).Value, "default") Then
                title = nowSheet.Name & " " & WaferArray(0, nowWafer)
            Else
                title = setSheet.Cells(2, 2).Value
                For j = 0 To UBound(mStrDef)
                    title = Replace(title, StrDef(j), mStrDef(j))
                Next j
                title = Replace(title, "[width]", StrW)
                title = Replace(title, "[length]", StrL)
                title = Replace(title, "[bias]", StrBias)
                title = Replace(title, "[wf]", WaferArray(0, nowWafer))
            End If
            '==========«Ø¥ßChartObject==========
            With nowChart.Chart
                '==========Â^¨úData=========
                For i = 1 To UBound(sec, 2)
                    .SeriesCollection.NewSeries
                    .SeriesCollection(.SeriesCollection.Count).XValues = sec(1, 0)
                    .SeriesCollection(.SeriesCollection.Count).Values = sec(1, i)
                    .SeriesCollection(.SeriesCollection.Count).Name = sec(0, i)
                Next i
                '==========®y¼Ð¶b³]©w=========
                .ChartType = xlXYScatterLinesNoMarkers
                .Axes(xlCategory).CrossesAt = 0
                '.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(sec4)
                '.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(sec4)
                .SetElement (msoElementPrimaryValueGridLinesNone)
                .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                .SetElement (msoElementPrimaryValueAxisTitleRotated)
                .SetElement (msoElementChartTitleAboveChart)
                .Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
                .Axes(xlValue).ScaleType = xlLogarithmic
                .Axes(xlValue).CrossesAt = 1
                .ChartTitle.Caption = title
                .Axes(xlValue).AxisTitle.Caption = "I" & "(" & Unit_y & ")"
                .Axes(xlCategory).AxisTitle.Caption = "V" & "(" & Unit_v1 & ")"
            End With
        End If
    Next nowWafer
                   
    Set nowSheet = Nothing
    Set setSheet = Nothing
    Call Unspeed
Exit Sub

ErrHandler:
    MsgBox "Wrong parameter input!"

End Sub

Public Sub QuickPlot()

    Dim i As Long, j As Long
    Dim Col_y As Integer
    Dim Col_v1 As Integer
    Dim Col_v2 As Integer
    Dim flag As Boolean
    Dim v
    Dim sysArray
    
    Call Speed
    On Error GoTo Err
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    
    For j = 1 To ActiveWorkbook.Worksheets.Count
        If Not isInArray(Worksheets(j).Name, sysArray) Then
            If Worksheets(j).Name = "!START" And ActiveWorkbook.Worksheets.Count > j Then
                Worksheets(j + 1).Activate
                Setting.Show
                Exit For
            End If
        End If
    Next j
    
    If goon = True Then
        For j = 1 To ActiveWorkbook.Worksheets.Count
            If Not isInArray(Worksheets(j).Name, sysArray) Then
                If flag = True Then
                    Worksheets(j).Activate
                    'Call Initial
                    For i = 1 To UBound(Params, 2) + 1
                        If Params(0, i - 1) = y Then Col_y = i
                        If Params(0, i - 1) = var1 Then Col_v1 = i
                        If Params(0, i - 1) = var2 Then Col_v2 = i
                    Next i
                                    
                    Call RevValue(Col_y, Col_v1, Col_v2, setABS1, setABS2)
                    
                    If setMerge Then
                        Call PlotByWafer(Col_y, Col_v1, Col_v2)
                    Else
                        Call PlotByVar2(Col_y, Col_v1, Col_v2)
                    End If
                End If
                If Worksheets(j).Name = "!START" Then flag = True
            End If
        Next j
    Else
        Call Unspeed: Exit Sub
    End If
    
    once = True
    
    Call Unspeed
Exit Sub

Err:
Call Initial
Resume

End Sub

Public Sub QuickPlotWR()

    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Dim nowRange As Range
    Dim nowChart As ChartObject
    Dim i As Long, j As Long, k As Long
    
    Dim nowWafer As Integer
    Dim SheetNum As Integer
    Dim nt As Long
    Dim Col_v1 As Integer
    
    Dim Unit_y As String
    Dim Unit_v1 As String
    
    Dim StrW As Double, StrL As Double, StrBias As Double
    Dim StrDef() As String
    Dim mStrDef() As String
    Dim ChartTitle As String
    Dim title As String
    
    Dim sec() As Range
    Dim sysArray
    Dim flag As Boolean
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    
    Call Speed
    
    For SheetNum = 1 To ActiveWorkbook.Worksheets.Count
        If Not isInArray(Worksheets(SheetNum).Name, sysArray) Then
            If Worksheets(SheetNum).Name = "!START" And ActiveWorkbook.Worksheets.Count > SheetNum Then
                Worksheets(SheetNum + 1).Activate
                FrmSelectPin.Show
                Exit For
            End If
        End If
    Next SheetNum
    
    On Error GoTo ErrHandler
    Set setSheet = Worksheets(sSheet)
    If setwafer = False Then Call SetWaferRange

    For SheetNum = 1 To ActiveWorkbook.Worksheets.Count
        If Not isInArray(Worksheets(SheetNum).Name, sysArray) Then
            If flag Then
                Worksheets(SheetNum).Activate
                Set nowSheet = Worksheets(SheetNum)
                                    
                If Not goon Then Call Unspeed: Exit Sub
                For i = 0 To UBound(Params, 2)
                    If Params(0, i) = FrmSelectPin.ComboX Then Col_v1 = i + 1: Exit For
                Next i
                '==========¨úµ´¹ï­È==========
                For j = 1 To UBound(Params, 2) + 1
                    If j = Col_v1 And nowSheet.Cells(3, j).Value < nowSheet.Cells(2, j).Value Then
                        For i = 1 To nowSheet.UsedRange.Rows.Count
                            If IsNumeric(nowSheet.Cells(i, j).Value) And Not (nowSheet.Cells(i, j).Value) = "" Then nowSheet.Cells(i, j).Value = -1 * (nowSheet.Cells(i, j).Value)
                        Next i
                    End If
                    If Not Params(1, j - 1) = "NO" And UCase(Left(Params(0, j - 1), 1)) = "I" Then
                        For i = 1 To nowSheet.UsedRange.Rows.Count
                            If IsNumeric(nowSheet.Cells(i, j).Value) And Not (nowSheet.Cells(i, j).Value) = "" Then nowSheet.Cells(i, j).Value = Abs(nowSheet.Cells(i, j).Value)
                        Next i
                    End If
                Next j
                '==========Auto-Naming¦r¦ê¨ú¼Ë==========
                If Not InStr(LCase(setSheet.Cells(1, 2).Value), "default") > 1 Then
                    StrDef = WorksheetSetting(setSheet.Range("B1"))
                    ReDim mStrDef(UBound(StrDef)) As String
                    For i = 0 To UBound(StrDef)
                        mStrDef(i) = SetStr(StrDef(i))
                    Next i
                    StrW = SetStr("[width]")
                    StrL = SetStr("[length]")
                    StrBias = SetStr("[bias]")
                End If
                '==========¨ú³æ¦ì==========
                    Unit_y = "A"
                    Unit_v1 = "V"
                '==========¥Dµ{¦¡==========
                For nowWafer = 0 To UBound(WaferArray, 2)
                    Set nowRange = nowSheet.Range("wafer_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 1) & "_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 2))
                    nt = nowRange.Rows.Count - 1
                    '==========«Ø¥ßChartObject==========
                    If WaferArray(1, nowWafer) <> "NO" Then
                        Set nowChart = nowSheet.ChartObjects.Add(Range(Columns(1), Columns(UBound(Params, 2) + 1)).Width, (nowSheet.ChartObjects.Count) * 200, 324, 200)
                        '==========¨ú±oData½d³ò==========
                        ReDim sec(1, 1) As Range
                        Set sec(0, 0) = nowRange.Cells(1, Col_v1)
                        Set sec(1, 0) = Range(nowRange.Cells(2, Col_v1), nowRange.Cells(1 + nt, Col_v1))
                        For i = 0 To UBound(Params, 2)
                            If Not Params(1, i) = "NO" And i + 1 <> Col_v1 Then
                                Set sec(0, UBound(sec, 2)) = nowRange.Cells(1, i + 1)
                                Set sec(1, UBound(sec, 2)) = Range(nowRange.Cells(2, i + 1), nowRange.Cells(1 + nt, i + 1))
                                ReDim Preserve sec(1, UBound(sec, 2) + 1) As Range
                            End If
                        Next i
                        ReDim Preserve sec(1, UBound(sec, 2) - 1) As Range
                        '==========¼ÐÃD¦WºÙ==========
                        If InStr(setSheet.Cells(2, 2).Value, "default") Then
                            title = nowSheet.Name & " " & WaferArray(0, nowWafer)
                        Else
                            title = setSheet.Cells(2, 2).Value
                            For j = 0 To UBound(mStrDef)
                                title = Replace(title, StrDef(j), mStrDef(j))
                            Next j
                            title = Replace(title, "[width]", StrW)
                            title = Replace(title, "[length]", StrL)
                            title = Replace(title, "[bias]", StrBias)
                            title = Replace(title, "[wf]", WaferArray(0, nowWafer))
                        End If
                        '==========«Ø¥ßChartObject==========
                        With nowChart.Chart
                            '==========Â^¨úData=========
                            For i = 1 To UBound(sec, 2)
                                .SeriesCollection.NewSeries
                                .SeriesCollection(.SeriesCollection.Count).XValues = sec(1, 0)
                                .SeriesCollection(.SeriesCollection.Count).Values = sec(1, i)
                                .SeriesCollection(.SeriesCollection.Count).Name = sec(0, i)
                            Next i
                            '==========®y¼Ð¶b³]©w=========
                            .ChartType = xlXYScatterLinesNoMarkers
                            .Axes(xlCategory).CrossesAt = 0
                            '.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(sec4)
                            '.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(sec4)
                            .SetElement (msoElementPrimaryValueGridLinesNone)
                            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                            .SetElement (msoElementPrimaryValueAxisTitleRotated)
                            .SetElement (msoElementChartTitleAboveChart)
                            .Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
                            .Axes(xlValue).ScaleType = xlLogarithmic
                            .Axes(xlValue).CrossesAt = 1
                            .ChartTitle.Caption = title
                            .Axes(xlValue).AxisTitle.Caption = "I" & "(" & Unit_y & ")"
                            .Axes(xlCategory).AxisTitle.Caption = "V" & "(" & Unit_v1 & ")"
                        End With
                    End If
                Next nowWafer
                Set nowSheet = Nothing
            End If
        End If
        If Worksheets(SheetNum).Name = "!START" Then flag = True
    Next SheetNum
    Set setSheet = Nothing
    Set nowChart = Nothing
    Call Unspeed
    Exit Sub

ErrHandler:
    Set setSheet = Nothing
    Set nowChart = Nothing
    Call Unspeed
    MsgBox "Wrong parameter input!"
    Exit Sub
End Sub

Public Sub runManualFunction()
    FrmManualFunction.Show
End Sub


Public Sub ChartSummary()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tempStr As String
    Dim nowSheet As Worksheet
    Dim oldSheet As Worksheet
    Dim sysArray
    Set nowSheet = AddSheet("All_Chart")

    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If Worksheets(i).Name = nowSheet.Name Then Exit For
        If Not isInArray(Worksheets(i).Name, sysArray) Then Worksheets(i).Activate
        If Not ActiveSheet.ChartObjects.Count = 0 Then
            Set oldSheet = ActiveSheet
            For k = 0 To oldSheet.ChartObjects.Count - 1
                oldSheet.Activate
                oldSheet.ChartObjects(k + 1).Select
                Selection.Copy
                nowSheet.Activate
                nowSheet.Cells(1 + 12 * k, 1 + 6 * j).Select
                nowSheet.Paste
                nowSheet.Cells(1 + 12 * k, 1 + 6 * j).Value = oldSheet.ChartObjects(k + 1).Chart.ChartTitle.Caption
            Next k
            Set oldSheet = Nothing
            j = j + 1
        End If
    Next i
    Set nowSheet = Nothing

End Sub

Public Sub ScaleSetting()

    Dim i As Integer
    Dim temp As Double
    
    On Error GoTo Err
    
    temp = InputBox("Maximum of X-axis", "Chart setting...", ActiveSheet.ChartObjects(1).Chart.Axes(xlCategory).MaximumScale)
    
    For i = 1 To ActiveSheet.ChartObjects.Count
        ActiveSheet.ChartObjects(i).Chart.Axes(xlCategory).MaximumScale = temp
    Next i
Exit Sub

Err:
MsgBox ("No ChartObject exists!")
Exit Sub

End Sub

Public Sub delChart()
    
    Dim oldSheet As Worksheet
    Dim i As Integer
    Dim sysArray
    Dim flag As Boolean
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    
    Call Speed
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If Not isInArray(Worksheets(i).Name, sysArray) Then
            If flag = True Then
                Worksheets(i).Activate
                If ActiveSheet.ChartObjects.Count <> 0 Then
                If ActiveSheet.ChartObjects.Count = 1 Then ActiveSheet.ChartObjects(1).Parent.Delete
                If ActiveSheet.ChartObjects.Count > 1 Then
                    ActiveSheet.Shapes.SelectAll
                    Selection.Delete
                End If
        End If
            End If
        End If
        If Worksheets(i).Name = "!START" Then flag = True
    Next i
    Call Unspeed

End Sub

Public Sub BVDtable()

    Dim Ilim As Double
    Dim Im As Integer
    Dim Vm As Integer
    
    Dim sysArray
    Dim nowSheet As Worksheet
    Dim nowRange As Range
    Dim nowWafer As Integer
    
    sysArray = Array("Setting", "PPT", "Precheck", "check", "All_Chart")
    If isInArray(ActiveSheet.Name, sysArray) Or Left(ActiveSheet.Name, 1) = "!" Then Exit Sub

    Set nowSheet = ActiveSheet
    
    Ilim = InputBox("Set Ilim: ", "Setting", 0.000001)
    Im = InputBox("Set ID Column: (Integer, A->1, B->2, and so on)", "Setting", 4)
    Vm = InputBox("Set VD Column: (Integer, A->1, B->2, and so on)", "Setting", 3)
    
    Call SetWaferRange
    
    Dim i As Integer
    Dim BVD() As Double
    ReDim BVD(UBound(WaferArray, 2))
    
    
    For nowWafer = 0 To UBound(WaferArray, 2)

        Set nowRange = nowSheet.Range("wafer_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 1) & "_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 2))
        
        For i = 2 To nowRange.Rows.Count
            If nowRange.Cells(i, Im) > Ilim Then
                BVD(nowWafer) = nowRange.Cells(i, Vm)
                Exit For
            End If
        Next i
    
    Next nowWafer
    
    Dim mLot As LotInfo
    mLot = getLotInfo(ActiveSheet.codeName)
    
    Set nowSheet = AddSheet("Data", False)
    Dim nowRow As Long
    
    For i = 1 To UBound(WaferArray, 2)
        If getCOL(WaferArray(0, 0), "-", 1) <> getCOL(WaferArray(0, i), "-", 1) Then Exit For
    Next i
    
    Dim lotFound As Long
    Dim waferFound As Long
    Dim waferName As String
    Dim siteNum As Integer
    Dim waferNum As Integer
    
    siteNum = i
    waferNum = (UBound(WaferArray, 2) + 1) / siteNum
    nowRow = 1
    
    For nowWafer = 1 To waferNum
    
        waferName = getCOL(WaferArray(0, (nowWafer - 1) * siteNum + 1), "-", 1)
        lotFound = -1
        waferFound = -1
    
        If nowSheet.Cells(1, 1) <> "<Process_ID>" Then
            GoSub printHeader
            nowRow = nowRow + 10
            GoSub printHeaderArray
            nowRow = nowRow + 1
            lotFound = 3
        Else

            For i = 1 To nowSheet.UsedRange.Rows.Count
                If Mid(nowSheet.Cells(i, 2), 2) = mLot.Lot Then lotFound = i: Exit For
            Next i
            
            For i = lotFound + 8 To nowSheet.UsedRange.Rows.Count
                If getCOL(getCOL(nowSheet.Cells(i, 4), "<", 2), "-", 1) = Mid(waferName, 2) Then waferFound = i: Exit For
                If nowSheet.Cells(i, 1) = "<Process_ID>" Then Exit For
            Next i
            
            If lotFound = -1 And waferFound = -1 Then
                nowRow = nowSheet.UsedRange.Rows.Count + 2
                GoSub printHeader
                nowRow = nowRow + 10
                GoSub printHeaderArray
                nowRow = nowRow + 1
            ElseIf lotFound > -1 And waferFound = -1 Then
                For i = lotFound + 8 To nowSheet.UsedRange.Rows.Count
                    If nowSheet.Cells(i, 1).Value = "<Process_ID>" Then
                        nowRow = i
                        Exit For
                    End If
                Next i
                If nowRow < nowSheet.UsedRange.Rows.Count Then
                    nowSheet.Rows(nowRow & ":" & nowRow + 2).Insert
                Else
                    nowRow = nowSheet.UsedRange.Rows.Count + 2
                End If
                GoSub printHeaderArray
                nowRow = nowRow + 1
            Else
                nowRow = nowSheet.Cells(waferFound, 1).CurrentRegion.Row + nowSheet.Cells(waferFound, 1).CurrentRegion.Rows.Count
                nowSheet.Rows(nowRow).Insert
                
            End If
        End If
                
        If Not IsNumeric(nowSheet.Cells(nowRow - 1, 1)) Then
            nowSheet.Cells(nowRow, 1) = 1
        Else
            nowSheet.Cells(nowRow, 1) = nowSheet.Cells(nowRow - 1, 1) + 1
        End If
        
        nowSheet.Cells(nowRow, 2) = mLot.sheetName

        For i = 1 To siteNum
            nowSheet.Cells(nowRow, 3 + i) = BVD((nowWafer - 1) * siteNum + i - 1)
        Next i

    Next nowWafer

Exit Sub

printHeader:
    nowSheet.Cells(nowRow + 0, 1).Value = "<Process_ID>"
    nowSheet.Cells(nowRow + 1, 1).Value = "<Product_ID>"
    nowSheet.Cells(nowRow + 2, 1).Value = "<Lot_ID>"
    nowSheet.Cells(nowRow + 3, 1).Value = "<Test_Plan_ID>"
    nowSheet.Cells(nowRow + 4, 1).Value = "<Limit_File>"
    nowSheet.Cells(nowRow + 5, 1).Value = "<Date/Time>"
    nowSheet.Cells(nowRow + 6, 1).Value = "( LONG REPORT )"
    nowSheet.Cells(nowRow + 7, 1).Value = "'-------------"
    nowSheet.Cells(nowRow + 8, 1).Value = "TYPE_SCALAR"
    nowSheet.Cells(nowRow + 9, 1).Value = "'-------------"
    
    nowSheet.Cells(nowRow + 0, 2).Value = ":" & mLot.Process
    nowSheet.Cells(nowRow + 1, 2).Value = ":x"
    nowSheet.Cells(nowRow + 2, 2).Value = ":" & mLot.Lot
    nowSheet.Cells(nowRow + 3, 2).Value = ":x"
    nowSheet.Cells(nowRow + 4, 2).Value = ":x"
    nowSheet.Cells(nowRow + 5, 2).Value = ":x"
    nowSheet.Cells(nowRow + 6, 2).Value = ":x"
    
Return

printHeaderArray:
    nowSheet.Cells(nowRow, 1).Value = "No./DataType"
    nowSheet.Cells(nowRow, 2).Value = "Parameter"
    nowSheet.Cells(nowRow, 3).Value = "Unit"
    
    For i = 1 To siteNum
        nowSheet.Cells(nowRow, 3 + i) = "<" & Mid(waferName, 2) & "-" & i & ">"
    Next i
    
    nowSheet.Cells(nowRow, siteNum + 4) = "W L"
    nowSheet.Cells(nowRow, siteNum + 5) = "RULE"
    
Return

End Sub
