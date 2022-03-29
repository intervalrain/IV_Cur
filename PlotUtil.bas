

Public Sub RevValue(Col_y As Integer, Col_v1 As Integer, Col_v2 As Integer, setABS1 As Boolean, setABS2 As Boolean)

    Dim nowSheet As Worksheet
    
    Set nowSheet = ActiveSheet

    If setABS1 = True Then
        For i = 1 To nowSheet.UsedRange.Rows.Count
            If IsNumeric(nowSheet.Cells(i, Col_y).Value) And Not (nowSheet.Cells(i, Col_y).Value) = "" Then nowSheet.Cells(i, Col_y).Value = Abs(nowSheet.Cells(i, Col_y).Value)
        Next i
    End If
    If nowSheet.Cells(3, Col_v1).Value > nowSheet.Cells(2, Col_v1) Then Set nowSheet = Nothing: Exit Sub
    If setABS2 = True Then
        For i = 1 To nowSheet.UsedRange.Rows.Count
            If IsNumeric(nowSheet.Cells(i, Col_v1).Value) And Not (nowSheet.Cells(i, Col_v1).Value) = "" Then nowSheet.Cells(i, Col_v1).Value = nowSheet.Cells(i, Col_v1).Value * -1
        Next i
    End If
    
    Set nowSheet = Nothing

End Sub

Public Function PlotByVar2(Col_y As Integer, Col_v1 As Integer, Col_v2 As Integer)

    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Dim nowRange As Range
    Dim nowChart As ChartObject
    Dim i As Long, j As Long, k As Long
    
    Dim nowWafer As Integer
    Dim nv1 As Long
    Dim nv2 As Long
    Dim nt As Long
    
    Dim Unit_y As String
    Dim Unit_v1 As String
    Dim Unit_v2 As String
    
    Dim StrW As Double, StrL As Double, StrBias As Double
    Dim StrDef() As String
    Dim mStrDef() As String
    Dim BVD() As Double
    Dim str_BVD() As String
    Dim ChartTitle As String
    
    Dim sec1 As Range, sec2 As Range, sec3 As Range, sec4 As Range, ChartRange As Range
                
    Set nowSheet = ActiveSheet
    Set setSheet = Worksheets("Setting")
    
    ReDim BVD(UBound(WaferArray, 2))
    ReDim str_BVD(UBound(WaferArray, 2))
    
    '==========Auto-Naming Setting==========
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
    
    '==========get Unit==========
        Unit_y = Unit(y)
        Unit_v1 = Unit(var1)
        Unit_v2 = Unit(var2)
    '==========Main==========
    For nowWafer = 0 To UBound(WaferArray, 2)
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
        '==========get BVD==========
        If setBVD And nv2 = 1 Then
            For i = 2 To nt + 1
                If Abs(nowRange.Cells(i, Col_y).Value) >= Abs(Icom) Then BVD(nowWafer) = nowRange.Cells(i, Col_v1).Value: Exit For
            Next i
            If Not BVD(nowWafer) = 0 Then str_BVD(nowWafer) = " ,BVD=" & CStr(BVD(nowWafer)) & "V"
        End If
        '==========Create ChartObject==========
        If WaferArray(1, nowWafer) <> "NO" Then
            Set nowChart = nowSheet.ChartObjects.Add(Range(Columns(1), Columns(UBound(Params, 2) + 1)).Width, (nowSheet.ChartObjects.Count) * 200, 324, 200)
        '==========get data range==========
            For i = 1 To nv2
                Set sec1 = nowRange.Cells(1, Col_y)
                Set sec2 = nowRange.Cells(1, Col_v1)
                Set sec3 = Range(nowRange.Cells(2 + (i - 1) * nv1, Col_y), nowRange.Cells((1 + i * nv1), Col_y))
                Set sec4 = Range(nowRange.Cells(2 + (i - 1) * nv1, Col_v1), nowRange.Cells((1 + i * nv1), Col_v1))
                Set ChartRange = Union(sec1, sec2, sec3, sec4)
        '==========Setup ChartObject==========
                With nowChart.Chart
                    If .SeriesCollection.Count = 0 Then
                        .ChartType = xlXYScatterLinesNoMarkers
                        .SetSourceData Source:=ChartRange
                        
                        'If ActiveChart.SeriesCollection.Count = 0 Then
                        '    ActiveChart.SetSourceData Source:=ChartRange
                        'ElseIf ActiveChart.SeriesCollection.Count > 0 Then
                        '    ActiveChart.SeriesCollection.NewSeries
                        'End If
                        If Not var2 = "" Then
                            .SeriesCollection(.SeriesCollection.Count).Name = var2 & "=" & nowRange.Cells(2, Col_v2).Value
                        Else
                            .SeriesCollection(.SeriesCollection.Count).Name = WaferArray(0, nowWafer)
                        End If
                        .SeriesCollection(.SeriesCollection.Count).XValues = sec4
                        .SeriesCollection(.SeriesCollection.Count).Values = sec3
                        .Axes(xlCategory).CrossesAt = 0
                        '.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(sec4)
                        '.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(sec4)
                        .SetElement (msoElementPrimaryValueGridLinesNone)
                        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                        .SetElement (msoElementPrimaryValueAxisTitleRotated)
                        .SetElement (msoElementChartTitleAboveChart)
                        If Left(Unit_y, 1) = "A" Then .Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
                        If setLog Then
                            .Axes(xlValue).ScaleType = xlLogarithmic
                            .Axes(xlValue).CrossesAt = 1
                        Else
                            .Axes(xlValue).ScaleType = xlLinear
                            .Axes(xlValue).CrossesAt = 0
                        End If
                        '==========ChartTitle==========
                        If InStr(setSheet.Cells(2, 2).Value, "default") Then
                            title = nowSheet.Name & " " & WaferArray(0, nowWafer) & str_BVD(nowWafer)
                        Else
                            title = setSheet.Cells(2, 2).Value
                            For j = 0 To UBound(mStrDef)
                                title = Replace(title, StrDef(j), mStrDef(j))
                            Next j
                            title = Replace(title, "[width]", StrW)
                            title = Replace(title, "[length]", StrL)
                            title = Replace(title, "[bias]", StrBias)
                            title = Replace(title, "[wf]", WaferArray(0, nowWafer))
                            title = Replace(title, "[BVD]", str_BVD(nowWafer))
                        End If
                        
                        If setName Then
                            .ChartTitle.Caption = title
                        Else
                            .ChartTitle.Caption = y & "-" & var1 & " " & WaferArray(0, nowWafer) & str_BVD(nowWafer)
                        End If
                        '==========AxisTitle==========
                        .Axes(xlValue).AxisTitle.Caption = y & "(" & Unit_y & ")"
                        .Axes(xlCategory).AxisTitle.Caption = var1 & "(" & Unit_v1 & ")"
                        If nowSheet.Cells(2, Col_v1).Value > nowSheet.Cells(3, Col_v1).Value Then
                            .Axes(xlValue).ReversePlotOrder = True
                            .Axes(xlCategory).ReversePlotOrder = True
                        End If
                    Else
                        .SeriesCollection.NewSeries
                        .SeriesCollection(.SeriesCollection.Count).Name = var2 & "=" & nowRange.Cells(2 + (i - 1) * nv1, Col_v2).Value
                        .SeriesCollection(.SeriesCollection.Count).XValues = sec4
                        .SeriesCollection(.SeriesCollection.Count).Values = sec3
                    End If
                End With
            Next i
        End If
    Next nowWafer
                   
    Set nowSheet = Nothing
    Set setSheet = Nothing
    
    On Error GoTo ErrHandler

    Exit Function
    
ErrHandler:
    MsgBox "Wrong parameter input!"
End Function

Public Function PlotByWafer(Col_y As Integer, Col_v1 As Integer, Col_v2 As Integer)

    Dim nowSheet As Worksheet
    Dim setSheet As Worksheet
    Dim nowRange As Range
    Dim nowChart As ChartObject
    Dim i As Long, j As Long, k As Long
    
    Dim nowWafer As Integer
    Dim nv1 As Long
    Dim nv2 As Long
    Dim nt As Long
    
    Dim Unit_y As String
    Dim Unit_v1 As String
    Dim Unit_v2 As String
    
    Dim StrW As Double, StrL As Double, StrBias As Double
    Dim StrDef() As String
    Dim mStrDef() As String
    Dim ChartTitle As String
    
    Dim sec1 As Range, sec2 As Range, sec3 As Range, sec4 As Range, ChartRange As Range
        
    Set nowSheet = ActiveSheet
    Set setSheet = Worksheets("Setting")
    
    '==========Auto-Naming Setting==========
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
    
    '==========Get Unit==========
        Unit_y = Unit(y)
        Unit_v1 = Unit(var1)
        Unit_v2 = Unit(var2)
    '==========Main==========
    nowWafer = 0
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
        '==========Create ChartObject==========
    For i = 1 To nv2
        Set nowChart = nowSheet.ChartObjects.Add(Range(Columns(1), Columns(UBound(Params, 2) + 1)).Width, (nowSheet.ChartObjects.Count) * 200, 324, 200)
        '==========Get data range==========
        For nowWafer = 0 To UBound(WaferArray, 2)
            If WaferArray(1, nowWafer) <> "NO" Then
                Set nowRange = nowSheet.Range("wafer_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 1) & "_" & getCOL(getCOL(WaferArray(0, nowWafer), "#", 2), "-", 2))
                Set sec1 = nowRange.Cells(1, Col_y)
                Set sec2 = nowRange.Cells(1, Col_v1)
                Set sec3 = Range(nowRange.Cells(2 + (i - 1) * nv1, Col_y), nowRange.Cells((1 + i * nv1), Col_y))
                Set sec4 = Range(nowRange.Cells(2 + (i - 1) * nv1, Col_v1), nowRange.Cells((1 + i * nv1), Col_v1))
                Set ChartRange = Union(sec1, sec2, sec3, sec4)
    '==========Setup ChartObject==========
                With nowChart.Chart
                    If .SeriesCollection.Count = 0 Then
                        .ChartType = xlXYScatterLinesNoMarkers
                        .SetSourceData Source:=ChartRange
                        
                        'If ActiveChart.SeriesCollection.Count = 0 Then
                        '    ActiveChart.SetSourceData Source:=ChartRange
                        'ElseIf ActiveChart.SeriesCollection.Count > 0 Then
                        '    ActiveChart.SeriesCollection.NewSeries
                        'End If
                        .SeriesCollection(.SeriesCollection.Count).Name = WaferArray(0, nowWafer)
                        .SeriesCollection(.SeriesCollection.Count).XValues = sec4
                        .SeriesCollection(.SeriesCollection.Count).Values = sec3
                        .Axes(xlCategory).CrossesAt = 0
                        '.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(sec4)
                        '.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(sec4)
                        .SetElement (msoElementPrimaryValueGridLinesNone)
                        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                        .SetElement (msoElementPrimaryValueAxisTitleRotated)
                        .SetElement (msoElementChartTitleAboveChart)
                        If Left(Unit_y, 1) = "A" Then .Axes(xlValue).TickLabels.NumberFormatLocal = "0.00E+00"
                        If setLog Then
                            .Axes(xlValue).ScaleType = xlLogarithmic
                            .Axes(xlValue).CrossesAt = 1
                        Else
                            .Axes(xlValue).ScaleType = xlLinear
                            .Axes(xlValue).CrossesAt = 0
                        End If
                        '==========ChartTitle==========
                        If InStr(setSheet.Cells(2, 2).Value, "default") Then
                            If Col_v2 <> 0 Then
                                title = nowSheet.Name & ", " & var2 & "=" & nowRange.Cells(2 + (i - 1) * nv1, Col_v2).Value
                            Else
                                title = nowSheet.Name
                            End If
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
                        
                        If setName Then
                            .ChartTitle.Caption = title
                        Else
                            .ChartTitle.Caption = y & "-" & var1 & " " & WaferArray(0, nowWafer)
                        End If
                        '==========AxisTitle==========
                        .Axes(xlValue).AxisTitle.Caption = y & "(" & Unit_y & ")"
                        .Axes(xlCategory).AxisTitle.Caption = var1 & "(" & Unit_v1 & ")"
                        If nowSheet.Cells(2, Col_v1).Value > nowSheet.Cells(3, Col_v1).Value Then
                            .Axes(xlValue).ReversePlotOrder = True
                            .Axes(xlCategory).ReversePlotOrder = True
                        End If
                    Else
                        .SeriesCollection.NewSeries
                        .SeriesCollection(.SeriesCollection.Count).Name = WaferArray(0, nowWafer)
                        .SeriesCollection(.SeriesCollection.Count).XValues = sec4
                        .SeriesCollection(.SeriesCollection.Count).Values = sec3
                    End If
                End With
            End If
        Next nowWafer
    Next i
                   
    Set nowSheet = Nothing
    Set setSheet = Nothing

    On Error GoTo ErrHandler

    Exit Function
    
ErrHandler:
    MsgBox "Wrong parameter input!"
End Function
