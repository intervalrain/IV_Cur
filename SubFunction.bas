Option Explicit

Public Function Unit(Params As String)

    Select Case UCase(Left(Params, 1))
        Case "I"
            Unit = "A"
        Case "V"
            Unit = "V"
        Case "C"
            Unit = "C/V"
        Case "R"
            Unit = "Ohm"
        Case "G"
            Unit = "A/V"
        Case Else
            Unit = ""
    End Select

End Function

Public Function SetStr(str As String)

    Dim setSheet As Worksheet
    Dim nowSheet As Worksheet
    Dim Ary_nowsheet() As String
    Dim Ary_setsheet() As String
    Dim Trimp As Boolean
    Dim i As Integer, j As Integer
    Dim tempStr(2) As String
    
    If str = "[length]" Or str = "[width]" Or str = "[bias]" Then Trimp = True
    
    Set setSheet = Worksheets("Setting")
    Set nowSheet = ActiveSheet
    Ary_nowsheet = Split(nowSheet.Name, "_")
    Ary_setsheet = Split(setSheet.Cells(1, 2).Value, "_")
    
    If UBound(Ary_nowsheet) > UBound(Ary_setsheet) Then
        j = UBound(Ary_setsheet)
    Else
        j = UBound(Ary_nowsheet)
    End If
    
    For i = 0 To j
        tempStr(0) = getCOL(Ary_setsheet(i), "[", 1)
        tempStr(2) = getCOL(Ary_setsheet(i), "]", 2)
        tempStr(1) = Replace(Ary_setsheet(i), tempStr(0), "")
        tempStr(1) = Replace(tempStr(1), tempStr(2), "")
        
        If Left(Ary_nowsheet(i), Len(tempStr(0))) = tempStr(0) Then Ary_nowsheet(i) = Replace(Ary_nowsheet(i), tempStr(0), "")
        If Right(Ary_nowsheet(i), Len(tempStr(2))) = tempStr(2) Then Ary_nowsheet(i) = Replace(Ary_nowsheet(i), tempStr(2), "")
    
        If tempStr(1) = str Then
            If Trimp = True Then Ary_nowsheet(i) = CDbl(Replace(Ary_nowsheet(i), "p", "."))
            SetStr = Ary_nowsheet(i)
            Exit For
        End If
    Next

End Function

Public Function WorksheetSetting(nowRange As Range)

   Dim s As Integer, E As Integer, i As Integer
   Dim Count As Integer
   Dim tempAry() As String

   For i = 1 To Len(nowRange.Text)
      If Mid(nowRange.Text, i, 1) = "[" Then s = i
      If Mid(nowRange.Text, i, 1) = "]" Then
         E = i
         If s > 0 Then
            Count = Count + 1
            ReDim Preserve tempAry(Count - 1) As String
            tempAry(UBound(tempAry)) = Mid(nowRange.Text, s, E - s + 1)
         Else
            Debug.Print "Define Error!!"
         End If
         s = 0
      End If
   Next i
   s = 0
   WorksheetSetting = tempAry

End Function

Public Sub SetWaferRange()
    Dim nowSheet As Worksheet
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim nowWafer As String
    Dim nowSite As String
    Dim temp As String
    Dim tempR() As Long
    Dim mCol As Integer
    Dim typeName As String
    Dim Item
    Dim myArray() As String
       
    On Error Resume Next
   
    'For i = ActiveWorkbook.Names.Count To 1 Step -1
    '    If InStr(ActiveWorkbook.Names(i), ActiveSheet.Name) > 0 Then ActiveWorkbook.Names(i).Delete
    'Next i
    ReDim myArray(1, 0)
    ReDim tempR(0) As Long
    
    typeName = "wafer_"
    Set nowSheet = ActiveSheet
    For i = 1 To nowSheet.UsedRange.Rows.Count
        If Left(trim(nowSheet.Cells(i, 1)), 3) = "===" Then
            typeName = "wafer_"
            temp = nowSheet.Cells(i, 1).Value
            nowSheet.Cells(i, 1).ClearContents
            Set nowRange = nowSheet.Cells(i - 1, 1).CurrentRegion
            For j = nowRange.Columns.Count To 1 Step -1
                If InStr(nowRange.Cells(1, j).Value, "Wfr") Then mCol = j: Exit For
            Next j
            nowWafer = trim(Replace(getCOL(nowRange.Cells(1, j).Value, ",", 2), "#", ""))
            nowSite = trim(Replace(nowRange.Cells(1, j + 1).Value, "#", ""))
            myArray(0, UBound(myArray, 2)) = "#" & nowWafer & "-" & nowSite
            ReDim Preserve myArray(1, UBound(myArray, 2) + 1)
            tempR(UBound(tempR)) = i
            ReDim Preserve tempR(UBound(tempR) + 1)
            nowSheet.Names.Add typeName & nowWafer & "_" & nowSite, nowRange
        End If
    Next i

    ReDim Preserve myArray(1, UBound(myArray, 2) - 1)
    ReDim Preserve tempR(UBound(tempR) - 1)
    
    For Each Item In tempR()
        nowSheet.Cells(Item, 1).Value = "'" & temp
    Next Item
    
    WaferArray = myArray
    Set nowSheet = Nothing
End Sub
Public Sub SetPrecheckWaferRange()
    Dim nowSheet As Worksheet
    Dim i As Long, j As Long
    Dim nowRange As Range
    Dim nowWafer As String
    Dim nowSite As String
    Dim temp
    Dim tempR() As Long
    Dim mCol As Integer
    Dim typeName As String
    Dim Item
    Dim myArray() As String
       
    On Error Resume Next
   
    ReDim myArray(1, 0)
    ReDim tempR(0) As Long
    
    typeName = "wafer_"
    Set nowSheet = ActiveSheet
    For i = 1 To Range("BoundRange").Rows.Count
        If Left(trim(nowSheet.Cells(i, 1)), 3) = "===" Then
            typeName = "wafer_"
            temp = nowSheet.Cells(i, 1).Value
            nowSheet.Rows(i).ClearContents
            Set nowRange = Cells(i - 1, 1).CurrentRegion
            For j = nowSheet.Range("BoundRange").Columns.Count To 1 Step -1
                If InStr(nowRange.Cells(1, j).Value, "Wfr") Then mCol = j: Exit For
            Next j
            nowWafer = trim(Replace(getCOL(nowRange.Cells(1, j).Value, ",", 2), "#", ""))
            nowSite = trim(Replace(nowRange.Cells(1, j + 1).Value, "#", ""))
            myArray(0, UBound(myArray, 2)) = "#" & nowWafer & "-" & nowSite
            ReDim Preserve myArray(1, UBound(myArray, 2) + 1)
            tempR(UBound(tempR)) = i
            ReDim Preserve tempR(UBound(tempR) + 1)
            nowSheet.Names.Add typeName & nowWafer & "_" & nowSite, nowRange
        End If
    Next i

    ReDim Preserve myArray(1, UBound(myArray, 2) - 1)
    ReDim Preserve tempR(UBound(tempR) - 1)
    
    For Each Item In tempR()
        For i = 0 To nowSheet.Cells(1, 1).CurrentRegion.Columns.Count / Range("BoundRange").Columns.Count - 1
            nowSheet.Cells(Item, 1 + i * Range("BoundRange").Columns.Count) = "'" & temp
        Next i
    Next Item
    
    WaferArray = myArray
    Set nowSheet = Nothing
End Sub
Sub getDataInfo()
   
    Dim nowSheet As Worksheet
    Dim i As Long
    Dim myParams() As String
    
    Set nowSheet = ActiveSheet
    ReDim myParams(1, 0)
       
    For i = 1 To nowSheet.Cells(1, 1).CurrentRegion.Columns.Count
        myParams(0, UBound(myParams, 2)) = nowSheet.Cells(1, i).Value
        ReDim Preserve myParams(1, UBound(myParams, 2) + 1)
        If InStr(1, myParams(0, UBound(myParams, 2) - 1), "Wfr") > 0 Then
            ReDim Preserve myParams(1, UBound(myParams, 2) - 2)
            Exit For
        End If
    Next
    
    Params = myParams
    Set nowSheet = Nothing
End Sub

Sub updateLotInfo(codeName As String)

    Dim lotSheet As Worksheet
    Dim nowSheet As Worksheet
    
    Set lotSheet = Worksheets("LotSheet")
    
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If Worksheets(i).codeName = codeName Then Set nowSheet = Worksheets(i)
    Next i
    
    For i = 1 To lotSheet.UsedRange.Rows.Count
        If lotSheet.Cells(i, 4).Value = codeName Then lotSheet.Cells(i, 3).Value = nowSheet.Name
    Next i
    
    Set nowSheet = Nothing
    Set lotSheet = Nothing
    
End Sub

Function getLotInfo(codeName As String) As LotInfo
    
    Dim i As Integer
    Dim mLot As LotInfo
    
    For i = 1 To Worksheets("LotSheet").UsedRange.Rows.Count
        If Worksheets("LotSheet").Cells(i, 4).Value = codeName Then
            mLot.Lot = Worksheets("LotSheet").Cells(i, 1).Value
            mLot.Process = Worksheets("LotSheet").Cells(i, 2).Value
            mLot.sheetName = Worksheets("LotSheet").Cells(i, 3).Value
            mLot.Index = codeName
        End If
    Next i
    
    getLotInfo = mLot
    
    
End Function
