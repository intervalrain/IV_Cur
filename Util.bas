Option Explicit

Public mlCalcStatus
Public mbInSpeed

Public Function IsExistSheet(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Integer
   For i = 1 To Worksheets.Count
      If UCase(Worksheets(i).Name) = UCase(sheetName) Then IsExistSheet = True: Exit Function
   Next
   IsExistSheet = False
End Function

Public Function LoadFile(FileAddress As Variant, NewWorkSheetName As String)

        Dim OpenFile As String
        Dim tempStr As String
        OpenFile = "TEXT;" & FileAddress
        
        If Len(NewWorkSheetName) > 31 Then NewWorkSheetName = Right(NewWorkSheetName, 31)
        AddSheet (NewWorkSheetName)
        
        With ActiveSheet.QueryTables.Add(Connection:=OpenFile, Destination:=Range("$A$1"))
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
       
    once = False

End Function

Public Function getCOL(ByVal mString As String, ByVal SplitChar As String, ByVal mNumber As Long)
   Dim tempA
   tempA = Split(mString, SplitChar)
   If mNumber - 1 > UBound(tempA) Or mNumber < 1 Then
      getCOL = ""
   Else
      getCOL = trim(tempA(mNumber - 1))
   End If
End Function

Public Function N2L(ByVal iNumber As Long)
   Dim iLetter As String
   Dim UpInt As Integer
   
   UpInt = (iNumber - 1) \ 26
   If UpInt > 0 Then iLetter = Chr(UpInt + 64)
   iLetter = iLetter & Chr(iNumber - UpInt * 26 + 64)
   N2L = iLetter
   
End Function

Public Function AddSheet(sheetName As String, Optional delOld As Boolean = True, Optional mSheet As String)
    Dim nowSheet As Worksheet
    Dim i As Integer
    If IsExistSheet(sheetName) Then
        If delOld = False Then
            Set AddSheet = Worksheets(sheetName)
            Exit Function
        ElseIf delOld = True Then
            Application.DisplayAlerts = False
            Application.Worksheets(sheetName).Delete
        End If
    End If
    On Error GoTo Err
    If mSheet = "" Then
        Set nowSheet = Worksheets.Add(, Worksheets(Worksheets.Count))
    Else
        Set nowSheet = Worksheets.Add(, Worksheets(mSheet))
    End If
    nowSheet.Name = sheetName
    Application.DisplayAlerts = True
    Set AddSheet = nowSheet
    Set nowSheet = Nothing
Exit Function
Err:
    mSheet = Worksheets(Worksheets.Count).Name
    Resume
End Function

Public Function DelSheet(sheetName As String)
   Dim nowSheet As Worksheet
   Dim i As Integer
   Application.DisplayAlerts = False
   For i = 1 To Worksheets.Count
      If Worksheets(i).Name = sheetName Then Worksheets(i).Delete: Exit For
   Next
   Application.DisplayAlerts = True
End Function

Public Sub Speed()
   On Error Resume Next
   If Not mbInSpeed Then
      Application.ScreenUpdating = False
      Application.DisplayAlerts = False
      Application.EnableEvents = False
      mlCalcStatus = Application.Calculation
      Application.Calculation = xlCalculationManual
      mbInSpeed = True
   Else
      'we are already in speed - don't do the settings again
   End If
End Sub

Public Sub Unspeed()
   On Error Resume Next
   Application.ScreenUpdating = True
   Application.DisplayAlerts = True
   Application.EnableEvents = True
   If mbInSpeed Then
      Application.Calculation = mlCalcStatus
   Else
      'this shouldn't be happening - put calc to auto - safest mode
      Application.Calculation = xlCalculationAutomatic
   End If
   mbInSpeed = False
End Sub

                
Public Sub CleanDataRange()
    Dim i As Integer
    
    For i = ActiveWorkbook.Names.Count To 1 Step -1
        If InStr(ActiveWorkbook.Names(i), ActiveSheet.Name) > 0 Then ActiveWorkbook.Names(i).Delete
    Next i
End Sub
Public Function getValue(SourceStr As String, SplitChar As String, ValueName As String, ValueChar As String)
    Dim tempA, tempB
    Dim i As Long
    tempA = Split(SourceStr, SplitChar)
    For i = 0 To UBound(tempA)
        If InStr(1, tempA(i), ValueChar) <> 0 Then
            tempB = Split(tempA(i), ValueChar)
            If UCase(trim(tempB(0))) = UCase(ValueName) Then getValue = trim(tempB(1)): Exit Function
        End If
    Next
    getValue = ""
End Function


Public Function getKA(mStart As Range, rowKey, colKey, Optional rowFirst As Boolean = True)

    Dim colFirst As Boolean
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim nowSheet As Worksheet
    Dim i As Integer
    
    Set nowSheet = mStart.Worksheet
    rowIndex = -32768
    colIndex = -32768
    
    If rowFirst = True Then
    
        If IsNumeric(rowKey) Then
            rowIndex = rowKey
        Else
            For i = 0 To nowSheet.UsedRange.Rows.Count - mStart.Row + 1
                If mStart.Offset(i, 0).Value = rowKey Then rowIndex = i: Exit For
            Next i
        End If
        If IsNumeric(colKey) Then
            colIndex = colKey
        Else
            For i = 0 To nowSheet.UsedRange.Columns.Count - mStart.Column + 1
                If mStart.Offset(rowIndex, i).Value = colKey Then colIndex = i: Exit For
            Next i
        End If
    Else
        If IsNumeric(colKey) Then
            colIndex = colKey
        Else
            For i = 0 To nowSheet.UsedRange.Columns.Count - mStart.Column + 1
                If mStart.Offset(0, i).Value = colKey Then colIndex = i: Exit For
            Next i
        End If
        mStart = mStart.Offset(0, colIndex)
        If IsNumeric(rowKey) Then
            rowIndex = rowKey
        Else
            For i = 0 To nowSheet.UsedRange.Rows.Count - mStart.Row + 1
                If mStart.Offset(i, colIndex).Value = rowKey Then rowIndex = i: Exit For
            Next i
        End If
    End If

    If rowIndex = -32768 Or colIndex = -32768 Then
        getKA = "#N/A"
    Else
        getKA = nowSheet.Name & "!" & mStart.Offset(rowIndex, colIndex).Address
    End If
    
End Function

Public Function isInArray(mValue As String, mArr)
    Dim v
    Dim tmpStr As String
    For Each v In mArr
        tmpStr = tmpStr & v & ":"
    Next v
    If InStr(tmpStr, mValue) Then
        isInArray = True
    Else
        isInArray = False
    End If

End Function
