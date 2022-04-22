Public Sub CheckFile()
    Dim i As Long, j As Long
    Dim WaferStr As String
    Dim Filename
    Dim FileID As Long
    Dim temp As String, tempStr As String
    Dim nowSheet As Worksheet
    
    Dim mProduct_ID As String
    Dim mLot_ID As String
    Dim mTester_ID As String
    Dim mTest_Plan_ID As String
    Dim mDateTime As String
    Dim mPreview(9) As String
    Dim mCount As Integer
    Dim mSiteCnt As Integer
    
    Filename = Application.GetOpenFilename("rpt File, *.rpt", 1, "Load rpt file", "Open", True)
    If VarType(Filename) = vbBoolean Then Exit Sub
    Set nowSheet = AddSheet("Check", , "Data")
    nowSheet.Cells(1, 1).Value = "Filename"
    nowSheet.Cells(1, 2).Value = "Shuttle"
    nowSheet.Cells(1, 3).Value = "Lot"
    nowSheet.Cells(1, 4).Value = "Tester_ID"
    nowSheet.Cells(1, 5).Value = "Recipe"
    nowSheet.Cells(1, 6).Value = "Date"
    nowSheet.Cells(1, 7).Value = "SiteNum"
    nowSheet.Cells(1, 8).Value = "WaferNum"
    nowSheet.Cells(1, 9).Value = "Wafer"
    nowSheet.Cells(1, 10).Value = "Preview"
    
    For i = 1 To UBound(Filename)
        
        FileID = FreeFile
        Open Filename(i) For Input As #FileID
        Do Until EOF(FileID)
            Line Input #FileID, temp
            If InStr(temp, "TYPE") Then
                mProduct_ID = getValue(temp, "   ", "TYPE", "=")
                mLot_ID = getValue(temp, "   ", "LOT", "=")
                mTester_ID = getValue(temp, "   ", "TESTER_ID", "=")
                mTest_Plan_ID = getValue(temp, "   ", "Recipe", "=")
                mDateTime = getValue(temp, "   ", "DATE", "=")
            ElseIf InStr(temp, "*** WAFER") Then
                tempStr = tempStr & ", #" & trim(Mid(temp, 11, 3))
                mSiteCnt = (((Len(temp) - Len(Replace(temp, vbTab, ""))) / Len(vbTab)) - 4 + 1) / 2
                mCount = 1
            ElseIf mCount > 0 And Not trim(getCOL(temp, vbTab, 3)) Like "*R*PC*" And Not temp Like "*TEM_offset*" And mCount < 11 Then
                mPreview(mCount - 1) = trim(getCOL(temp, vbTab, 3))
                mCount = mCount + 1
            End If
        Loop
        nowSheet.Cells(i + 1, 1).Value = Mid(Filename(i), InStrRev(Filename(i), "\") + 1)
        nowSheet.Cells(i + 1, 2).Value = mProduct_ID
        nowSheet.Cells(i + 1, 3).Value = mLot_ID
        nowSheet.Cells(i + 1, 4).Value = mTester_ID
        nowSheet.Cells(i + 1, 5).Value = mTest_Plan_ID
        nowSheet.Cells(i + 1, 6).Value = mDateTime
        nowSheet.Cells(i + 1, 7).Value = mSiteCnt
        nowSheet.Cells(i + 1, 8).Value = Len(tempStr) - Len(Replace(tempStr, "#", ""))
        nowSheet.Cells(i + 1, 9).Value = Mid(tempStr, 3)
        tempStr = ""
        For j = 0 To 9
            tempStr = tempStr & mPreview(j) & Chr(10)
        Next j
        nowSheet.Columns(10).ColumnWidth = 100
        nowSheet.Cells(i + 1, 10) = Left(tempStr, Len(tempStr) - Len(Chr(10)))
        tempStr = ""
        Close #FileID
    Next i
    nowSheet.Cells.Font.Name = "Arial"
    nowSheet.Cells.Font.Size = 10
    ActiveWindow.Zoom = 75
    nowSheet.Cells.Columns.AutoFit
    nowSheet.Rows("2:" & nowSheet.UsedRange.Rows.Count).RowHeight = 12.75 * 3
    Set nowSheet = Nothing
    MsgBox "Finished"
End Sub


Public Sub trimSheet()

    Dim srcSheet As Worksheet
    
    Set srcSheet = ActiveSheet
    
    
    Dim i As Long
    Dim n As Integer
    
    n = 153
    
    Dim siteNum As Integer
    siteNum = srcSheet.UsedRange.Rows.Count / n
  
    For i = siteNum To 1 Step -1
        If i Mod 2 = 1 Then
            Range(srcSheet.Rows(i * n), srcSheet.Rows((i - 1) * n + 1)).Delete
        End If
    Next i


End Sub


Sub genWaferName()
    
    Dim nowSheet As Worksheet
    Set nowSheet = Worksheets("Group")
    
    nowSheet.Cells(1, 1) = "SiteName"
    nowSheet.Cells(1, 2) = "SetName"
    
    On Error GoTo Err
    Dim i As Integer
    For i = 0 To UBound(WaferArray, 2)
        nowSheet.Cells(i + 2, 1) = WaferArray(0, i)
        nowSheet.Cells(i + 2, 2) = WaferArray(0, i)
    Next i
Exit Sub
Err:
    MsgBox ("Try to Initial a rawdata sheet before this action.")
End Sub

Sub changeName()

    Dim i As Integer
    For i = 1 To Worksheets.Count
        If Worksheets(i).Cells(1, 1) <> "VG" Then
        Else
            Worksheets(i).Range("H1225") = "Wfr, # 1A,Site"
            Worksheets(i).Range("H1378") = "Wfr, # 1A,Site"
            Worksheets(i).Range("H1531") = "Wfr, # 4A,Site"
            Worksheets(i).Range("H1684") = "Wfr, # 4A,Site"
        End If
        Worksheets(i).Activate
        Call Initial
        
    Next i
    
End Sub

Function setWaferName(ByRef siteName As Object)
    Dim nowSheet As Worksheet
    If Not IsExistSheet("Group") Then Exit Function
    Set nowSheet = Worksheets("Group")
    
    Dim i As Long
    Dim n As Long
    
    On Error Resume Next
    
    For i = 2 To nowSheet.UsedRange.Rows.Count
        If nowSheet.Cells(1, 1) = "" Then Exit For
        siteName.Add nowSheet.Cells(i, 1).Value, nowSheet.Cells(i, 2).Value
    Next i
    
End Function
