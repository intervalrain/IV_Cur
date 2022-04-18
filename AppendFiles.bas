Option Explicit

Sub AppendFile()

    Dim File
    Dim FilePath As String
    Dim trimWord As String
    Dim Filename As String
    
    Dim tmpAry() As String
    
    Dim i As Integer
    Dim j As Integer

    File = Application.GetOpenFilename("txt File, *.txt", 1, "Load files to combine", "Open", True)
    'In case of close the window
    If VarType(Filename) = vbBoolean Then Exit Sub
        
    'To setup the rule of combination
    tmpAry = Split(File(1), "\")
    FilePath = Replace(File(1), tmpAry(UBound(tmpAry)), "")
    Filename = Replace(Replace(File(1), FilePath, ""), ".txt", "")
    trimWord = Replace(Filename, InputBox("Please enter the part you want to keep, it will become the final name", "Trim word", Filename), "")
    
    'Derive Output Files' name
    Dim oFile() As String
    ReDim oFile(1 To 1) As String
    For i = 1 To UBound(File)
        If InStr(File(i), trimWord) Then
            oFile(i) = getCOL(getCOL(File(i), trimWord, 2), ".txt", 1)
            ReDim Preserve oFile(1 To UBound(oFile) + 1)
        Else
            Exit For
        End If
    Next i
    ReDim Preserve oFile(1 To UBound(oFile) - 1)
    
    Dim mFileName As String
    Dim tmpStr As String
    For i = 1 To UBound(oFile)
        mFileName = oFile(i)
        For j = 1 To UBound(File)
            If Right(Replace(File(j), ".txt", ""), Len(mFileName)) = oFile(i) Then
                tmpStr = ReadTextFile(CStr(File(j)))
                Call WriteTextFile(FilePath & "\" & oFile(i) & ".txt", tmpStr)
            End If
        Next j
    Next i
        
End Sub

Public Sub WriteTextFile(ByVal Filename As String, ByVal mStr As String, Optional ByVal Override As Boolean = False)
    Dim FS As New Scripting.FileSystemObject
    Dim f As TextStream
    Dim Iomode As Integer
    
    Iomode = 2                              'ForAppending
    If Override = True Then Iomode = 8      'ForWriting
    Set f = FS.OpenTextFile(Filename, 8, True)
    f.Write mStr
    f.Close
    Set FS = Nothing
    Set f = Nothing
End Sub

Public Function ReadTextFile(mFileName As String)
    Dim FSO As New Scripting.FileSystemObject
    Dim f As TextStream
    
    Set f = FSO.OpenTextFile(mFileName, 1)   'ForReading
    ReadTextFile = f.ReadAll
    f.Close
    Set FSO = Nothing
    Set f = Nothing
End Function
