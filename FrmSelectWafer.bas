
Private Function putWafer()
   
   Call getDataInfo
   Call SetWaferRange

   ListActive.Clear
   ListNonActive.Clear
   On Error GoTo myError
   For i = 0 To UBound(WaferArray, 2)
      If WaferArray(1, i) <> "NO" Then
         ListActive.AddItem trim(WaferArray(0, i))
      Else
         ListNonActive.AddItem trim(WaferArray(0, i))
      End If
   Next i
   Exit Function
myError:
   Call Initial
   Resume
End Function

Public Sub CmdAdd_Click()
   Dim i As Integer
   
   For i = 0 To ListNonActive.ListCount - 1
      If ListNonActive.Selected(i) Then
         ListActive.AddItem ListNonActive.List(i)
      End If
   Next i
   For i = ListNonActive.ListCount - 1 To 0 Step -1
      If ListNonActive.Selected(i) Then
         ListNonActive.RemoveItem (i)
      End If
   Next i

End Sub

Public Sub CmdAddAll_Click()
   Dim i As Integer
   
   For i = 0 To ListNonActive.ListCount - 1
      ListActive.AddItem ListNonActive.List(i)
   Next i
   For i = ListNonActive.ListCount - 1 To 0 Step -1
      ListNonActive.RemoveItem (i)
   Next i

End Sub

Public Sub cmdOK_Click()
   Dim i As Long
   Dim j As Long
   
   'Active
   For i = 0 To ListActive.ListCount - 1
      For j = 0 To UBound(WaferArray, 2)
         If WaferArray(0, j) = ListActive.List(i) Then WaferArray(1, j) = ""
      Next j
   Next i
   'Non-Active
   For i = 0 To ListNonActive.ListCount - 1
      For j = 0 To UBound(WaferArray, 2)
         If WaferArray(0, j) = ListNonActive.List(i) Then WaferArray(1, j) = "NO"
      Next j
   Next i
   
   setwafer = True
   
   Unload Me
End Sub

Public Sub CmdRemove_Click()
   Dim i As Integer
   
   For i = 0 To ListActive.ListCount - 1
      If ListActive.Selected(i) Then
         ListNonActive.AddItem ListActive.List(i)
      End If
   Next i
   For i = ListActive.ListCount - 1 To 0 Step -1
      If ListActive.Selected(i) Then
         ListActive.RemoveItem (i)
      End If
   Next i
End Sub

Public Sub CmdRemoveAll_Click()
   Dim i As Integer
   
   For i = 0 To ListActive.ListCount - 1
      ListNonActive.AddItem ListActive.List(i)
   Next i
   For i = ListActive.ListCount - 1 To 0 Step -1
      ListActive.RemoveItem (i)
   Next i
End Sub


Public Sub UserForm_Initialize()
   Call putWafer
End Sub
