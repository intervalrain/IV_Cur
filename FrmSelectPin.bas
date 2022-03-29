
Public Sub UserForm_Initialize()
   Call putPins
End Sub

Private Function putPins()

Dim i As Integer, j As Integer
Dim myParams() As String

Call getDataInfo

myParams = Params

ListActive.Clear
ListNonActive.Clear
ComboX.Clear

On Error GoTo myError
For i = 0 To UBound(myParams, 2)
    If myParams(1, i) <> "NO" Then
        ListActive.AddItem trim(myParams(0, i))
    Else
        ListNonActive.AddItem trim(myParams(0, i))
    End If
        ComboX.AddItem trim(myParams(0, i))
Next i
                
For i = 0 To UBound(myParams, 2)
    If myParams(0, i) = "VD" Then Me.ComboX.Value = "VD": Exit For
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
With ListActive
    For i = 0 To .ListCount - 1
        For j = 0 To UBound(Params, 2)
            If Params(0, j) = .List(i) Then Params(1, j) = ""
        Next j
    Next i
End With
   'Non-Active
With ListNonActive
    For i = 0 To .ListCount - 1
        For j = 0 To UBound(Params, 2)
            If Params(0, j) = .List(i) Then Params(1, j) = "NO"
        Next j
    Next i
End With
   
    goon = True
    Me.Hide
    
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

Public Sub ComboX_Change()

Dim i As Integer

ListActive.Clear
ListNonActive.Clear
'ComboX.Clear

For i = 0 To UBound(Params, 2)
    If Params(1, i) <> "NO" And Params(0, i) <> ComboX.Value Then
        ListActive.AddItem trim(Params(0, i))
    ElseIf Params(0, i) <> ComboX.Value Then
        ListNonActive.AddItem trim(Params(0, i))
    End If
Next i

End Sub

Public Sub UserForm_QueryClose(cancel As Integer, closemode As Integer)

goon = False

End Sub
