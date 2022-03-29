Option Explicit
    Dim UF As UserForm
    Dim LB() As MSForms.Label
    Dim CB() As MSForms.CheckBox


Public Sub CB1_Click()
    Dim i As Integer
    
    For i = 0 To UBound(AryTrimStr) + 1
        trimOrNot(i) = CB(i).Value
    Next
    Unload Setting2
End Sub

Public Sub CB2_Click()
    Unload Setting2
End Sub

Public Sub UserForm_Activate()
    
    Dim i, j As Integer
    
    If i > 3 Then
        For i = 4 To UF.Controls.Count
            UF.Controls(i).Caption = ""
        Next
    End If

    ReDim LB(UBound(AryTrimStr) + 1) As MSForms.Label
    ReDim CB(UBound(AryTrimStr) + 1) As MSForms.CheckBox
        
        Set UF = Setting2
        
    For j = 0 To UBound(AryTrimStr) + 1
        Set LB(j) = UF.Controls.Add("Forms.Label.1")
        Set CB(j) = UF.Controls.Add("Forms.Checkbox.1")
        
    '******************Label Setting******************
        If j = UBound(AryTrimStr) + 1 Then
            LB(j).Caption = "(Tail)"
        Else
            LB(j).Caption = AryTrimStr(j)
        End If
        LB(j).Width = Len(LB(j).Caption) * 8
        If j = 0 Then
            LB(0).Left = 8
        Else
            LB(j).Left = LB((j - 1)).Left + Len(AryTrimStr(j - 1)) * 8
        End If
        LB(j).Top = 20
    '******************CheckBox Setting******************
        CB(j).Caption = ""
        CB(j).Left = LB(j).Left
        CB(j).Top = 8 + LB(0).Top
        CB(j).Value = True
    Next
    
    CB(0).Value = False
    CB(1).Value = False
    
End Sub

