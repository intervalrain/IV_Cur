Option Explicit

Public Sub CheckBox2_Click()

    If Setting.CheckBox2 = True Then Setting.CheckBox3 = False
    
End Sub

Public Sub checkbox3_click()

    If Setting.CheckBox3 = True Then Setting.CheckBox2 = False
    If Setting.CheckBox3 = True Then ComboBox3.Text = ""

End Sub

Public Sub CommandButton1_Click()

    y = Setting.ComboBox1
    var1 = Setting.ComboBox2
    var2 = Setting.ComboBox3
    
    setLog = Setting.CheckBox1                          'Set scale of Y-axis as log
    setMerge = Setting.CheckBox2                        'Combine dies together
    setBVD = Setting.CheckBox3                          'Derive BVD value
    setName = Setting.CheckBox4                         'Name by user's setting
    setABS1 = Setting.CheckBox5                         'Make value1 absolute
    setABS2 = Setting.CheckBox6                         'Make value2 absolute
    If TextBox1 <> "" Then Icom = Setting.TextBox1      'Define current value of BVD
                
    Me.Hide
    
    goon = True
        
End Sub

Public Sub CommandButton2_Click()

    Me.Hide
    
    goon = False
    

End Sub


Public Sub UserForm_Activate()
   
Dim i As Integer

Call getDataInfo
    
    '******************Initialize for UserForm******************
    If once = False Then
        Setting.ComboBox1.Clear
        Setting.ComboBox2.Clear
        Setting.ComboBox3.Clear
        Setting.ComboBox3.AddItem ""
    
        For i = 1 To UBound(Params, 2) + 1
            Setting.ComboBox1.AddItem Params(0, i - 1) 'Add options for y
            Setting.ComboBox2.AddItem Params(0, i - 1) 'Add options for var1
            Setting.ComboBox3.AddItem Params(0, i - 1) 'Add options for var2
        Next
    End If
                
                        
        TextBox1.Text = 0.000001                  'Set Itar = 1E-6 for initial UserForm
        CheckBox4.Value = True                    'Set Auto-Naming = true for initial UserForm
        
End Sub


Public Sub UserForm_QueryClose(cancel As Integer, closemode As Integer)

goon = False
once = False

End Sub

