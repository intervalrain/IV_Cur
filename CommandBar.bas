Option Explicit


Private Sub Workbook_Open()

Dim CB As CommandBar
Dim CBCtrl As CommandBarControl
Dim CBB As CommandBarButton

    Application.ScreenUpdating = False

Set CB = Nothing

On Error Resume Next
    
    Application.CommandBars("IVCur").Delete
    
On Error GoTo 0

Set CB = Application.CommandBars.Add(Name:="IVCur", Temporary:=False)

'Button
'Order1: Loadivcurve
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "LoadTxt"
    .FaceId = 109
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!LoadTxt"
    .Enabled = True
End With

'Order2: Initial
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Initial"
    .FaceId = 601
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Initial"
    .Enabled = True
End With

'Order3: Select_Wafer
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Select Wafer"
    .FaceId = 98
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Select_Wafer"
    .Enabled = True
End With


'Order4: Run
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Run"
    .FaceId = 350
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!run"
    .Enabled = True
End With

'Order5: PlotAllPins
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "PlotAllPins"
    .FaceId = 422
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!PlotWR"
    .Enabled = True
End With

'Order6: GenPPT
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "GenPPT"
    .FaceId = 267
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!GenPPT"
    .Enabled = True
End With

'Order7: ManualFunction
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "ManualFuncion"
    .FaceId = 176
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!runManualFunction"
    .Enabled = True
End With

'Order8: Version
Set CBB = CB.Controls.Add(Type:=msoControlButton)

With CBB
    .Caption = "Ver. 6.2"
    .FaceId = 487
    .Style = msoButtonIconAndCaption
    .BeginGroup = True
    .OnAction = ActiveWorkbook.Name & "!Version"
    .Enabled = True
End With

'Menu
With Application.CommandBars("IVCur")
    .Visible = True
    .Position = msoBarTop
End With

    Application.ScreenUpdating = True

End Sub

Private Sub Workbook_Deactivate()

On Error Resume Next
    
    Application.CommandBars("IVCur").Visible = False

On Error GoTo 0

End Sub

Private Sub Workbook_activate()

On Error Resume Next
    
    Application.CommandBars("IVCur").Visible = True

On Error GoTo 0

End Sub



