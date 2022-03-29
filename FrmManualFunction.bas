Option Explicit

Public Sub CB_QuickPlot_Click()
    Me.Hide
    Call QuickPlot
End Sub

Public Sub CB_Precheck_Click()
    Me.Hide
    Call LoadPrecheck
End Sub

Public Sub CB_ChartSummary_Click()
    Me.Hide
    Call ChartSummary
End Sub

Public Sub CB_ScaleSetting_Click()
    Me.Hide
    Call ScaleSetting
End Sub

Public Sub CB_QuickPlotPins_Click()
    Me.Hide
    Call QuickPlotWR
End Sub

Public Sub CB_DelAllChart_Click()
    Me.Hide
    Call delChart
End Sub


Private Sub CB_CheckFile_Click()
    Me.Hide
    Call CheckFile
End Sub

Private Sub CB_BVDSummary_Click()
    Me.Hide
    Call BVDtable
End Sub
