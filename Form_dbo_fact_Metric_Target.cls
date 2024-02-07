VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dbo_fact_Metric_Target"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command11_Click()
    Me.Value_Number.SetFocus
    DoCmd.RunCommand acCmdPaste
End Sub

Private Sub Form_Unload(Cancel As Integer)
    exec_sp "[usp_load_00_refresh_all_metric_data_and_goal]"
End Sub

Private Sub Metric_Key_DblClick(Cancel As Integer)
    Me.Filter = "Metric_Key= " & Me.Metric_Key
    Me.FilterOn = True
    Me.Refresh
End Sub
