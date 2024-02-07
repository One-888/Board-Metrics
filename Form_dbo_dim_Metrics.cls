VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_dbo_dim_Metrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Combo16_Change()
    DoCmd.OpenForm "dbo_fact_Metric_Target", acNormal, , "Metric_Key=" & Me.Metric_ID.Value _
                 & " And YEAR([Date])=" & Str(Me.Combo16.Value)
End Sub

Private Sub Command11_Click()
    Me.Metric_Key.SetFocus
    DoCmd.RunCommand acCmdPaste
End Sub

Private Sub Command13_Click()
    
    Dim myOkCancelValue As Integer
    myOkCancelValue = MsgBox("Do you want To add a Goal Line? ", vbOKCancel)
    
    If myOkCancelValue = vbOK Then
        Dim cmd     As ADODB.Command
        Set cmd = New ADODB.Command
        
        With cmd
            .NamedParameters = True
            .ActiveConnection = "Provider=sqloledb;Server=SQL-STORAGE1;Database=Warehouse;Trusted_Connection=yes;"
            .CommandType = adCmdStoredProc
            .CommandText = "usp_load_00_goal_line_starter_to_target"
            .Parameters.Append .CreateParameter("@Metric", adInteger, adParamInput, , Me.Metric_ID.Value)
            .Execute
        End With
        
        MsgBox "Add Goal (User Defined) Line For " & Me.Metric_ID.Value
        
        Set cmd = Nothing
    End If
    
    'CurrentProject.Connection.Execute "[dbo].[usp_load_00_goal_line_starter_to_target] " & Me.Metric_ID.Value
End Sub

Private Sub Command14_Click()
    exec_sp "[usp_load_00_refresh_all_metric_data_and_goal]"
    
    MsgBox "Refresh All Data"
End Sub

Private Sub Return_Value_Test()
    Dim text_return As String
    Dim cmd         As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .NamedParameters = True
        .ActiveConnection = "Provider=sqloledb;Server=SQL-STORAGE1;Database=Warehouse;Trusted_Connection=yes;"
        .CommandType = adCmdStoredProc
        .CommandText = "usp_test_sp"
        Set Param1 = .CreateParameter("@In_param", adVarChar, adParamInput, 50, Me.Metric_Key.Value)
        .Parameters.Append Param1
        Set Param2 = .CreateParameter("@Out_param", adVarChar, adParamReturnValue, 50)
        .Parameters.Append Param2
        .Execute
        MsgBox .Parameters("@Out_param").Value
    End With
    
    MsgBox text_return
    
    Set cmd = Nothing
End Sub

Private Sub Command18_Click()
    DoCmd.OpenForm "dbo_view_all_metrics1", acNormal, , "Metric_ID=" & Me.Metric_ID.Value _
                 & " And YEAR([Date])=" & Str(Me.Combo16.Value)
End Sub

Private Sub Command19_Click()
    DoCmd.OpenForm "dbo_fact_Metric_Data", acNormal, , "Metric_Key=" & Me.Metric_ID.Value _
                 & " And YEAR([Date])=" & Str(Me.Combo16.Value)
End Sub

Private Sub Command20_Click()
    Me.Filter = "[Metric_Key] <>        'Not Used'"
    Me.FilterOn = True
    Me.OrderBy = "[Metric_Key]"
    Me.OrderByOn = True
    Me.Refresh
End Sub

Private Sub Command21_Click()
    Me.Filter = "[Metric_Key] =        'Not Used' OR [Metric_Key] IS NULL"
    Me.FilterOn = True
    Me.OrderBy = "[Metric_Key]"
    Me.OrderByOn = True
    Me.Refresh
End Sub

Private Sub Command22_Click()
    Me.Filter = "[Metric_Short] LIKE '%Target%'"
    Me.FilterOn = True
    Me.OrderBy = "[Metric_Key]"
    Me.OrderByOn = True
    Me.Refresh
End Sub

Private Sub Command23_Click()
    Me.Filter = "[Metric_Short] LIKE '%AWWA%'"
    Me.FilterOn = True
    Me.OrderBy = "[Metric_Key]"
    Me.OrderByOn = True
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExportAllCode
End Sub

Private Sub Metric_ID_DblClick(Cancel As Integer)
    DoCmd.OpenForm "dbo_view_all_metrics1", acNormal, , "Metric_ID=" & Me.Metric_ID.Value
End Sub

Private Sub Metric_Key_DblClick(Cancel As Integer)
    Me.Filter = "[Metric_Key] =        '" & Me.Metric_Key & "'"
    Me.FilterOn = True
    Me.Refresh
End Sub

Private Sub Metric_Short_DblClick(Cancel As Integer)
    DoCmd.OpenForm "dbo_fact_Metric_Target", acNormal, , "Metric_Key=" & Me.Metric_ID.Value
End Sub

