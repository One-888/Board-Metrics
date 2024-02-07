Attribute VB_Name = "Module1"
Option Compare Database

Public Sub exec_sp(sSP_Name As String)
    Dim cmd         As ADODB.Command
    Set cmd = New ADODB.Command
    
    With cmd
        .NamedParameters = True
        .ActiveConnection = "Provider=sqloledb;Server=SQL-STORAGE1;Database=Warehouse;Trusted_Connection=yes;"
        .CommandType = adCmdStoredProc
        .CommandText = sSP_Name
        .Execute
    End With
    
    Set cmd = Nothing
End Sub

Public Sub ExportAllCode()
    'Add 5/14/2020
    'Click to Export All Codes to two folders
    
    Dim c           As VBComponent
    Dim Sfx         As String        'Suffix
    
    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select
        
        If Sfx <> "" Then
            c.Export _
                     FileName:=CurrentProject.Path & "\" & _
                     c.Name & Sfx
            
            On Error Resume Next
            c.Export _
                     FileName:="C:\Users\vsayakanit\OneDrive - Pittsburgh Water and Sewer Authority\Documents\PowerBI\Board All Codes\vba" & "\" & _
                     c.Name & Sfx
            
            'Export to Source Only Git
            'MsgBox "Save to: " & "C:\Users\vsayakanit\OneDrive - Pittsburgh Water and Sewer\Git\__Update Git\IOS_Source_Only\" & _
            c.Name & Sfx
            
        End If
    Next c
End Sub

