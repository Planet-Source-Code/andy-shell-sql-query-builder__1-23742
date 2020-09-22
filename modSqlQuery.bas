Attribute VB_Name = "modSqlQuery"
Option Explicit

Global gdbSQLQ As ADODB.Connection
Global gsApplName As String
Global gsDBConnection As String

Global gsPathLogs As String

Public Sub Main()
    Set gdbSQLQ = New ADODB.Connection
    
    gbErrorHandSwitch = True
    
    If App.PrevInstance Then Exit Sub
    gsApplName = App.EXEName
    gsPathLogs = App.Path
    frmMain.Show

End Sub
