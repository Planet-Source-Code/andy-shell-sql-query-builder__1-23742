Attribute VB_Name = "modCommon"
Option Explicit


Public gbErrorHandSwitch As Boolean

Public gsUserName As String

Public Enum gDatabaseType
    gdbtAccess
    gdbtSQLServer
End Enum

Public Enum gDriverList
    gAll
    gLocal
End Enum

Public Enum gFormSettingType
    gfstAll
    gfstPositionOnly
    gfstSizeOnly
End Enum


'------------------------------------------------------------------------------
'Description:   Loads a forms top, left, height and width from a sub key under
'               the default application area.
'               Also ensures that some part of the form is visible in case the
'               user changes screen resolution.
'Parameters:    frm - the form to load.
'               fst - the type of settings to load.
'
'Example Usage:
'qLoadFormSettings Me
'qLoadFormSettings Me, gfstPositionOnly
'------------------------------------------------------------------------------
'
Public Sub LoadFormSettings(frm As Form, Optional fst As gFormSettingType = gfstAll, Optional bChangeColours As Boolean = True, Optional bCheckTag As Boolean = False)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim lWindow As Long
    Dim sWindow As String
    Dim i As Integer, s() As String, sStr As String
    Dim ctl As Control
    Const sNOT_SET As String = "not"
    Const iTOP = 0
    Const iLEFT = 1
    Const iHEIGHT = 2
    Const iWIDTH = 3
    
    sWindow = "0,0," & frm.Height & "," & frm.Width & ","
    sStr = GetSetting(gsApplName, "Forms", frm.Name, sWindow)
    s = Split(sStr, ",")
    
'    frm.Top = s(iTOP)
'    frm.Left = s(iLEFT)
    
    ' only position and size the form if it has been set and is in a normal state
        
    If fst = gfstAll Or fst = gfstSizeOnly Then
        If UBound(s()) >= iHEIGHT Then If Val(s(iHEIGHT)) > 0 Then frm.Height = Val(s(iHEIGHT))
        If UBound(s()) >= iWIDTH Then If Val(s(iWIDTH)) > 0 Then frm.Width = Val(s(iWIDTH))
    End If
    
    If UBound(s()) >= iTOP Then
        If Val(s(iTOP)) > 0 And Val(s(iTOP)) < Screen.Height Then
            frm.Top = Val(s(iTOP))
        Else
            frm.Top = 0
        End If
    End If
    
    If UBound(s()) >= iLEFT Then
        If Val(s(iLEFT)) > 0 And Val(s(iLEFT)) < Screen.Width Then
            frm.Left = Val(s(iLEFT))
        Else
            frm.Left = 0
        End If
    End If
    ' ensure the form is in the visible area of the screen
    If frm.Top > Screen.Height - frm.Height Then frm.Top = 0
    If frm.Top < 0 Then frm.Top = 0
    If frm.Left > (Screen.Width - frm.Width) Then frm.Left = 0
    If frm.Left < 0 Then frm.Left = 0
    
    Clear frm

ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("VB5Common", 0, "LoadFormSettings")
End Sub
'------------------------------------------------------------------------------
'Description:   Return the application version in a standard format, ie
'               M.mm.rrrr
'               where   M = App.Major
'                       m = App.Minor
'                       r = App.Revision
'
'Example Usage:
'sString = qGetAppVersion()
'------------------------------------------------------------------------------
'
Public Function GetAppVersion() As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    GetAppVersion = App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "GetAppVersion")
End Function
'=====================================================================
'Description:       Converts Date into American Format
'Parameters         d           - Date
'                   dbt         - Database Format
'
Public Function AMDate(d As Date, Optional dbt As gDatabaseType = gdbtSQLServer) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    If dbt = gdbtSQLServer Then
        AMDate = "'" & Month(d) & "/" & Day(d) & "/" & Year(d) & "'"
    Else
        AMDate = "#" & Month(d) & "/" & Day(d) & "/" & Year(d) & "#"
    End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "AMDate")
End Function
'=====================================================================
'Description:       Converts Date/Time into American Format
'Parameters         d           - Date
'                   dbt         - Database Format
'
Public Function AMDateTime(ByVal d As Date, Optional dbt As gDatabaseType = gdbtSQLServer) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    If dbt = gdbtSQLServer Then
        AMDateTime = "'" & Month(d) & "/" & Day(d) & "/" & Year(d) & " " & TimeValue(d) & "'"
    Else
        AMDateTime = "#" & Month(d) & "/" & Day(d) & "/" & Year(d) & " " & TimeValue(d) & "#"
    End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "AMDateTime")
End Function
'=====================================================================
'Description:       Converts Date into American Format
'Parameters         dbt       - Database Format
'
Public Function AMTS(Optional dbt As gDatabaseType = gdbtSQLServer) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    If dbt = gdbtSQLServer Then
        AMTS = "'" & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & TimeValue(Time) & "'"
    Else
        AMTS = "#" & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & " " & TimeValue(Time) & "#"
    End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "AMTS")
End Function

'=====================================================================
'Description:       Converts Time into American Standard
'Parameters         t       - Time
'                   dbt     - Database Format
'
Public Function AMTime(ByVal t As Date, Optional dbt As gDatabaseType = gdbtSQLServer) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    If dbt = gdbtSQLServer Then
        AMTime = "'" & TimeValue(t) & "'"
    Else
        AMTime = "#" & TimeValue(t) & "#"
    End If
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "AMTime")
End Function

'=====================================================================
'Description:       Executes SQL Query
'Parameters         db              - Database
'                   sSql            - SQL String
'                   dbt         - Database Format
'GLOBALs            gQLogInd        - Indicator 0- (Default) no Log; 1-Print Log for ALL SQL
'                   gQLogPath       - QLogPath
'
Public Sub SQLExecute(DB As ADODB.Connection, sSql As String, Optional dbt As gDatabaseType = gdbtSQLServer)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim TimerStart As Double, TimerTotal As Double, sPrintLine As String
    TimerStart = Timer
        
    If DB.State = adStateOpen Then
'        DB.BeginTrans
        DB.Execute sSql, , adCmdText + adExecuteNoRecords
'        DB.CommitTrans
    End If
    

rrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("VB5Common", 0, "SQLExecute", sSql)
End Sub
'=====================================================================
'Description:       Executes SQL Query and Creates ADO Recordset
'Parameters         db              - Database
'                   sSql            - SQL String
'GLOBALs            gQLogInd        - Indicator 0- (Default) no Log; 1-Print Log for ALL SQL
'                   gQLogPath       - QLogPath
'
Public Function SQLOpenRecordset(DB As ADODB.Connection, sSql As String) As ADODB.Recordset
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim l As Long
    Dim TimerStart As Double, TimerTotal As Double, sPrintLine As String
    TimerStart = Timer
    
    If DB.State = adStateOpen Then
        Set SQLOpenRecordset = DB.Execute(sSql, , adCmdText)
    End If
    
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "SQLOpenRecordset", sSql)
End Function

'=======================================================
'   System Version in format V.0.0.0000
'
'
Public Function SystemVersion() As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    SystemVersion = "V." & App.Major & "." & App.Minor & "." & App.Revision
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "SystemVersion")
End Function


'========================================================================
'Description:   Error Handler Process and generate Error messages in Log files
'PARAMETERS:    sFormName       - Form Name
'               ind             - Group or line Index variable (used mainly if Subs with Index parameter
'               SubName         - Procedure number withing a Form
'               sSql            - SQl Statement or another String passed from the form
'               bLogOnly        - Put only in Database Log
'
'
Public Sub ErrorHandler(sFormName As String, ind As Integer, SubName As String, Optional sSql As String, Optional bLogOnly As Boolean)
    Dim iFreeFileN As Integer    ' Next Free File Number
    Dim sPrintLine As String     ' Line to be printed
    Dim sLogName As String       ' Log file name
    Dim lErrNum As Long, sErrDesc As String, sErrSource As String
                
    ' Show Indicator
    Screen.MousePointer = vbDefault
    sErrDesc = Err.Description
    lErrNum = Err.Number
    sErrSource = Err.Source
    sPrintLine = SystemVersion & " - " & " - " & sFormName & "  -  " & SubName & " - Section - " & CStr(ind) & " - " & sErrDesc & " Error # " & lErrNum & "  " & sErrSource & " " & sSql
    
    'Forming Log File Name and open it
    If Dir(gsPathLogs, vbDirectory) = "" Or gsPathLogs = "" Then gsPathLogs = App.Path & "\"
    sLogName = gsPathLogs & "log" & Year(Date) & Month(Date) & Day(Date) & ".txt"
    
    iFreeFileN = FreeFile
    Open sLogName For Append Shared As #iFreeFileN
    ' Printing Error Message
    Print #iFreeFileN, sPrintLine, Time
    MsgBox sPrintLine, , "ERROR MESSAGE"
    Close #iFreeFileN
    
End Sub


'================================================================================
'Description:   Converts SQL String into The appropriate format
'PARAMETERS:    sStr - Source String
'
'
Public Function SQLCheck(sStr As String) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim s As String
    s = Replace(sStr, "'", "''")
    s = Replace(s, "", """")
    SQLCheck = "'" & s & "'"
ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "SQLCheck")
End Function

'=================================================================
' SUB - Cleanup ALL controls on the form and Setup TDBGrid
' PARAMETERS:   frm         - Form Object
'
'
Public Sub Clear(frm As Form)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim ctl As Control, a
    For Each ctl In frm
        Select Case UCase(TypeName(ctl))
'            Case "LABEL":       ctl.Caption = ""
            Case "COMBOBOX":    If ctl.Tag <> "No" Then ctl.Text = ""
            Case "TEXTBOX":     ctl.Text = ""
            Case "CHECKBOX":    ctl.Value = vbUnchecked
            Case "FPDATETIME":  ctl.Text = ""
            
        End Select
    Next ctl
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("VB5Common", 0, "Clear")
End Sub

'------------------------------------------------------------------------------
'Description:   Saves a forms top, left, height and width to a sub key under
'               the default application area.
'Parameters:    frm - the form to save.
'               fst - the type of settings to save.
'
'Example Usage:
'qSaveFormSettings Me
'qSaveFormSettings Me, gfstPositionOnly
'------------------------------------------------------------------------------
'
Public Sub SaveFormSettings(frm As Form, Optional fst As gFormSettingType = gfstPositionOnly)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim sWindow As String
    
    If frm.Top < 0 Or frm.Left < 0 Or frm.Top > Screen.Height Or frm.Left > Screen.Width Then
        sWindow = 0 & "," & 0 & "," & -1 & "," & -1
    Else
        sWindow = frm.Top & "," & frm.Left & "," & frm.Height & "," & frm.Width
    End If
'    SetRegistryString frm.Name, sWindow, gsFORM_SETTINGS_KEY
    Call SaveSetting(gsApplName, "Forms", frm.Name, sWindow)
    
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("VB5Common", 0, "SaveFormSettings")
End Sub


'=======================================================
'   Find Data Type
'   Parameters: rst     - recordset
'               dbt     - Database Type (SQLServer or ACCESS)
'
'
Public Function DBDataType(iDatatType As Integer, Optional dbt As gDatabaseType = gdbtSQLServer) As Integer
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim i As Integer
    
    Const i_TEXT = 0
    Const i_DATE = 1
    Const i_NUMERIC = 2
    Const i_ELSE = 3
    
    Select Case iDatatType
    Case adSmallInt, adInteger, adSingle, adDouble, _
        adCurrency, adTinyInt, adUnsignedTinyInt, _
        adVariant, adBigInt, adNumeric:
        
                                                                    DBDataType = i_NUMERIC
    Case adDate, adDBTimeStamp:                                     DBDataType = i_DATE
    Case adTypeText, adChar, adVarChar, adWChar, adVarWChar:      DBDataType = i_TEXT
    Case Else:                                                      DBDataType = i_ELSE
    End Select
    
ErrExit:      Exit Function
ErrHandler:   Call ErrorHandler("IMSMain", 0, "DBDataType")
End Function

'================================================================================
'
'   Parses the filename out of a directory string
'   PARAMETERS:    FileName - Full File Path (C:\aaa\asd\as.txt)
'
'
Public Function GetFileName(sFileName As String) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
     
     Dim i As Integer
     
     For i = Len(sFileName) To 1 Step -1
       If Mid(sFileName, i, 1) = "\" Then
         Exit For
       End If
     Next
     
     GetFileName = Mid(sFileName, i + 1)

ErrExit:         Exit Function
ErrHandler:      Call ErrorHandler("VB5Common", 0, "GetFileName")
End Function
'----------------------------------------------------------*
' Name       : UnloadAllForms (MDI)                        *
'----------------------------------------------------------*
' Purpose    : Unloads all forms in an application and     *
'            : sets them to Nothing.                       *
'----------------------------------------------------------*
' Parameters : frmMain Required. Parent form variable.     *
'----------------------------------------------------------*
Public Sub UnloadAllForms()
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim frm As Form

    For Each frm In Forms
       Unload frm
       Set frm = Nothing
    Next frm

ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler("VB5Common", 0, "UnloadAllForms")
End Sub
