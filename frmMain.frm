VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   8130
   ClientLeft      =   5115
   ClientTop       =   1860
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   0
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Print"
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   41
      ToolTipText     =   "Print "
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Assign"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   9360
      TabIndex        =   37
      ToolTipText     =   "Create Assign Statements for ALL Fields for selected Table"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Fields"
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   10320
      TabIndex        =   36
      ToolTipText     =   "Show ALL Fields for selected Table"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Check"
      Height          =   375
      Index           =   8
      Left            =   6480
      TabIndex        =   35
      ToolTipText     =   "Checks the Query and removes all rubish"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Paste"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   34
      ToolTipText     =   "Paste from Clipboard"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Help"
      Height          =   375
      Index           =   6
      Left            =   720
      TabIndex        =   33
      ToolTipText     =   "Help"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Clean"
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   31
      ToolTipText     =   "Cleans the Box"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Copy"
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   30
      ToolTipText     =   "Copy to Clipboard"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Format"
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   29
      ToolTipText     =   "Format SELECT Query"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   27
      ToolTipText     =   "Exit the Application"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Connect"
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   23
      ToolTipText     =   "Connect to Database"
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Build SQL"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   22
      ToolTipText     =   "Build SQL Queries"
      Top             =   7680
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Format"
      Height          =   1380
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   9135
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   43
         ToolTipText     =   "Add Constant Prefix"
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   5
         Left            =   6600
         TabIndex        =   42
         Text            =   "txtFormat"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   39
         Text            =   "txtFormat"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   38
         ToolTipText     =   "Use defined Recordset Name"
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Dim"
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   28
         ToolTipText     =   "Include Dim Part"
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   26
         ToolTipText     =   "Add Defined number of Spaces"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Update"
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   21
         ToolTipText     =   "Include Update Statement"
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   20
         ToolTipText     =   "Include Delete Statement"
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Insert"
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   19
         ToolTipText     =   "Include Insert Statement"
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   3
         Left            =   6600
         TabIndex        =   18
         Text            =   "txtFormat"
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   16
         ToolTipText     =   "User Text Box Name"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "Use defined SQL String Check Function"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkFormat 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         ToolTipText     =   "Use String Variable Name"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   13
         Text            =   "txtFormat"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   11
         Text            =   "txtFormat"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFormat 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   10
         Text            =   "txtFormat"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Constant Prefix :"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   44
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Recordset Name :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Spaces :"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Text Box Name :"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "SQL String Check Function :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "String Variable Name :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Settings"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12015
      Begin VB.CheckBox chkContinue 
         Caption         =   "Use Continuation"
         Height          =   255
         Left            =   8760
         TabIndex        =   46
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkPartial 
         Caption         =   "Partial Format"
         Height          =   255
         Left            =   10560
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCall 
         Caption         =   "..."
         Height          =   300
         Left            =   11520
         TabIndex        =   32
         ToolTipText     =   "Build Connect Statement"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtDBName 
         Height          =   285
         Left            =   6960
         TabIndex        =   25
         Text            =   "txtDBName"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDB 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "txtDB"
         Top             =   720
         Width           =   9975
      End
      Begin VB.ComboBox cboDB 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Text            =   "cboDB"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "DB Variable Name :"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Connection String :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Database Name :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4965
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8758
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.ListBox lstTables 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msTableName As String

Private Const miC_INSERT As Integer = 0
Private Const miC_DELETE As Integer = 1
Private Const miC_UPDATE As Integer = 2
Private Const miC_DIM As Integer = 3

Private Const miT_STRING_VAR    As Integer = 0
Private Const miT_SQL_CHECK     As Integer = 1
Private Const miT_TEXT_BOX      As Integer = 2
Private Const miT_SPACES        As Integer = 3
Private Const miT_RECORDSET     As Integer = 4
Private Const miT_PREFIX        As Integer = 5

Private msDBName(15) As String
Private msDBConnect(15) As String
Private miDBNameNum As Integer


Private Sub cboDB_Click()
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim i As Integer
    
    For i = 0 To 15
        If cboDB = msDBName(i) Then
            miDBNameNum = i
            txtDB = msDBConnect(i)
            Call SaveSetting("SQLS", "Settings", "DBNum", i)
            cmd(2) = True
            Exit Sub
        End If
    Next i

ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "cboDB_Click")
End Sub

Private Sub cmd_Click(Index As Integer)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim i As Integer
    Dim rstTables As ADODB.Recordset
    Dim rstCols As ADODB.Recordset
    Dim sSql As String, sData As String, sDecl As String
    Dim sLeft As String, s As String
    Dim sView As String, sFunc1(3) As String, sFunc2(3) As String
    Dim iDType As Integer
    
    Const iB_BUILD = 0
    Const iB_EXIT = 1
    Const iB_CONNECT = 2
    Const iB_FORMAT = 3
    Const iB_COPY = 4
    Const iB_CLEAN = 5
    Const iB_HELP = 6
    Const iB_PASTE = 7
    Const iB_CHECK = 8
    Const iB_FIELDS = 9
    Const iB_ASSIGN = 10
    Const iB_PRINT = 11

    Const i_TEXT = 0
    Const i_DATE = 1
    Const i_NUMERIC = 2
    Const i_ELSE = 3
    

    Select Case Index
    Case iB_CLEAN:      rtb.Text = ""
    Case iB_HELP:       frmHelp.Show
    Case iB_PRINT
        cd.Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtb.SelLength = 0 Then
           cd.Flags = cd.Flags + cdlPDAllPages
        Else
           cd.Flags = cd.Flags + cdlPDSelection
        End If
        cd.ShowPrinter
        rtb.SelPrint cd.hDC
        
    
    Case iB_ASSIGN
        MousePointer = vbHourglass
        
        
        If msTableName = "" Then
            MsgBox "Please select Table and try again."
            MousePointer = vbDefault
            Exit Sub
        End If
        
        rtb.Text = ""
        sView = ""
        Set rstCols = gdbSQLQ.OpenSchema(adSchemaColumns)
        
        sSql = "' == " & UCase(msTableName) & " == " & vbCrLf & vbCrLf
        sSql = sSql & "sSql = ""select * from " & UCase(msTableName) & "  where """ & vbCrLf
        sSql = sSql & "Set R_S_T = SQLOpenrecordsetADO(" & txtDBName & ",sSql)" & vbCrLf & vbCrLf
        
        Do Until rstCols.EOF
            If rstCols.Fields(2) = msTableName Then
                sSql = sSql & "T_X_T(m_i_T_" & UCase(rstCols.Fields(3)) & ")  = R_S_T!" & rstCols.Fields(3) & vbCrLf
            End If
            rstCols.MoveNext
        Loop
        sView = sSql
        
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "sSql = ", Space(Val(txtFormat(miT_SPACES))) & " sSql = ")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "T_X_T", Space(Val(txtFormat(miT_SPACES))) & " T_X_T")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "Set R_S_T = ", Space(Val(txtFormat(miT_SPACES))) & " Set R_S_T = ")
        If chkFormat(miT_TEXT_BOX) Then sView = Replace(sView, "T_X_T", txtFormat(miT_TEXT_BOX))
        If chkFormat(miT_RECORDSET) Then sView = Replace(sView, "R_S_T", txtFormat(miT_RECORDSET))
        If chkFormat(miT_PREFIX) Then sView = Replace(sView, "m_i_T_", txtFormat(miT_PREFIX))
        If chkFormat(miT_STRING_VAR) Then sView = Replace(sView, "sSql", txtFormat(miT_STRING_VAR))
        
        rtb.Text = sView
        
    Case iB_FIELDS
        MousePointer = vbHourglass
        
        
        If msTableName = "" Then
            MsgBox "Please select Table and try again."
            MousePointer = vbDefault
            Exit Sub
        End If
        
        rtb.Text = ""
        sView = ""
        Set rstCols = gdbSQLQ.OpenSchema(adSchemaColumns)
        
        sSql = " == " & UCase(msTableName) & " == " & vbCrLf & vbCrLf
        
        Do Until rstCols.EOF
            If rstCols.Fields(2) = msTableName Then
                sSql = sSql & rstCols.Fields(3) & vbCrLf
            End If
            rstCols.MoveNext
        Loop
        
        sView = sSql
        
        rtb.Text = sSql
    
    Case iB_EXIT:
        Unload Me
        End
        
    Case iB_COPY:
        With rtb
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Clipboard.SetText (.SelText)
        End With
    
    Case iB_PASTE:
        rtb.Text = Clipboard.GetText
        
    Case iB_CHECK:
        sView = rtb.Text
        
        If InStr(LCase(sView), "select ") = 0 Then
            MsgBox "It's not Select Query. Put the right one in and try again."
            Exit Sub
        End If
        
        sView = Replace(sView, Chr(13), "")
        sView = Replace(sView, Chr(10), "")
        sView = Replace(sView, "dbo.", "")
        sView = Replace(sView, "from ", " FROM ")
        sView = Replace(sView, "From ", " FROM ")
        sView = Replace(sView, "FROM ", " FROM ")
        
        rtb.Text = sView
        Clipboard.SetText sView
    
    Case iB_FORMAT:
        sView = rtb.Text
        
        If InStr(LCase(sView), "select ") = 0 Then
            MsgBox "It's not Select Query. Put the right one in and try again."
            Exit Sub
        End If
        
        sView = Replace(sView, Chr(13), "")
        sView = Replace(sView, Chr(10), "")
        sView = Replace(sView, "dbo.", "")
        sView = Replace(sView, "sSql = sSql & ", "")
        sView = Replace(sView, "sSql = ", "")
        If chkFormat(miT_STRING_VAR) Then sView = Replace(sView, txtFormat(miT_STRING_VAR) & " = " & txtFormat(miT_STRING_VAR) & " & ", "")
        If chkFormat(miT_STRING_VAR) Then sView = Replace(sView, txtFormat(miT_STRING_VAR) & " = ", "")
        
        sView = SQLFormat(sView)
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "sSql = ", Space(Val(txtFormat(miT_SPACES))) & "sSql = ")
        If chkContinue <> 0 And chkFormat(miT_SPACES) Then sView = Replace(sView, "& """, Space(Val(txtFormat(miT_SPACES))) & "& """)
        sSql = ""
        sSql = sSql & sView
        
        If chkFormat(miT_STRING_VAR) Then sView = Replace(sView, "sSql", txtFormat(miT_STRING_VAR))
        
        rtb.Text = sView
        Clipboard.SetText sView
        
    Case iB_BUILD
        MousePointer = vbHourglass
        
        
        If msTableName = "" Then
            MsgBox "Please select Table and try again."
            MousePointer = vbDefault
            Exit Sub
        End If
        
        rtb.Text = ""
        sView = ""
        Set rstCols = gdbSQLQ.OpenSchema(adSchemaColumns)
        i = 0
        
        sFunc2(i_TEXT) = ")"
        sFunc2(i_DATE) = ")"
        sFunc2(i_NUMERIC) = ")"
        sFunc2(i_ELSE) = ""
        
        If miT_SQL_CHECK Then
            sFunc1(i_TEXT) = txtFormat(miT_SQL_CHECK) & "("
        Else
            sFunc1(i_TEXT) = "SQLCheck("
        End If
        sFunc1(i_DATE) = "AMDateTime("
        sFunc1(i_NUMERIC) = "Val("
        sFunc1(i_ELSE) = ""
        
        ' Creates Insert
            sDecl = ""
            sData = ""
            sSql = vbCrLf
            sSql = sSql & "' INSERT STATEMENT"
            sSql = sSql & vbCrLf
            sSql = sSql & "sSql = ""insert into " & msTableName & " (""" & vbCrLf
            
            Do Until rstCols.EOF
                If rstCols.Fields(2) = msTableName Then
                    sDecl = sDecl & "Private const m_i_T_" & UCase(rstCols.Fields(3)) & " = " & CStr(i) & vbCrLf
                    sSql = sSql & "sSql = sSql & """ & rstCols.Fields(3) & ",""" & vbCrLf
                    iDType = DBDataType(rstCols.Fields(11))
                    sData = sData & "sSql = sSql & " & sFunc1(iDType) & "T_X_T(m_i_T_" & UCase(rstCols.Fields(3)) & ")" & sFunc2(iDType) & " & "",""" & vbCrLf
                    i = i + 1
                End If
                                                            
                rstCols.MoveNext
            Loop
            i = InStrRev(sSql, ",", Len(sSql))
            If i > 0 Then
                sSql = Mid(sSql, 1, i - 1) & """" & vbCrLf
                sSql = sSql & "sSql = sSql  & "") values(""" & vbCrLf
            End If
            If Trim(sData) <> "" Then
                i = InStrRev(sData, ",", Len(sData))
                If i > 0 Then
                    sData = Mid(sData, 1, i - 1) & """" & vbCrLf
                    sData = sData & "sSql = sSql  & "")""" & vbCrLf
                End If
            End If
            
            sData = sData & vbCrLf
            
            sData = sData & "Call SQLExecute(gdb,sSql)" & vbCrLf
            
            If chkType(miC_DIM) Then sView = sView & sDecl & vbCrLf & vbCrLf
            If chkType(miC_INSERT) Then sView = sView & sSql & sData
            
            rstCols.MoveFirst
        '==================================
        
        ' Include Delete Statement
            sDecl = ""
            sData = ""
            sSql = vbCrLf
            sSql = sSql & "' DELETE STATEMENT"
            sSql = sSql & vbCrLf
            
            sSql = sSql & "sSql = ""delete from " & msTableName & " where """ & vbCrLf
            sSql = sSql & vbCrLf
            sSql = sSql & "Call SQLExecute(gdb,sSql)" & vbCrLf
            
            If chkType(miC_DELETE) Then sView = sView & sSql
        '==================================
        
        ' Include Update Statement
            sSql = vbCrLf
            sSql = sSql & "' UPDATE STATEMENT"
            sSql = sSql & vbCrLf
            
            
            sSql = sSql & "sSql = ""update " & msTableName & " set """ & vbCrLf
            
            Do Until rstCols.EOF
                If rstCols.Fields(2) = msTableName Then
                    iDType = DBDataType(rstCols.Fields(11))
                    sSql = sSql & "sSql = sSql & """ & rstCols.Fields(3) & " = " & """ & " & sFunc1(iDType) & "T_X_T(m_i_T_" & UCase(rstCols.Fields(3)) & ")" & sFunc2(iDType) & " & "",""" & vbCrLf
                    i = i + 1
                End If
                                                            
                rstCols.MoveNext
            Loop
            i = InStrRev(sSql, ",", Len(sSql))
            If i > 0 Then
                sSql = Mid(sSql, 1, i - 1) & """" & vbCrLf
                sSql = sSql & "sSql = sSql  & "" where """ & vbCrLf
                sSql = sSql & vbCrLf
                sSql = sSql & "Call SQLExecute(gdb,sSql)" & vbCrLf
            End If
            
            If chkType(miC_UPDATE) Then sView = sView & sSql
        '==================================
        
        rstCols.Close
        
        ' Format
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "' DELETE", Space(Val(txtFormat(miT_SPACES))) & "' DELETE")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "' INSERT", Space(Val(txtFormat(miT_SPACES))) & "' INSERT")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "' UPDATE", Space(Val(txtFormat(miT_SPACES))) & "' UPDATE")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "Call ", Space(Val(txtFormat(miT_SPACES))) & "Call ")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "Private const ", Space(Val(txtFormat(miT_SPACES))) & "Private const ")
        If chkFormat(miT_SPACES) Then sView = Replace(sView, "sSql = ", Space(Val(txtFormat(miT_SPACES))) & "sSql = ")
        
        sView = Replace(sView, "Call SQLExecute(gdb", "Call SQLExecute(" & txtDBName)
        If chkFormat(miT_STRING_VAR) Then sView = Replace(sView, "sSql", txtFormat(miT_STRING_VAR))
        If chkFormat(miT_SQL_CHECK) Then sView = Replace(sView, "sSql", txtFormat(miT_STRING_VAR))
        If chkFormat(miT_TEXT_BOX) Then sView = Replace(sView, "T_X_T", txtFormat(miT_TEXT_BOX))
        If chkFormat(miT_PREFIX) Then sView = Replace(sView, "m_i_T_", txtFormat(miT_PREFIX))
        
        ' Copy to Clipboard
        Clipboard.SetText sView
        
        ' Display
        rtb.Text = sView
        
        MousePointer = vbDefault
    
    
    Case iB_CONNECT
        If txtDB = "" Then
            MsgBox "Please create and save Connection String"
            Exit Sub
        Else
            gsDBConnection = txtDB
        End If
        
        i = 0
NewOpen:
        If gdbSQLQ.State <> adStateOpen Then
           gdbSQLQ.CommandTimeout = 60
           gdbSQLQ.Open gsDBConnection
        Else
            gdbSQLQ.Close
            i = i + 1
            If i > 2 Then
                MsgBox "Can not connect to Database"
                Exit Sub
            Else
                GoTo NewOpen
            End If
            
        End If
        lstTables.Clear
        
        Set rstTables = gdbSQLQ.OpenSchema(adSchemaTables)
        Do Until rstTables.EOF
            If rstTables.Fields(3) = "TABLE" Then
                lstTables.AddItem rstTables.Fields(2)
            End If

            i = i + 1
            rstTables.MoveNext
        Loop
        
        rstTables.Close
        cmd(iB_BUILD).Enabled = True
        cmd(iB_ASSIGN).Enabled = True
        cmd(iB_FIELDS).Enabled = True
        
    End Select
    MousePointer = vbDefault


ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "cmd_Click")
End Sub


Private Sub cmdCall_Click()
    frmDBConnection.Show
End Sub

Private Sub Edit_Click()

End Sub

Private Sub Exit_Click()

End Sub

Private Sub Form_Load()
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim i As Integer
    
    LoadFormSettings Me
    
    Caption = "QUERY BUILDER " & SystemVersion
    
    miDBNameNum = GetSetting("SQLS", "Settings", "DBNum", 0)
    
    cboDB.Clear
    
    For i = 0 To 15
        msDBName(i) = GetSetting("SQLS", "Settings", "DBName_" & CStr(i), "")
        msDBConnect(i) = GetSetting("SQLS", "Settings", "DBConnect_" & CStr(i), "")
        cboDB.AddItem msDBName(i), i
        If i = miDBNameNum Then
            cboDB = msDBName(i)
            txtDB = msDBConnect(i)
        End If
    Next i
    
    For i = 0 To miDBNameNum - 1
        msDBName(i) = GetSetting("SQLS", "Settings", "DBName_" & CStr(i), "")
        cboDB.AddItem msDBName(i), i
    Next i
    
'    txtDB = GetSetting("SQLS", "Settings", "ConnectionString", "Provider=SQLOLEDB.1;Password=tripled;Persist Security Info=True;User ID=sa;Initial Catalog=IMS_BENGALLA;Data Source=W2KIMSTEST")
    txtDBName = GetSetting("SQLS", "Settings", "DBName", "gdb")
    txtFormat(miT_STRING_VAR) = GetSetting("SQLS", "Settings", "sSql", "sSql")
    txtFormat(miT_SQL_CHECK) = GetSetting("SQLS", "Settings", "SQLCheck", "SQLCheck")
    txtFormat(miT_TEXT_BOX) = GetSetting("SQLS", "Settings", "txtBOX", "txtBOX")
    txtFormat(miT_SPACES) = GetSetting("SQLS", "Settings", "Spaces", "4")
    txtFormat(miT_RECORDSET) = GetSetting("SQLS", "Settings", "Recordset", "rst")
    txtFormat(miT_PREFIX) = GetSetting("SQLS", "Settings", "Prefix", "miT_")
    
    chkFormat(miT_STRING_VAR).Value = 1
    chkFormat(miT_SQL_CHECK).Value = 1
    chkFormat(miT_TEXT_BOX).Value = 1
    chkFormat(miT_SPACES).Value = 1
    chkFormat(miT_RECORDSET).Value = 1
    chkFormat(miT_PREFIX).Value = 1
    
    chkType(miC_INSERT).Value = 1
    chkType(miC_DELETE).Value = 1
    chkType(miC_UPDATE).Value = 1
    chkType(miC_DIM).Value = 1
    rtb.Text = ""
    

    
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    
    SaveFormSettings Me
    SaveSetup
    UnloadAllForms
    End
    
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "Form_Unload")
End Sub


Private Sub SaveSetup()
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim i As Integer, iEmpty As Integer
    Call SaveSetting("SQLS", "Settings", "ConncetionString", txtDB)
    Call SaveSetting("SQLS", "Settings", "DBName", txtDBName)
    Call SaveSetting("SQLS", "Settings", "sSql", txtFormat(miT_STRING_VAR))
    Call SaveSetting("SQLS", "Settings", "SQLCheck", txtFormat(miT_SQL_CHECK))
    Call SaveSetting("SQLS", "Settings", "txtBOX", txtFormat(miT_TEXT_BOX))
    Call SaveSetting("SQLS", "Settings", "Spaces", txtFormat(miT_SPACES))
    Call SaveSetting("SQLS", "Settings", "Recordset", txtFormat(miT_RECORDSET))
    Call SaveSetting("SQLS", "Settings", "Prefix", txtFormat(miT_PREFIX))
    
    iEmpty = -1
    
    For i = 0 To 15
        If cboDB = msDBName(i) Then
            miDBNameNum = i
            Call SaveSetting("SQLS", "Settings", "DBNum", i)
            Exit Sub
        End If
        If msDBName(i) = "" And iEmpty = -1 Then iEmpty = i
    Next i
        
    For i = 15 To 1 Step -1
        msDBName(i) = msDBName(i - 1)
        msDBConnect(i) = msDBConnect(i - 1)
        Call SaveSetting("SQLS", "Settings", "DBName_" & CStr(i), msDBName(i))
        Call SaveSetting("SQLS", "Settings", "DBConnect_" & CStr(i), msDBConnect(i))
    Next i
    i = 0
    msDBName(i) = cboDB
    msDBConnect(i) = txtDB
    Call SaveSetting("SQLS", "Settings", "DBName_" & CStr(i), msDBName(i))
    Call SaveSetting("SQLS", "Settings", "DBConnect_" & CStr(i), msDBConnect(i))
    Call SaveSetting("SQLS", "Settings", "DBNum", i)

ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "SaveSetup")
End Sub
Private Sub lstTables_Click()
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    msTableName = lstTables.Text
ErrExit:      Exit Sub
ErrHandler:   Call ErrorHandler(Name, 0, "lstTables_Click")
End Sub

Private Function SQLFormat(sSql As String) As String
If gbErrorHandSwitch Then On Error GoTo ErrHandler
    Dim s As String, i As Integer
    Dim iStart As Integer, iEnd As Integer
    
    s = sSql
    
    s = Replace(s, " _" & vbCrLf, "")
    s = Replace(s, "_ ", "")
    s = Replace(s, vbCrLf, "")
    s = Replace(s, "& """, "")
    s = Replace(s, """ &", "")
    s = Replace(s, "_ &", "")
    s = Replace(s, """", "")
    s = Replace(s, "Select ", " SELECT ")
    s = Replace(s, "select ", " SELECT ")
    s = Replace(s, Chr(32), " ")
    s = Replace(s, " On ", " ON ")
    s = Replace(s, " on ", " ON ")
    s = Replace(s, " Inner ", " INNER ")
    s = Replace(s, " inner ", " INNER ")
    s = Replace(s, " Join", " JOIN ")
    s = Replace(s, " join", " JOIN ")
    s = Replace(s, " left ", " LEFT ")
    s = Replace(s, " left ", " LEFT ")
    s = Replace(s, " Outer ", " OUTER ")
    s = Replace(s, " outer ", " OUTER ")
    s = Replace(s, " Right ", " RIGHT ")
    s = Replace(s, " right ", " RIGHT ")
    s = Replace(s, " Union ", " UNION ")
    s = Replace(s, " union ", " UNION ")
    s = Replace(s, " From ", " FROM ")
    s = Replace(s, " from ", " FROM ")
    s = Replace(s, " Where ", " WHERE ")
    s = Replace(s, " where ", " WHERE ")
    s = Replace(s, " Order ", " ORDER ")
    s = Replace(s, " order ", " ORDER ")
    s = Replace(s, " Group ", " GROUP ")
    s = Replace(s, " group ", " GROUP ")
    s = Replace(s, " By ", " BY ")
    s = Replace(s, " by ", " BY ")
    s = Replace(s, " Or ", " OR ")
    s = Replace(s, " or ", " OR ")
    s = Replace(s, " And ", " AND ")
    s = Replace(s, " and ", " AND ")
    For i = 1 To 20
        s = Replace(s, Space(2), Space(1))
    Next
    
    If chkPartial = 0 Then
        s = Replace(s, "SELECT ", vbCrLf & " SELECT ")
        s = Replace(s, Chr(32), " ")
        s = Replace(s, " ON ", vbCrLf & " ON ")
    '    s = Replace(s, "on ", vbCrLf & "ON ")
        s = Replace(s, " INNER ", vbCrLf & " INNER ")
'        s = Replace(s, " JOIN ", vbCrLf & " JOIN ")
        s = Replace(s, " LEFT ", vbCrLf & " LEFT ")
'        s = Replace(s, " OUTER ", vbCrLf & " OUTER ")
        s = Replace(s, " RIGHT ", vbCrLf & " RIGHT ")
        s = Replace(s, " UNION  ", vbCrLf & " UNION ")
        s = Replace(s, " FROM ", vbCrLf & " FROM ")
        s = Replace(s, " WHERE ", vbCrLf & " WHERE ")
        s = Replace(s, " ORDER ", vbCrLf & " ORDER ")
        s = Replace(s, " GROUP ", vbCrLf & " GROUP ")
        s = Replace(s, " BY ", vbCrLf & " BY ")
        s = Replace(s, " OR ", vbCrLf & " OR ")
        s = Replace(s, " AND ", vbCrLf & " AND ")
        s = Replace(s, ",", "," & vbCrLf & Space(8))
        s = Replace(s, vbCrLf, """" & vbCrLf & "sSql = sSql & """)
        
        For i = 1 To 10
            s = Replace(s, Space(9), Space(8))
        Next
        s = Replace(s, "sSql = sSql & "" SELECT", "sSql = ""SELECT")
        s = Right(s, Len(s) - 3)
    End If
    
    If chkContinue <> 0 Then
        s = Replace(s, vbCrLf, " _" & vbCrLf)
        s = Replace(s, "sSql = sSql & ", " & ")
    End If
    
    SQLFormat = s
    
ErrExit:      Exit Function
ErrHandler:   Call ErrorHandler(Name, 0, "SQLFormat")
End Function

