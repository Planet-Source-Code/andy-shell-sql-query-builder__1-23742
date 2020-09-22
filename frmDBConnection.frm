VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDBConnection 
   Caption         =   "DB Connection"
   ClientHeight    =   3555
   ClientLeft      =   3510
   ClientTop       =   1545
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5595
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmd 
         Caption         =   "Build"
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   16
         ToolTipText     =   "Build Connection String from the Settings"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Exit"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         ToolTipText     =   "Exit the Form"
         Top             =   2880
         Width           =   855
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   240
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other DB Types"
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   5175
         Begin VB.OptionButton opt 
            Caption         =   "SQL Server 2000"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox txtConnect 
            Height          =   285
            Index           =   4
            Left            =   1200
            TabIndex        =   13
            Text            =   "txtConnect"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox txtConnect 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   12
            Text            =   "txtConnect"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txtConnect 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Text            =   "txtConnect"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtConnect 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Text            =   "txtConnect"
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "PWD :"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Login :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Server Name :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "DB Name :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Access DB"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5175
         Begin VB.CommandButton cmdCall 
            Caption         =   "..."
            Height          =   300
            Left            =   4680
            TabIndex        =   4
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtConnect 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   3
            Text            =   "txtConnect"
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "DB Path :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDBPassword As String

Private Const miT_ACCESS = 0
Private Const miT_SERVER = 1
Private Const miT_DB = 2
Private Const miT_LOGIN = 3
Private Const miT_PWD = 4
Private Const miT_TOTAL = 4

Private Sub cmd_Click(Index As Integer)
    Dim i As Integer
    Const iB_EXIT = 0
    Const iB_BUILD = 1
    
    Select Case Index
    Case iB_EXIT:
    Case iB_BUILD:
        If Trim(txtConnect(miT_ACCESS)) <> "" Then
            ' Access DB
            gsDBConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DDD;Persist Security Info=False"
            gsDBConnection = Replace(gsDBConnection, "DDD", txtConnect(miT_ACCESS))
            
        Else
            ' Other DB
            ' Check Fields
            For i = miT_SERVER To miT_TOTAL
                If Trim(txtConnect(i)) = "" Then
                    MsgBox "Please Fill ALL fields and try again"
                    Exit Sub
                End If
            Next
            gsDBConnection = "Provider=SQLOLEDB.1;Password=PPP;Persist Security Info=True;User ID=UUU;Initial Catalog=DDD;Data Source=SSS"
            gsDBConnection = Replace(gsDBConnection, "DDD", txtConnect(miT_DB))
            gsDBConnection = Replace(gsDBConnection, "PPP", txtConnect(miT_PWD))
            gsDBConnection = Replace(gsDBConnection, "UUU", txtConnect(miT_LOGIN))
            gsDBConnection = Replace(gsDBConnection, "SSS", txtConnect(miT_SERVER))
        End If
        frmMain.txtDB = gsDBConnection
        
    End Select
    Unload Me
    
End Sub

Private Sub cmdCall_Click()
    Dim iCheck As Integer
    cd.CancelError = True
    cd.FileName = "*.mdb"
    cd.DialogTitle = "Open Access 97/2000 Database"
    cd.InitDir = ""
        
    cd.Filter = "All Files (*.*)|*.*|Access 97/2000 Database Files (*.mdb)|*.mdb"
    cd.FilterIndex = 2
    cd.ShowOpen
    
    txtConnect(miT_ACCESS) = cd.FileName
End Sub

Private Sub Form_Load()
    LoadFormSettings Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSettings Me

End Sub
