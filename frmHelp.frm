VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   4515
   ClientLeft      =   6030
   ClientTop       =   3840
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6465
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6800
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim s As String
    s = App.Path & "\Help.rtf"
    rtb.LoadFile s
End Sub
