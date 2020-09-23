VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   2520
   ClientLeft      =   7140
   ClientTop       =   4965
   ClientWidth     =   3660
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   1845
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   3555
   End
   Begin VB.Label Label2 
      Caption         =   "Syntax: Repl123.exe *.sss_File  [Options]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Command Line Options."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   795
      TabIndex        =   1
      Top             =   360
      Width           =   2085
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
Unload EReplicator
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "Help, /?        -  This Screen." & vbCrLf & _
"*.sss_File      -  Define session file to run." & vbCrLf & _
"[/S]               -  Start Replication." & vbCrLf & _
"[/N]               -  No Override Files." & vbCrLf & _
"[/P]               -  Build Path." & vbCrLf & _
"[/X]               -  Close Replicator. Only with [/S]. "


End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
Unload Me
End Sub
