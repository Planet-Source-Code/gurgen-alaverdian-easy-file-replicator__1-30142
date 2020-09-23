VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   6060
   ClientTop       =   4725
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   795
      Picture         =   "About.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   315
      Width           =   3045
   End
   Begin VB.Label Label4 
      BackColor       =   &H00004000&
      Caption         =   "Licence:        Free"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   772
      TabIndex        =   3
      Top             =   1800
      Width           =   3090
   End
   Begin VB.Label Label3 
      BackColor       =   &H00004000&
      Caption         =   "Version:        1.2.02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   772
      TabIndex        =   2
      Top             =   1530
      Width           =   3090
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      Height          =   2025
      Left            =   90
      Top             =   90
      Width           =   4425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "Author:         Gurgen Alaverdian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   772
      TabIndex        =   1
      Top             =   1245
      Width           =   3090
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Label1_Click()
Shell "rundll32 url.dll,FileProtocolHandler mailto:gurgen@bellatlantic.net"
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub
