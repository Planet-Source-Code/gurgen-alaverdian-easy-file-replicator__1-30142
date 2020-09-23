VERSION 5.00
Begin VB.Form FileSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse...."
   ClientHeight    =   6240
   ClientLeft      =   9885
   ClientTop       =   2595
   ClientWidth     =   5220
   Icon            =   "FileSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5220
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5025
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FileSelect.frx":030A
         Left            =   150
         List            =   "FileSelect.frx":031D
         TabIndex        =   7
         Text            =   "*.*"
         Top             =   5070
         Width           =   1695
      End
      Begin VB.CommandButton Close 
         Caption         =   "&Close"
         Height          =   345
         Left            =   3480
         TabIndex        =   5
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton OK 
         Height          =   345
         Left            =   300
         TabIndex        =   4
         Top             =   5640
         Width           =   1335
      End
      Begin VB.FileListBox File1 
         Height          =   4185
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   720
         Width           =   1755
      End
      Begin VB.DirListBox Dir1 
         Height          =   4140
         Left            =   2040
         TabIndex        =   2
         Top             =   750
         Width           =   2865
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   5070
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Folders"
         Height          =   165
         Left            =   2670
         TabIndex        =   9
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Files"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Title 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub Combo1_Change()
File1.Pattern = Combo1.Text
End Sub
Private Sub Combo1_Click()
File1.Pattern = Combo1.Text
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
ChDir Dir1.Path
End Sub
Private Sub Drive1_Change()
On Error GoTo YourError
Dir1.Path = Drive1.Drive
ChDrive Drive1.Drive
Exit Sub
YourError:
MyError = MsgBox("Drive is not available", vbExclamation, "Off Line")
Drive1.Drive = "C:"
End Sub
Private Sub File1_DblClick()
Call FileSelector
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
End Sub

Private Sub OK_Click()
Call FileSelector
End Sub
Function FileSelector()
Dim I           As Integer
Dim L           As Integer
Dim ItemtoAdd   As String
Dim fPath       As String
Dim hList       As String
Dim rLine       As String
L = 0
If FileSelect.Title = "Select Files to replicate." Then
    For I = 0 To File1.ListCount - 1
        If File1.Selected(I) = True Then
            ItemtoAdd = File1.List(I)
            If Not FileExists(ItemtoAdd) Then
                NotFound = MsgBox("File does not exist!", vbOKOnly, "Error:")
                Exit Function
            End If
            fPath = Dir1.Path
            If Len(fPath) = 3 Then fPath = Left(fPath, 2)
            EReplicator.FileList.AddItem fPath & "\" & ItemtoAdd
        End If
    Next
Else
    hList = File1.FileName
    Open hList For Input As #2
    Do While Not EOF(2)
        Line Input #2, rLine
        EReplicator.List1.AddItem Trim(rLine)
    Loop
    Close #2
    Do
        If EReplicator.List1.List(L) = "" Then
            EReplicator.List1.RemoveItem (L)
        Else
            L = L + 1
        End If
    Loop Until L > EReplicator.List1.ListCount - 1
    EReplicator.HCount.Caption = EReplicator.List1.ListCount
End If
End Function
