VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EReplicator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replicator 1-2-3"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11430
   ForeColor       =   &H00000000&
   Icon            =   "EReplicator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ClearLog 
      Caption         =   "Clear Log"
      Height          =   330
      Left            =   8580
      TabIndex        =   57
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar PrBar 
      Height          =   255
      Left            =   1050
      TabIndex        =   24
      Top             =   7110
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton HideLog 
      Caption         =   "&Hide Log"
      Height          =   330
      Left            =   9840
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Save File"
            Object.ToolTipText     =   "Create a New Session"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open a New Session"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save a Session"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help Me!"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3780
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":030A
               Key             =   "New"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":041C
               Key             =   "Help"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":052E
               Key             =   "Help1"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":0640
               Key             =   "Save1"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":095A
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "EReplicator.frx":0A6C
               Key             =   "Save"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Status 
      Caption         =   "Status"
      Height          =   6390
      Left            =   7710
      TabIndex        =   30
      Top             =   480
      Width           =   3615
      Begin VB.CommandButton Transfer 
         Caption         =   "&Transfer"
         Height          =   330
         Left            =   2130
         TabIndex        =   45
         Top             =   5370
         Width           =   1200
      End
      Begin VB.CommandButton ShowLog 
         Caption         =   "Show &Log"
         Height          =   315
         Left            =   2160
         TabIndex        =   37
         Top             =   5880
         Width           =   1155
      End
      Begin VB.CommandButton ClearStat 
         Caption         =   "Clear &Status"
         Height          =   330
         Left            =   390
         TabIndex        =   36
         Top             =   5370
         Width           =   1200
      End
      Begin VB.ListBox FConn 
         Height          =   4545
         ItemData        =   "EReplicator.frx":0B7E
         Left            =   1890
         List            =   "EReplicator.frx":0B80
         TabIndex        =   32
         Top             =   600
         Width           =   1545
      End
      Begin VB.ListBox SConn 
         Height          =   4545
         ItemData        =   "EReplicator.frx":0B82
         Left            =   180
         List            =   "EReplicator.frx":0B84
         TabIndex        =   31
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Failed Connections"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1920
         TabIndex        =   34
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Successfull connections"
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   150
         TabIndex        =   33
         Top             =   180
         Width           =   1635
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   23
      Top             =   7080
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1765
            TextSave        =   "8:57 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   13220
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1941
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3088
            MinWidth        =   3088
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame ListFile 
      Caption         =   "Hosts List"
      Height          =   6375
      Left            =   3630
      TabIndex        =   9
      Top             =   480
      Width           =   3975
      Begin VB.CommandButton getDom 
         Caption         =   "List Domains"
         Height          =   330
         Left            =   2370
         TabIndex        =   55
         Top             =   2190
         Width           =   1200
      End
      Begin VB.CommandButton AddHosts 
         Caption         =   "Add H&osts"
         Height          =   330
         Left            =   540
         TabIndex        =   53
         Top             =   5850
         Width           =   1200
      End
      Begin VB.TextBox EndN 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   49
         Top             =   5430
         Width           =   495
      End
      Begin VB.TextBox StartN 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   510
         TabIndex        =   48
         Top             =   5430
         Width           =   495
      End
      Begin VB.TextBox Prefix 
         Height          =   285
         Left            =   180
         TabIndex        =   47
         Top             =   4980
         Width           =   1635
      End
      Begin MSComDlg.CommonDialog SaveDialog 
         Left            =   1140
         Top             =   930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Save Session As"
         Flags           =   1
         InitDir         =   "C:\"
      End
      Begin MSComDlg.CommonDialog OpenDialog 
         Left            =   360
         Top             =   900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open Session"
         Flags           =   1
         InitDir         =   "C:\"
      End
      Begin VB.CommandButton Replicate 
         BackColor       =   &H00BED977&
         Caption         =   "&Replicate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2430
         Picture         =   "EReplicator.frx":0B86
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Fire Up!"
         Top             =   5190
         Width           =   1185
      End
      Begin VB.CommandButton Clear 
         Caption         =   "C&lear All"
         Height          =   330
         Left            =   2355
         TabIndex        =   22
         Top             =   1740
         Width           =   1200
      End
      Begin VB.CommandButton Remove 
         Caption         =   "Re&move"
         Height          =   330
         Left            =   2355
         TabIndex        =   21
         Top             =   1290
         Width           =   1200
      End
      Begin VB.CommandButton Browse 
         Caption         =   "Bro&wse..."
         Height          =   330
         Left            =   2355
         TabIndex        =   18
         Top             =   870
         Width           =   1200
      End
      Begin VB.TextBox ComChar 
         Height          =   285
         Left            =   2190
         TabIndex        =   15
         Top             =   3720
         Width           =   1545
      End
      Begin VB.ComboBox DomainList 
         Height          =   315
         ItemData        =   "EReplicator.frx":0E90
         Left            =   2190
         List            =   "EReplicator.frx":0E92
         TabIndex        =   14
         Top             =   2910
         Width           =   1575
      End
      Begin VB.CommandButton Discover 
         Caption         =   "&Discover"
         Height          =   330
         Left            =   2400
         TabIndex        =   13
         Top             =   4170
         Width           =   1200
      End
      Begin VB.CommandButton AddHost 
         Caption         =   "Add &Host"
         Height          =   330
         Left            =   390
         TabIndex        =   12
         Top             =   4140
         Width           =   1200
      End
      Begin VB.TextBox ManualAdd 
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   3720
         Width           =   1635
      End
      Begin VB.ListBox List1 
         Height          =   2400
         ItemData        =   "EReplicator.frx":0E94
         Left            =   165
         List            =   "EReplicator.frx":0E96
         MultiSelect     =   1  'Simple
         TabIndex        =   10
         Top             =   630
         Width           =   1650
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   1005
         Left            =   2280
         Top             =   5070
         Width           =   1485
      End
      Begin VB.Label Label16 
         Caption         =   "To:"
         Height          =   225
         Left            =   1050
         TabIndex        =   52
         Top             =   5460
         Width           =   285
      End
      Begin VB.Label Label15 
         Caption         =   "From:"
         Height          =   195
         Left            =   90
         TabIndex        =   51
         Top             =   5460
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Hosts Prefix"
         Height          =   255
         Left            =   300
         TabIndex        =   50
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Selected Hosts"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Host Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   35
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Hosts List"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   315
         TabIndex        =   28
         Top             =   1125
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Selected"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   930
         TabIndex        =   20
         Top             =   3150
         Width           =   795
      End
      Begin VB.Label HCount 
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   570
         TabIndex        =   19
         Top             =   3150
         Width           =   285
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Host Name Prefix"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2250
         TabIndex        =   17
         Top             =   3450
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Select Domain"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2250
         TabIndex        =   16
         Top             =   2670
         Width           =   1395
      End
      Begin VB.Image Image3 
         Height          =   570
         Left            =   3270
         Picture         =   "EReplicator.frx":0E98
         ToolTipText     =   "Just a pretty button"
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination"
      Height          =   2085
      Left            =   90
      TabIndex        =   1
      Top             =   4770
      Width           =   3435
      Begin VB.CheckBox CheckFolder 
         Caption         =   "Build Path If Not Exist"
         Height          =   285
         Left            =   150
         TabIndex        =   54
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox RShare 
         Height          =   315
         ItemData        =   "EReplicator.frx":15C1
         Left            =   150
         List            =   "EReplicator.frx":15CE
         TabIndex        =   41
         Text            =   "C$\"
         Top             =   510
         Width           =   1485
      End
      Begin VB.TextBox DestFolder 
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1230
         Width           =   2325
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Remote Share"
         Height          =   195
         Left            =   270
         TabIndex        =   42
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Destination Folder Path"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   27
         Top             =   960
         Width           =   1725
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   2760
         Picture         =   "EReplicator.frx":15E6
         ToolTipText     =   "Just a pretty button"
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files to Replicate"
      Height          =   4200
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   3435
      Begin VB.CheckBox CheckFilesOverride 
         Caption         =   "Override Destination Files"
         Height          =   255
         Left            =   150
         TabIndex        =   56
         Top             =   3810
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CommandButton FileRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   2070
         TabIndex        =   7
         Top             =   3420
         Width           =   1200
      End
      Begin VB.ListBox FileList 
         Height          =   1230
         ItemData        =   "EReplicator.frx":1D05
         Left            =   150
         List            =   "EReplicator.frx":1D07
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2100
         Width           =   3105
      End
      Begin VB.CommandButton Add 
         Caption         =   "&Add"
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   1380
         Width           =   1200
      End
      Begin VB.TextBox FLocation 
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   930
         Width           =   3105
      End
      Begin VB.CommandButton allClear 
         Caption         =   "&Clear All"
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   3420
         Width           =   1200
      End
      Begin VB.CommandButton FBrowse 
         Caption         =   "&Browse..."
         Height          =   330
         Left            =   2070
         TabIndex        =   2
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Path to File(s)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   870
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "File(s) to be replicated"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   810
         TabIndex        =   26
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Files to Replicate"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   1830
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   2760
         Picture         =   "EReplicator.frx":1D09
         ToolTipText     =   "Just a pretty button"
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.TextBox Log 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   3630
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   43
      Top             =   570
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label LogPath 
      Height          =   375
      Left            =   3660
      TabIndex        =   46
      Top             =   6330
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuNew 
         Caption         =   "&New Session"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Session"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Session"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Session &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu munBr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Replicator 1-2 3"
      End
   End
End
Attribute VB_Name = "EReplicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'Replicator 1-2-3 v1.2.02   *
'Author:  Gurgen Alaverdian *
'****************************


Private Sub Form_Load()
Dim cName           As String
Dim FileNameOnly    As String
Dim cmdArray        As Variant
Dim AutoStart       As Boolean
Dim AutoShut        As Boolean
Dim endTime         As Single

If Not Command = "" Then
    cmdArray = Split(Command, "/", -1, vbTextCompare)
    For Each cmdOption In cmdArray
        If cmdOption = "" Then
                DoEvents
            ElseIf UCase(Right(Trim(cmdOption), 3)) = "SSS" Then
                FileToOpen = cmdOption
            ElseIf UCase(Trim(cmdOption)) = "S" Then
                AutoStart = True
            ElseIf UCase(Trim(cmdOption)) = "X" Then
                AutoShut = True
                noSound = True
            ElseIf UCase(Trim(cmdOption)) = "N" Then
                CheckFilesOverride.Value = 0
            ElseIf UCase(Trim(cmdOption)) = "P" Then
                CheckFolder.Value = 1
            Else
                frmHelp.Show
                Me.Hide
                Exit Sub
        End If
    Next
End If
ReDim cmdArray(0)
cName = Environ$("COMPUTERNAME")
StatusBar1.Panels(4).Text = "Ready"
If FileToOpen = "" Then
   Me.Caption = "Replicator 1-2-3 (Untitled.sss)": GoTo skipCheck
   ElseIf Not FileExists(FileToOpen) Then
        MsgBox "File ''" & FileToOpen & _
        "'' can not be found! Please check the file name, " & _
        "path and try again.", vbExclamation, "Error:": Unload Me: End
   Else
   FileNameOnly = Right(FileToOpen, Len(FileToOpen) - InStrRev(FileToOpen, "\"))
   Me.Caption = "Replicator 1-2-3 (" & FileNameOnly & ")"
   Call OpenFile(FileToOpen)
End If
skipCheck:
NewLogFile = App.Path & "\replicator.log"
If Not FileExists(NewLogFile) Then
       Open NewLogFile For Output As #1
       Print #1, "                                     Replicator  Log  on   " & _
                 "''" & "\\" & cName & "''"
       Print #1, "============================================================="
       Close #1
End If
If AutoStart Then EReplicator.Show: Call Replicate_Click
If AutoShut Then
endTime = Timer + 3
Do While Timer < endTime
DoEvents
Loop
Unload Me
End
End If
End Sub

Private Sub Add_Click()
Dim AddFile As String

If FLocation.Text = "" Then
   NoFile = MsgBox("Specify path to a file.", vbInformation, "Note")
   Exit Sub
End If
AddFile = Trim(FLocation.Text)
If AddFile <> "" Then
    If InStr(FilePath, "*") <> 0 Then GoTo Cont
    If Not FileExists(AddFile) Then
            MsgBox "File ''" & AddFile & "'' can not be found!", vbOKOnly, "Error:"
            Exit Sub
    End If
Cont:
    FileList.AddItem AddFile
End If
End Sub
Private Sub AddHost_Click()
If ManualAdd.Text = "" Then
    EmptyError = MsgBox("You forgot to type a host name!", vbInformation, "Note")
    Exit Sub
End If
List1.AddItem (Trim(ManualAdd.Text))
HCount.Caption = List1.ListCount
End Sub
Private Sub AddHosts_Click()
Dim StartNumber As Integer
Dim EndNumber As Integer
Dim HostsPrefix As String

On Error GoTo ErrorHandle
If StartN.Text = "" Or EndN.Text = "" Or Prefix.Text = "" Then
    NoEntry = MsgBox("Parameters are missing! Please enter Hosts Prefix, " & _
    "Start and End numbers.", vbExclamation, "Error")
    Exit Sub
End If
StartNumber = Trim(StartN.Text)
EndNumber = Trim(EndN.Text)
HostsPrefix = Trim(Prefix.Text)
For I = StartNumber To EndNumber
    If Len(I) = 2 Then
        List1.AddItem (HostsPrefix & "0" & I)
    ElseIf Len(I) = 1 Then
        List1.AddItem (HostsPrefix & "00" & I)
    Else
        List1.AddItem (HostsPrefix & I)
    End If
Next
HCount.Caption = List1.ListCount
Exit Sub
ErrorHandle:
MsgBox "Entry in 'From' and 'To' must be numeric.", vbInformation, "Note"
End Sub
Private Sub Browse_Click()
FileSelect.OK.Caption = "&Load"
FileSelect.Title = "Select Machines List File."
FileSelect.Show
End Sub
Private Sub FBrowse_Click()
FileSelect.OK.Caption = "&Select"
FileSelect.Title = "Select Files to replicate."
FileSelect.Show
End Sub
Private Sub Clear_Click()
If List1.ListCount = 0 Then
    StopTheBull = MsgBox("Do you see anything in the list " & _
    "box?", vbExclamation, "Do not tamper!")
End If
List1.Clear
ManualAdd.Text = ""
HCount.Caption = "0"
Prefix.Text = ""
StartN = ""
EndN = ""
End Sub
Private Sub allClear_Click()
If FileList.ListCount = 0 Then
    MeError = MsgBox("Do you see anything in the list " & _
    "box?", vbExclamation, "Do not tamper!")
    Exit Sub
End If
FileList.Clear
End Sub
Private Sub getDom_Click()
On Error Resume Next
DomainList.Clear
StatusBar1.Panels(4).Text = "Building Domain List..."
Set Container = GetObject("WinNT:")
For Each eDomain In Container
    DomainList.AddItem (eDomain.Name)
Next
DomainList.Text = DomainList.List(1)
StatusBar1.Panels(4).Text = "Ready"
End Sub
Private Sub Discover_Click()
Dim ListCheck   As Integer
Dim ChrIn       As String
Dim ChrNumber   As Integer
Dim cName       As String

If ComChar.Text = "" Or DomainList.Text = "" Then
   NoChr = MsgBox("You need to enter a prefix for the host names " & _
   "and Domain Name.", vbInformation, "Note:")
   Exit Sub
End If
On Error Resume Next
ListCheck = List1.ListCount
StatusBar1.Panels(4).Text = "Working...."
Set Container = GetObject("WinNT://" & Trim(DomainList.Text))
Container.Filter = Array("Computer")
ChrIn = Trim(ComChar.Text)
ChrIn = LCase(ChrIn)
ChrNumber = Len(ChrIn)
For Each Computer In Container
    cName = Computer.Name
    cName = LCase(cName)
    If Left(cName, ChrNumber) = ChrIn Then
        List1.AddItem cName
    End If
Next
If ListCheck = List1.ListCount Then
    MsgBox "Did not find hosts to match, or domain  ''" & _
    DomainList.Text & "''  is not available.", vbExclamation, "Note:"
End If
StatusBar1.Panels(4).Text = "Ready"
HCount.Caption = List1.ListCount
End Sub
Private Sub FileRemove_Click()
If FileList.ListCount = 0 Then
    MeError = MsgBox("Nothing to remove!", vbExclamation, "Do not tamper!")
    Exit Sub
End If
If FileList.ListIndex <> -1 Then
    FileList.RemoveItem FileList.ListIndex
End If
End Sub
Private Sub Remove_Click()
On Error GoTo YourError
Dim I As Integer
I = 0
Do
    If List1.Selected(I) = True Then
        List1.RemoveItem (I)
    Else
        I = I + 1
    End If
Loop Until I > List1.ListCount - 1
HCount.Caption = List1.ListCount
Exit Sub
YourError:
MeError = MsgBox("Nothing to remove!", vbExclamation, "Do not tamper!")
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "New"
  mnuNew_Click
Case "Open"
  mnuOpen_Click
Case "Save"
  mnuSave_Click
Case "Help"
  mnuHHelp_Click
End Select
End Sub
Private Sub mnuNew_Click()
FLocation.Text = ""
FileList.Clear
DestFolder.Text = ""
List1.Clear
ManualAdd.Text = ""
ComChar.Text = ""
SConn.Clear
FConn.Clear
HCount.Caption = "0"
EReplicator.Caption = "Replicator 1-2-3 (Untitled.sss)"
FileToOpen = ""
End Sub
Private Sub mnuOpen_Click()
OpenDialog.Filter = "Session Files (*.sss)|*.sss"
OpenDialog.FileName = ""
OpenDialog.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
OpenDialog.ShowOpen
If OpenDialog.FileName <> "" Then
    FileToOpen = OpenDialog.FileName
    Call OpenFile(FileToOpen)
    FileNameOnly = Right(FileToOpen, Len(FileToOpen) - InStrRev(FileToOpen, "\"))
    EReplicator.Caption = "Replicator 1-2-3 (" & FileNameOnly & ")"
End If
End Sub
Private Sub mnuSave_Click()
If FileToOpen = "" Then
    Call mnuSaveAs_Click
Else
    Call SaveFile(FileToOpen)
End If
End Sub
Private Sub mnuSaveAs_Click()
Dim Cancel          As Boolean
Dim FileToSave      As String

On Error GoTo ErrorHandler
Cancel = False
With SaveDialog
    .DefaultExt = ".sss"
    .Filter = "Session Files (*.sss)|*.sss"
    .FileName = "Untitled1"
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .ShowSave
End With
If Not Cancel Then
    FileToSave = SaveDialog.FileName
    Call SaveFile(FileToSave)
FileNameOnly = Right(FileToSave, Len(FileToSave) - InStrRev(FileToSave, "\"))
EReplicator.Caption = "Replicator 1-2-3 (" & FileNameOnly & ")"
FileToOpen = FileToSave
End If
ErrorHandler:
If Err.Number = cmdCancel Then
    Cancel = True
    Resume Next
End If
End Sub
Private Sub mnuHHelp_Click()
On Error GoTo ErrorHandle
Shell "rundll32 url.dll,FileProtocolHandler " & App.Path & "\help\help.htm"
Exit Sub
ErrorHandle:
YourError = MsgBox("Help File can not be found!", vbExclamation, "Note")
End Sub
Private Sub mnuAbout_Click()
About.Show
End Sub
Private Sub mnuExit_Click()
End
Unload Me
End Sub
Private Sub Replicate_Click()
Dim I           As Integer
Dim J           As Integer
Dim NextHost    As String
Dim ipResolve   As String
Dim pingRet     As Long
Dim ECHO        As ICMP_ECHO_REPLY
Dim chkFolder   As String
Dim fileOnly    As String
Dim Override    As Boolean
Dim strShare    As String
Dim strFolder   As String
Dim absDest     As String
Dim absPath     As String

On Error Resume Next
If FileList.ListCount = 0 Or List1.ListCount = 0 Then
    NotComplete = MsgBox("You need to specify at list 1 file to replicate, " & _
    "and destination host name.", vbExclamation, "Replicator 1-2-3")
    Exit Sub
End If
Override = False
If CheckFilesOverride.Value = 0 Then Override = True

strShare = Trim(RShare.Text)
strFolder = Trim(DestFolder.Text)
If Not Right(strShare, 1) = "\" Then strShare = strShare & "\"
If Right(strFolder, 1) = "\" Then strFolder = Left(strFolder, Len(strFolder) - 1)

NewLogFile = App.Path & "\replicator.log"
FConn.Clear
SConn.Clear
Open NewLogFile For Append As #1

Print #1,
Print #1, "=================="
Print #1, Now
Print #1, "=================="
Print #1,
Print #1, "Replication of the following file(s):"
Print #1,

J = 0
Do
    Print #1, Trim(FileList.List(J))
    J = J + 1
Loop Until J > FileList.ListCount - 1

Print #1,
Print #1, "to  ''" & RShare.Text & DestFolder.Text & "''   failed on the following hosts:"
Print #1,
PrBar.Min = 0
PrBar.Max = List1.ListCount - 1
DoEvents

I = 0
Do
    NextHost = Trim(List1.List(I))
    StatusBar1.Panels(3).Text = "Remain:  " & List1.ListCount - I
    StatusBar1.Panels(4).Text = "Current: " & NextHost
    DoEvents
    StatusBar1.Refresh
    PrBar.Value = I
    PrBar.Refresh
    Call SocketsInitialize
    ipResolve = GetIPFromHostName(NextHost)
    If Not ipResolve = Empty Then
        pingRet = Ping(ipResolve, "send this", ECHO)
        If pingRet = 0 Then
                    DoEvents
                If Not FileExists("\\" & NextHost & "\" & strShare & strFolder) Then
                    If CheckFolder.Value = 1 Then
                            If Not MakeTree(strFolder, "\\" & NextHost & "\" & strShare) Then
                                FConn.AddItem NextHost
                                Print #1, NextHost & vbTab & vbTab & _
                                "-  Invalid share or, permission denied."
                                GoTo SkipCopy
                            End If
                    Else
                        FConn.AddItem NextHost
                        Print #1, NextHost & vbTab & vbTab & _
                        "-  Dest. folder does not exist, or permission denied."
                        GoTo SkipCopy
                    End If
                End If
            J = 0
            Do
                absPath = FileList.List(J)
                If strFolder = "" Then
                    absDest = "\\" & NextHost & "\" & strShare
                    Else: absDest = "\\" & NextHost & "\" & strShare & strFolder & "\"
                End If
                ReplicateFiles absPath, absDest, Override
                J = J + 1
            Loop Until J > FileList.ListCount - 1
            SConn.AddItem NextHost
        Else
            FConn.AddItem NextHost
            Print #1, NextHost & vbTab & vbTab & "-  " & GetStatusCode(pingRet)
            GoTo SkipCopy
        End If
    Else
        FConn.AddItem NextHost
        Print #1, NextHost & vbTab & vbTab & "-  Can not resolve Host Name to IP address."
        GoTo SkipCopy
    End If
SkipCopy:
    StatusBar1.Refresh
    DoEvents
    I = I + 1
Loop Until I > List1.ListCount - 1
  
StatusBar1.Panels(4).Text = "Ready"
StatusBar1.Panels(3).Text = ""
PrBar.Value = 0
TimeRemainStatus = ""
Close #1
If Not noSound Then Call PlaySound(101)
End Sub
Private Function MakeTree(FolderIn, HostShare)
Dim NextFolder As String
Dim SA As SECURITY_ATTRIBUTES
Dim FolderArray As Variant

FolderArray = Split(FolderIn, "\", -1, vbTextCompare)
NextFolder = Left(HostShare, Len(HostShare) - 1)
For Each SubFolder In FolderArray
    NextFolder = NextFolder & "\" & SubFolder
    If Not FileExists(NextFolder) Then
        createSuccess = CreateDirectory(NextFolder, SA)
    End If
Next
If createSuccess = 0 Then
        MakeTree = False
Else:   MakeTree = True
End If
End Function
Private Sub HideLog_Click()
LogPath.Visible = False
HideLog.Visible = False
Log.Visible = False
ClearLog.Visible = False
ListFile.Visible = True
Status.Visible = True
End Sub
Private Sub ShowLog_Click()
On Error GoTo ErrorHandler
Open NewLogFile For Input As #1
File_Length = LOF(1)
Read_Buffer = Input(File_Length, #1)
LogPath.Caption = NewLogFile
LogPath.Visible = True
HideLog.Visible = True
ClearLog.Visible = True
Log.Visible = True
ListFile.Visible = False
Status.Visible = False
Log.Text = Read_Buffer
Close #1
Exit Sub
ErrorHandler:
LogError = MsgBox("Can not open log file Replicator.log", vbCritical, "Error")
End Sub
Private Sub ClearLog_Click()
Open NewLogFile For Output As #1
Print #1, "                                     Replicator  Log  on   " & _
"''" & "\\" & Environ$("COMPUTERNAME") & "''"
Print #1, "============================================================="
Close #1
Open NewLogFile For Input As #1
File_Length = LOF(1)
Read_Buffer = Input(File_Length, #1)
Log.Text = Read_Buffer
Close #1
End Sub
Private Sub ClearStat_Click()
FConn.Clear
SConn.Clear
End Sub
Private Sub Transfer_Click()
If MsgBox("You are about to Transfer list of failed hosts to the 'Selected hosts' window!", _
    vbOKCancel, "Note") = vbCancel Then Exit Sub
    List1.Clear
    I = 0
Do
    List1.AddItem FConn.List(I), I
    I = I + 1
Loop Until I > FConn.ListCount - 1
FConn.Clear
If List1.List(0) = "" Then List1.Clear
HCount.Caption = List1.ListCount
End Sub
Private Sub FileList_Click()
I = 0
Do
    If FileList.Selected(I) Then
        FileList.ToolTipText = FileList.List(I)
        Exit Do
    End If
    I = I + 1
Loop Until I > FileList.ListCount - 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
Unload Me
End Sub

