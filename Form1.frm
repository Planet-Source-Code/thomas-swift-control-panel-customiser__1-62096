VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TAS System Menu Customiser Version 1.0"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3855
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3150
      Width           =   930
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8100
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2940
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "&Add"
      TabPicture(0)   =   "Form1.frx":39B2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelFile"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelUUID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ComInstall"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SelFile"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TextName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ToolTipText"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Remove"
      TabPicture(1)   =   "Form1.frx":39CE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelOldUUID"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "OldFile"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FileBackups"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ComRemove"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame2 
         Caption         =   "Appears In or On"
         Height          =   585
         Left            =   -70965
         TabIndex        =   15
         Top             =   1155
         Width           =   4140
         Begin VB.Label InMyComputer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "My Computer"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label InControlPanel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Control Panel"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1530
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label OnDesktop 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desktop"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   195
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Add Shortcut To"
         Height          =   555
         Left            =   2347
         TabIndex        =   11
         Top             =   1935
         Width           =   3870
         Begin VB.CheckBox CheckMyComputer 
            Caption         =   "My Computer"
            Height          =   270
            Left            =   2550
            TabIndex        =   14
            Top             =   180
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox CheckControl 
            Caption         =   "Control Panel"
            Height          =   270
            Left            =   1170
            TabIndex        =   13
            Top             =   180
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox CheckAddDesk 
            Caption         =   "&Desktop"
            Height          =   270
            Left            =   105
            TabIndex        =   12
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.CommandButton ComRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   -70185
         TabIndex        =   10
         Top             =   1920
         Width           =   2550
      End
      Begin VB.TextBox ToolTipText 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   825
         Width           =   8370
      End
      Begin VB.TextBox TextName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2445
         TabIndex        =   8
         Top             =   465
         Width           =   3630
      End
      Begin VB.FileListBox FileBackups 
         Height          =   2040
         Left            =   -74850
         Pattern         =   "*.UUID"
         TabIndex        =   5
         Top             =   375
         Width           =   3615
      End
      Begin VB.CommandButton SelFile 
         Caption         =   "&File"
         Height          =   285
         Left            =   7665
         TabIndex        =   2
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton ComInstall 
         Caption         =   "&Install"
         Height          =   300
         Left            =   3120
         TabIndex        =   1
         Top             =   2535
         Width           =   2325
      End
      Begin VB.Label OldFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Old File Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71070
         TabIndex        =   7
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label LabelOldUUID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Old UUID"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71055
         TabIndex        =   6
         Top             =   405
         Width           =   4455
      End
      Begin VB.Label LabelUUID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UUID"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   165
         TabIndex        =   4
         Top             =   1680
         Width           =   8235
      End
      Begin VB.Label LabelFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   1275
         Width           =   7500
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckAddDesk_Click()
Text1.SetFocus
End Sub

Private Sub CheckControl_Click()
Text1.SetFocus
End Sub

Private Sub CheckMyComputer_Click()
Text1.SetFocus
End Sub

Private Sub ComInstall_Click()
Text1.SetFocus
If CheckAddDesk.Value = 0 And CheckControl.Value = 0 And CheckMyComputer.Value = 0 Then
MsgBox "You must select the places where you want the shortcut to appear by checking one of the boxes under 'Add Shortcut To' !"
Exit Sub
End If
ComInstall.Enabled = False
LabelUUID.Caption = GenerateUUID
CreateEntryToSystemPanels LabelUUID.Caption, TextName.Text, ToolTipText.Text, LabelFile.Caption & ",0", LabelFile.Caption
SaveFile
End Sub
Private Sub ComRemove_Click()
Text1.SetFocus
DeleteEntryFromSystemPanels LabelOldUUID
Kill App.Path & "\Backups\" & FileBackups.FileName
FileBackups.Refresh
If FileBackups.ListCount > -1 Then
FileBackups.Selected(0) = True
ComRemove.Enabled = True
Else
ComRemove.Enabled = False
End If
End Sub
Private Sub FileBackups_Click()
LoadFile App.Path & "\Backups\" & FileBackups.FileName
End Sub
Private Sub Form_Load()
ComInstall.Enabled = False
FileBackups.Path = App.Path & "\Backups"
SSTab1.Tab = 0
End Sub
Private Sub SelFile_Click()
Text1.SetFocus
On Error GoTo ErrorTrap
With CommonDialog1
     .DefaultExt = "*.*)"
     .InitDir = "c:\"
     .ShowOpen
End With
LabelFile.Caption = CommonDialog1.FileName
LabelUUID.Caption = ""
TextName.Text = Left(GetFileName(LabelFile.Caption), Len(GetFileName(LabelFile.Caption)) - Len(GetFileExtension(LabelFile.Caption)) - 1)
ToolTipText.Text = TextName.Text
ComInstall.Enabled = True
ErrorTrap:
End Sub
Private Sub SaveFile()
TheFile = FreeFile()
Open App.Path & "\Backups\" & LabelUUID.Caption & ".UUID" For Output As #TheFile
Print #TheFile, TextName.Text
Print #TheFile, ToolTipText.Text
Print #TheFile, LabelFile.Caption
Print #TheFile, LabelUUID.Caption
Print #TheFile, CheckAddDesk.Value
Print #TheFile, CheckControl.Value
Print #TheFile, CheckMyComputer.Value
Close #TheFile
End Sub
Private Function LoadFile(xPath As String)
Dim TempStr As String
On Error Resume Next
TheFile = FreeFile()
Open xPath For Input As #TheFile
Line Input #TheFile, TempStr
'ToolTipText.Text = TempStr
Line Input #TheFile, TempStr
'ToolTipText.Text = TempStr
Line Input #TheFile, TempStr
OldFile.Caption = TempStr
Line Input #TheFile, TempStr
LabelOldUUID.Caption = TempStr

Line Input #TheFile, TempStr
If TempStr = 1 Then
OnDesktop.BackColor = &H80FF80
Else
OnDesktop.BackColor = &HE0E0E0
End If
Line Input #TheFile, TempStr
If TempStr = 1 Then
InControlPanel.BackColor = &H80FF80
Else
InControlPanel.BackColor = &HE0E0E0
End If
Line Input #TheFile, TempStr
If TempStr = 1 Then
InMyComputer.BackColor = &H80FF80
Else
InMyComputer.BackColor = &HE0E0E0
End If
Close #TheFile
End Function
Public Property Get GetFileName(file As String) As String
    Dim m As Long
    Dim GetChr0 As String
    Dim GetChr1 As String
    GetFileName = ""
    For m = 1 To Len(file)
        GetChr0 = Right(file, m)
        GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then
            GetFileName = Right(GetChr0, m - 1)
            GetFileName = Replace(GetFileName, "^.***", "")
            Exit Property
        End If
    Next m
End Property
Public Property Get GetFileExtension(FileName As String) As String
    On Error Resume Next
    Dim TempStr As String
    GetFileExtension = ""
    TempStr = Right(FileName, 2)
    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Property
    Else
        TempStr = Right(FileName, 3)
        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Property
        Else
            TempStr = Right(FileName, 4)
            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Property
            Else
                TempStr = Right(FileName, 5)
                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Property
                Else
                    GetFileExtension = "000"
                End If
            End If
        End If
    End If
End Property
Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Text1.SetFocus
If SSTab1.Tab = 1 Then
FileBackups.Refresh
If FileBackups.ListCount > -1 Then
FileBackups.Selected(0) = True
ComRemove.Enabled = True
Else
ComRemove.Enabled = False
End If
End If
End Sub

