VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBProject Pro v2.5"
   ClientHeight    =   3720
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMove 
      Height          =   1455
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   5880
      Width           =   5415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   2040
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   1815
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtPFile 
      Height          =   1455
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4080
      Width           =   5415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO!"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   3240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Project"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameGen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameMake"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Forms"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstForms"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Modules"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstModules"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Classes"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lstClasses"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Controls"
      TabPicture(4)   =   "frmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lstControls"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "CUSTOM"
      TabPicture(5)   =   "frmMain.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstCustom"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdNew"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdEdit"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdDelete"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Comments"
      TabPicture(6)   =   "frmMain.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame2"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Something Else Here..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -72840
         TabIndex        =   50
         Top             =   480
         Width           =   6735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   1695
         Begin VB.CheckBox chkCControls 
            Caption         =   "User Controls"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox chkCClasses 
            Caption         =   "Classes"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox chkCModules 
            Caption         =   "Modules"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chkCForms 
            Caption         =   "Forms"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frameMake 
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   4320
         TabIndex        =   31
         Top             =   480
         Width           =   4575
         Begin VB.Frame frameVersion 
            Caption         =   "Version Number"
            Height          =   855
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txtRevision 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   41
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtMinor 
               Enabled         =   0   'False
               Height          =   285
               Left            =   720
               TabIndex        =   40
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtMajor 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblRevision 
               AutoSize        =   -1  'True
               Caption         =   "Revision:"
               Height          =   195
               Left            =   1320
               TabIndex        =   44
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lblMinor 
               AutoSize        =   -1  'True
               Caption         =   "Minor:"
               Height          =   195
               Left            =   720
               TabIndex        =   43
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lblMajor 
               AutoSize        =   -1  'True
               Caption         =   "Major:"
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   435
            End
         End
         Begin VB.Frame frameApp 
            Caption         =   "Application"
            Height          =   855
            Left            =   2280
            TabIndex        =   35
            Top             =   240
            Width           =   2175
            Begin VB.Label lblIcon 
               Caption         =   "lblIcon"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   540
               Width           =   1935
            End
            Begin VB.Label lblTitle 
               Caption         =   "lblTitle"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame frameInfo 
            Caption         =   "Version Information"
            Height          =   975
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   4335
            Begin VB.TextBox txtValue 
               Height          =   640
               Left            =   1920
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Top             =   240
               Width           =   2295
            End
            Begin VB.ListBox lstType 
               Height          =   640
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   1695
            End
         End
      End
      Begin VB.Frame frameGen 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   3975
         Begin VB.Label lblProjectType 
            Caption         =   "lblProjectType"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lblProjectName 
            Caption         =   "lblProjectName"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   675
            Width           =   3735
         End
         Begin VB.Label lblProjectStartup 
            Caption         =   "lblProjectStartup"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1020
            Width           =   3735
         End
         Begin VB.Label lblProjectHelp 
            Caption         =   "lblProjectHelp"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1350
            Width           =   3735
         End
         Begin VB.Label lblProjectDescription 
            Caption         =   "lblProjectDescription"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Rule"
         Height          =   375
         Left            =   -72480
         TabIndex        =   4
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Rule"
         Height          =   375
         Left            =   -73680
         TabIndex        =   3
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Rule"
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
      End
      Begin ComctlLib.ListView lstCustom 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Extension"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   5106
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Move Path"
            Object.Width           =   7742
         EndProperty
      End
      Begin ComctlLib.ListView lstForms 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   3343
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Move Path"
            Object.Width           =   7742
         EndProperty
      End
      Begin ComctlLib.ListView lstModules 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   3343
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Move Path"
            Object.Width           =   7742
         EndProperty
      End
      Begin ComctlLib.ListView lstClasses 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   3343
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Move Path"
            Object.Width           =   7742
         EndProperty
      End
      Begin ComctlLib.ListView lstControls 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File"
            Object.Width           =   3343
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Move Path"
            Object.Width           =   7742
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "*Double Click To Edit Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   2715
         Width           =   2760
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Double Click To Edit Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   10
         Top             =   2715
         Width           =   2760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "*Double Click To Edit Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   8
         Top             =   2715
         Width           =   2760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "*Double Click To Edit Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   6
         Top             =   2720
         Width           =   2760
      End
   End
   Begin VB.Label lblCompany 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label lblProduct 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lblTrademarks 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label lblCopyright 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label lblDescription 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label lblComments 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuhelpmain 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: frmMain
'     Module File: Forms\frmMain.frm
'     Module Type: Form
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================


Public Sub LoadDefaults()
cmdGo.Enabled = False
Me.Caption = "VBProject Pro v2.5"

'=== Load Project Defaults === '
lblProjectType.Caption = ""
lblProjectName.Caption = ""
lblProjectStartup.Caption = ""
lblProjectHelp.Caption = ""
lblProjectDescription = ""
txtMajor.Text = ""
txtMinor.Text = ""
txtRevision.Text = ""
lstType.Clear
txtValue.Text = ""
lblTitle.Caption = ""
lblIcon.Caption = ""
txtPFile.Text = ""

'=== Load Forms Defaults === '
lstForms.ListItems.Clear

'=== Load Modules Defaults === '
lstModules.ListItems.Clear

'=== Load Classes Defaults === '
lstClasses.ListItems.Clear

'=== Load Controls Defaults === '
lstControls.ListItems.Clear

'=== Load CUSTOM Defaults === '
lstCustom.ListItems.Clear
cmdNew.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False

End Sub

Private Sub cmdDelete_Click()
If lstCustom.ListItems.Count <= 0 Then Exit Sub

For x = lstCustom.ListItems.Count To 1 Step -1
    If lstCustom.ListItems(x).Selected = True Then
        lstCustom.ListItems.Remove (x)
    End If
Next x

End Sub

Private Sub cmdEdit_Click()
If lstCustom.ListItems.Count <= 0 Then Exit Sub

For x = 1 To lstCustom.ListItems.Count
    If lstCustom.ListItems(x).Selected = True Then
        frmRule.Mode = False
        frmRule.Caption = "Edit Rule"
        frmRule.txtExtensions.Text = lstCustom.ListItems(x).Text
        frmRule.txtDescription.Text = lstCustom.ListItems(x).SubItems(1)
        frmRule.txtMove.Text = lstCustom.ListItems(x).SubItems(2)
        frmRule.tIndex = x
        frmRule.cmdOK.Caption = "&Save"
        frmRule.Show vbModal, Me
        Exit For
    End If
Next x

End Sub

Private Sub cmdGo_Click()
On Error GoTo errHandler
    cmdGo.Enabled = False
    pb1.Value = 0
    pb1.Max = (lstForms.ListItems.Count + lstModules.ListItems.Count + lstClasses.ListItems.Count + lstControls.ListItems.Count + lstCustom.ListItems.Count) + 1
    
    If MsgBox("Would you like to backup any files being moved or changed?", vbYesNo + vbQuestion) = vbYes Then
        Me.Caption = "VBProject Pro v2.5 - Backing Up Files"
        BackupFiles
    End If
    
    pb1.Visible = True
    MoveForms
    MoveModules
    MoveClasses
    MoveControls
    MoveCustom
    Me.Caption = "VBProject Pro v2.5 - Creating VBP File"
    CreateVBP

MsgBox "VBProject Pro Completed Successfully", vbInformation
LoadDefaults
pb1.Visible = False
Exit Sub

errHandler:
If Err Then
    MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical
    pb1.Visible = False
    Me.MousePointer = vbDefault
    cmdGo.Enabled = True
    Exit Sub
End If

End Sub

Private Sub cmdNew_Click()
frmRule.Mode = True
frmRule.Caption = "Create New Rule"
frmRule.Show vbModal, Me
End Sub


Private Sub Form_Load()
LoadDefaults
InitMainForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lstClasses_DblClick()
If lstClasses.ListItems.Count <= 0 Then Exit Sub

For s = 1 To lstClasses.ListItems.Count
    If lstClasses.ListItems(s).Selected = True Then
        frmProperties.tIndex = s
        frmProperties.tType = "Class"
        frmProperties.txtName = lstClasses.ListItems(s).Text
        frmProperties.txtFile = lstClasses.ListItems(s).SubItems(1)
        frmProperties.txtPath = lstClasses.ListItems(s).SubItems(2)
        frmProperties.Caption = "Edit " & lstClasses.ListItems(s).Text & " Properties"
        frmProperties.Show vbModal, Me
        Exit For
    End If
Next s
End Sub

Private Sub lstControls_DblClick()
If lstControls.ListItems.Count <= 0 Then Exit Sub

For s = 1 To lstControls.ListItems.Count
    If lstControls.ListItems(s).Selected = True Then
        frmProperties.tIndex = s
        frmProperties.tType = "Control"
        frmProperties.txtName = lstControls.ListItems(s).Text
        frmProperties.txtFile = lstControls.ListItems(s).SubItems(1)
        frmProperties.txtPath = lstControls.ListItems(s).SubItems(2)
        frmProperties.Caption = "Edit " & lstControls.ListItems(s).Text & " Properties"
        frmProperties.Show vbModal, Me
        Exit For
    End If
Next s
End Sub

Private Sub lstForms_DblClick()
If lstForms.ListItems.Count <= 0 Then Exit Sub

For s = 1 To lstForms.ListItems.Count
    If lstForms.ListItems(s).Selected = True Then
        frmProperties.tIndex = s
        frmProperties.tType = "Form"
        frmProperties.txtName = lstForms.ListItems(s).Text
        frmProperties.txtFile = lstForms.ListItems(s).SubItems(1)
        frmProperties.txtPath = lstForms.ListItems(s).SubItems(2)
        frmProperties.Caption = "Edit " & lstForms.ListItems(s).Text & " Properties"
        frmProperties.Show vbModal, Me
        Exit For
    End If
Next s

End Sub

Private Sub lstModules_DblClick()
If lstModules.ListItems.Count <= 0 Then Exit Sub

For s = 1 To lstModules.ListItems.Count
    If lstModules.ListItems(s).Selected = True Then
        frmProperties.tIndex = s
        frmProperties.tType = "Module"
        frmProperties.txtName = lstModules.ListItems(s).Text
        frmProperties.txtFile = lstModules.ListItems(s).SubItems(1)
        frmProperties.txtPath = lstModules.ListItems(s).SubItems(2)
        frmProperties.Caption = "Edit " & lstModules.ListItems(s).Text & " Properties"
        frmProperties.Show vbModal, Me
        Exit For
    End If
Next s
End Sub

Private Sub lstType_Click()
Select Case lstType.ListIndex
Case 0
    txtValue.Text = lblComments.Caption
Case 1
    txtValue.Text = lblCompany.Caption
Case 2
    txtValue.Text = lblDescription.Caption
Case 3
    txtValue.Text = lblCopyright.Caption
Case 4
    txtValue.Text = lblTrademarks.Caption
Case 5
    txtValue.Text = lblProduct.Caption
End Select

End Sub

Private Sub mnuabout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
On Error GoTo errHandler
cd1.Filter = "Visual Basic Projects (*.vbp) | *.vbp"
cd1.CancelError = True
cd1.ShowOpen

If Trim$(cd1.FileName) <> "" Then
    theFile = cd1.FileName
Else
    Exit Sub
End If

LoadDefaults
frmMain.MousePointer = vbHourglass
LoadProject
LoadForms
LoadModules
LoadClasses
LoadControls
frmMain.MousePointer = vbDefault
lstType.AddItem "Comments"
lstType.AddItem "Company Name"
lstType.AddItem "File Description"
lstType.AddItem "Legal Copyright"
lstType.AddItem "Legal Trademarks"
lstType.AddItem "Product Name"
cmdNew.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdGo.Enabled = True
File1.Path = Left(theFile, InStrRev(theFile, "\"))
Me.Caption = Me.Caption & " - " & theFile
errHandler:
If Err.Number = 32755 Then
    Exit Sub
ElseIf Err Then
    MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical
End If
End Sub


Private Sub SSTab1_DblClick()
frmRule.Caption = "Edit Rule"
End Sub
