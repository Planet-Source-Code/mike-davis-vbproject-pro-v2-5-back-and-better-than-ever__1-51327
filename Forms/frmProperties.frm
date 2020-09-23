

VERSION 5.00
Begin VB.Form frmProperties
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1960
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1960
      Width           =   855
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Move Path:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmProperties"
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
'     Module Name: frmProperties
'     Module File: Forms\frmProperties.frm
'     Module Type: Form
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public tIndex
Public tType

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim tPath2 As String

tPath2 = Left(theFile, InStrRev(theFile, "\") - 1)

If Right(tPath2, 1) = ":" Then
    tPath2 = tPath2 & "\"
Else
    'Nothing
End If

Select Case tType
Case "Form"
    frmMain.lstForms.ListItems(tIndex).Text = txtName.Text
    frmMain.lstForms.ListItems(tIndex).SubItems(1) = txtFile.Text
    
    If Right(txtPath.Text, 1) = "\" Then
        frmMain.lstForms.ListItems(tIndex).SubItems(2) = txtPath.Text
    Else
        frmMain.lstForms.ListItems(tIndex).SubItems(2) = txtPath.Text & "\"
    End If

    frmMain.lstForms.ListItems(tIndex).Tag = RelativePath(tPath2, frmMain.lstForms.ListItems(tIndex).SubItems(2) & Mid(frmMain.lstForms.ListItems(tIndex).SubItems(1), InStrRev(frmMain.lstForms.ListItems(tIndex).SubItems(1), "\") + 1))

Case "Module"
    frmMain.lstModules.ListItems(tIndex).Text = txtName.Text
    frmMain.lstModules.ListItems(tIndex).SubItems(1) = txtFile.Text
    
    If Right(txtPath.Text, 1) = "\" Then
        frmMain.lstModules.ListItems(tIndex).SubItems(2) = txtPath.Text
    Else
        frmMain.lstModules.ListItems(tIndex).SubItems(2) = txtPath.Text & "\"
    End If

    frmMain.lstModules.ListItems(tIndex).Tag = RelativePath(tPath2, frmMain.lstModules.ListItems(tIndex).SubItems(2) & Mid(frmMain.lstModules.ListItems(tIndex).SubItems(1), InStrRev(frmMain.lstModules.ListItems(tIndex).SubItems(1), "\") + 1))

Case "Class"
    frmMain.lstClasses.ListItems(tIndex).Text = txtName.Text
    frmMain.lstClasses.ListItems(tIndex).SubItems(1) = txtFile.Text
    
    If Right(txtPath.Text, 1) = "\" Then
        frmMain.lstClasses.ListItems(tIndex).SubItems(2) = txtPath.Text
    Else
        frmMain.lstClasses.ListItems(tIndex).SubItems(2) = txtPath.Text & "\"
    End If

    frmMain.lstClasses.ListItems(tIndex).Tag = RelativePath(tPath2, frmMain.lstClasses.ListItems(tIndex).SubItems(2) & Mid(frmMain.lstClasses.ListItems(tIndex).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(tIndex).SubItems(1), "\") + 1))

Case "Control"
    frmMain.lstControls.ListItems(tIndex).Text = txtName.Text
    frmMain.lstControls.ListItems(tIndex).SubItems(1) = txtFile.Text
    
    If Right(txtPath.Text, 1) = "\" Then
        frmMain.lstControls.ListItems(tIndex).SubItems(2) = txtPath.Text
    Else
        frmMain.lstControls.ListItems(tIndex).SubItems(2) = txtPath.Text & "\"
    End If

    frmMain.lstControls.ListItems(tIndex).Tag = RelativePath(tPath2, frmMain.lstControls.ListItems(tIndex).SubItems(2) & Mid(frmMain.lstControls.ListItems(tIndex).SubItems(1), InStrRev(frmMain.lstControls.ListItems(tIndex).SubItems(1), "\") + 1))

End Select


Unload Me
End Sub

Private Sub cmdPath_Click()
Dim tPath As String
tPath = BrowseForFolder(frmRule, "Please select a folder to move " & txtName.Text & " to.", ROOTDIR_ALL, , txtPath.Text, True, False)
If tPath <> "" Then
    If Right(tPath, 1) = "\" Then
        txtPath.Text = tPath
    Else
        txtPath.Text = tPath & "\"
    End If
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon
End Sub
