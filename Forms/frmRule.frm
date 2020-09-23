

VERSION 5.00
Begin VB.Form frmRule
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caption Here"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbPre 
      Height          =   315
      ItemData        =   "frmRule.frx":0000
      Left            =   1200
      List            =   "frmRule.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtMove 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtExtensions 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Predefined:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2560
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "* Seperate each extension with a comma ',' (eg: txt,doc,rtf)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Extension(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRule"
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
'     Module Name: frmRule
'     Module File: Forms\frmRule.frm
'     Module Type: Form
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public Mode As Boolean
Public tIndex As Integer
Public Sub CheckDefined()
If txtExtensions.Text = "bmp,jpg,jpeg,gif,tif,tiff,png,ico" And txtDescription.Text = "Image Files (bmp,jpg,jpeg,gif,tif,tiff,png,ico)" Then
    cmbPre.Text = "Images"
ElseIf txtExtensions.Text = "txt,rtf,doc,htm,html" And txtDescription.Text = "Document Files (txt,rtf,doc,htm,html)" Then
    cmbPre.Text = "Documents"
ElseIf txtExtensions.Text = "mpg,mpeg,avi,asf,wmf" And txtDescription.Text = "Video Files (mpg,mpeg,avi,wmf)" Then
    cmbPre.Text = "Video"
ElseIf txtExtensions.Text = "wav,mp3,cda,mid,midi,aac,ogg,wma,aiff" And txtDescription.Text = "Audio Files (wav,mp3,cda,mid,midi,aac,ogg,wma,aiff)" Then
    cmbPre.Text = "Audio"
Else
    cmbPre.Text = "Custom"
End If
End Sub
Private Sub cmbPre_Click()
Dim mPath As String
Dim mPos As Integer

If Trim$(txtMove.Text) <> "" Then
    If Right(txtMove.Text, 1) = "\" Then
        mPath = txtMove.Text
    Else
        mPath = txtMove.Text & "\"
    End If
Else
    If Right(App.Path, 1) = "\" Then
        mPath = App.Path
    Else
        mPath = App.Path & "\"
    End If
End If


Select Case cmbPre.Text
Case "Images"
    txtExtensions.Text = "bmp,jpg,jpeg,gif,tif,tiff,png,ico"
    txtDescription.Text = "Image Files (bmp,jpg,jpeg,gif,tif,tiff,png,ico)"
    'txtMove.Text = mPath & "Images\"
Case "Documents"
    txtExtensions.Text = "txt,rtf,doc,htm,html"
    txtDescription.Text = "Document Files (txt,rtf,doc,htm,html)"
    'txtMove.Text = mPath & "Documents\"
Case "Video"
    txtExtensions.Text = "mpg,mpeg,avi,asf,wmf"
    txtDescription.Text = "Video Files (mpg,mpeg,avi,wmf)"
    'txtMove.Text = mPath & "Video\"
Case "Audio"
    txtExtensions.Text = "wav,mp3,cda,mid,midi,aac,ogg,wma,aiff"
    txtDescription.Text = "Audio Files (wav,mp3,cda,mid,midi,aac,ogg,wma,aiff)"
    'txtMove.Text = mPath & "Audio\"
End Select
     
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdMove_Click()
Dim tPath As String
tPath = BrowseForFolder(frmRule, "Please select a folder for the rule:", ROOTDIR_ALL, , Left(theFile, InStrRev(theFile, "\")), True, False)
If tPath <> "" Then
    If Right(tPath, 1) = "\" Then
        txtMove.Text = tPath
    Else
        txtMove.Text = tPath & "\"
    End If
Else
    Exit Sub
End If
End Sub

Private Sub cmdOK_Click()
Dim Complete As Boolean
Dim newRule As ListItem

If Trim$(txtExtensions.Text) <> "" Then
    If Trim$(txtMove.Text) <> "" Then
        Complete = True
    Else
        MsgBox "Please Provide A Move Path", vbCritical
        Exit Sub
        Complete = False
    End If
Else
    MsgBox "Please Provide An Extension(s)", vbCritical
    Exit Sub
    Complete = False
End If

If InStr(1, txtExtensions.Text, ".") <> 0 Then
    MsgBox "Incorrect Extension Format" & Chr(13) & Chr(13) & "Specify Extensions Using This Format (###,###,###)" & Chr(13) & Chr(13) & "*Note: Do Not use periods (.), only the extension"
    Exit Sub
    Complete = False
End If

If Mode = True Then
    Set newRule = frmMain.lstCustom.ListItems.Add(, , Trim$(txtExtensions.Text))
    newRule.SubItems(1) = txtDescription.Text
    If Right(txtMove.Text, 1) = "\" Then
        newRule.SubItems(2) = txtMove.Text
    Else
        newRule.SubItems(2) = txtMove.Text & "\"
    End If
    
Else
    frmMain.lstCustom.ListItems(tIndex).Text = txtExtensions.Text
    frmMain.lstCustom.ListItems(tIndex).SubItems(1) = txtDescription.Text
    If Right(txtMove.Text, 1) = "\" Then
        frmMain.lstCustom.ListItems(tIndex).SubItems(2) = txtMove.Text
    Else
        frmMain.lstCustom.ListItems(tIndex).SubItems(2) = txtMove.Text & "\"
    End If
End If

Unload Me

End Sub

Private Sub Form_Load()
frmRule.Icon = frmMain.Icon
End Sub

Private Sub txtDescription_Change()
CheckDefined
End Sub

Private Sub txtExtensions_Change()
CheckDefined
End Sub

Private Sub txtMove_Change()
CheckDefined
End Sub
