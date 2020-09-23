Attribute VB_Name = "modLoad"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: modLoad
'     Module File: Modules\modLoad.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public theFile As String

Public Sub LoadProject()
Dim tLine
Dim TypeFound As Boolean

    TypeFound = False
    Close #1
    Open theFile For Input As #1
    Input #1, tLine
    Do Until EOF(1)
    If InStr(1, LCase(tLine), "type=") <> 0 Or InStr(1, LCase(tLine), "name=") <> 0 Or InStr(1, LCase(tLine), "startup=") <> 0 Or InStr(1, LCase(tLine), "helpfile=") <> 0 Or InStr(1, LCase(tLine), "description=") <> 0 Then
        
        If InStr(1, LCase(tLine), "type=") <> 0 And TypeFound = False Then
        frmMain.lblProjectType.Caption = "Project Type: " & Mid(tLine, InStr(1, LCase(tLine), "type=") + 5)
        TypeFound = True
        ElseIf InStr(1, LCase(tLine), "name=") <> 0 Then
        frmMain.lblProjectName.Caption = "Project Name: " & Mid(tLine, InStr(1, LCase(tLine), "name=") + 5)
        ElseIf InStr(1, LCase(tLine), "startup=") <> 0 Then
        frmMain.lblProjectStartup.Caption = "Project Startup Object: " & Mid(tLine, InStr(1, LCase(tLine), "startup=") + 8)
        ElseIf InStr(1, LCase(tLine), "helpfile=") <> 0 Then
        frmMain.lblProjectHelp.Caption = "Project Help File: " & Mid(tLine, InStr(1, LCase(tLine), "helpfile=") + 9)
        ElseIf InStr(1, LCase(tLine), "description=") <> 0 Then
        frmMain.lblProjectDescription.Caption = "Project Description: " & Mid(tLine, InStr(1, LCase(tLine), "description=") + 12)
        Exit Do
        End If
    End If
    Line Input #1, tLine
    Loop
    Close #1
    
    Open theFile For Input As #1
    Input #1, tLine
    Do Until EOF(1)
    If InStr(1, LCase(tLine), "majorver=") <> 0 Or InStr(1, LCase(tLine), "minorver=") <> 0 Or InStr(1, LCase(tLine), "revisionver=") <> 0 Or InStr(1, LCase(tLine), "title=") <> 0 Or InStr(1, LCase(tLine), "iconform=") <> 0 Or InStr(1, LCase(tLine), "versioncomments=") <> 0 Or InStr(1, LCase(tLine), "versioncompanyname=") <> 0 Or InStr(1, LCase(tLine), "versionfiledescription=") <> 0 Or InStr(1, LCase(tLine), "versionlegalcopyright=") <> 0 Or InStr(1, LCase(tLine), "versionlegaltrademarks=") <> 0 Or InStr(1, LCase(tLine), "versionproductname=") <> 0 Then
        
        If InStr(1, LCase(tLine), "majorver=") <> 0 Then
        frmMain.txtMajor.Text = Mid(tLine, InStr(1, LCase(tLine), "majorver=") + 9)
        ElseIf InStr(1, LCase(tLine), "minorver=") <> 0 Then
        frmMain.txtMinor.Text = Mid(tLine, InStr(1, LCase(tLine), "minorver=") + 9)
        ElseIf InStr(1, LCase(tLine), "revisionver=") <> 0 Then
        frmMain.txtRevision.Text = Mid(tLine, InStr(1, LCase(tLine), "revisionver=") + 12)
        ElseIf InStr(1, LCase(tLine), "title=") <> 0 Then
        frmMain.lblTitle = "Title: " & Mid(tLine, InStr(1, LCase(tLine), "title=") + 6)
        ElseIf InStr(1, LCase(tLine), "iconform=") <> 0 Then
        frmMain.lblIcon = "Icon Form: " & Mid(tLine, InStr(1, LCase(tLine), "iconform=") + 9)
        ElseIf InStr(1, LCase(tLine), "versioncomments=") <> 0 Then
        frmMain.lblComments = Mid(tLine, InStr(1, LCase(tLine), "versioncomments=") + 16)
        ElseIf InStr(1, LCase(tLine), "versioncompanyname=") <> 0 Then
        frmMain.lblCompany = Mid(tLine, InStr(1, LCase(tLine), "versioncompanyname=") + 19)
        ElseIf InStr(1, LCase(tLine), "versionfiledescription=") <> 0 Then
        frmMain.lblDescription = Mid(tLine, InStr(1, LCase(tLine), "versionfiledescription=") + 23)
        ElseIf InStr(1, LCase(tLine), "versionlegalcopyright=") <> 0 Then
        frmMain.lblCopyright = Mid(tLine, InStr(1, LCase(tLine), "versionlegalcopyright=") + 22)
        ElseIf InStr(1, LCase(tLine), "versionlegaltrademarks=") <> 0 Then
        frmMain.lblTrademarks = Mid(tLine, InStr(1, LCase(tLine), "versionlegaltrademarks=") + 23)
        ElseIf InStr(1, LCase(tLine), "versionproductname=") <> 0 Then
        frmMain.lblProduct = Mid(tLine, InStr(1, LCase(tLine), "versionproductname=") + 19)
        End If
    End If
    Line Input #1, tLine
    Loop
    Close #1
End Sub

Public Sub LoadForms()
Dim s As Integer
Dim tLine
Dim newForm As ListItem
Dim tPath As String, tPath2 As String
Dim tForm As String
Dim Done As Boolean

tPath = Left(theFile, InStrRev(theFile, "\"))

tPath2 = Left(theFile, InStrRev(theFile, "\") - 1)

If Right(tPath2, 1) = ":" Then
    tPath2 = tPath2 & "\"
Else
    'Nothing
End If

Close #1
Open theFile For Input As #1
Input #1, tLine
Do Until EOF(1)
    If InStr(1, LCase(tLine), "iconform=") <> 0 Then
            Line Input #1, tLine
    ElseIf InStr(1, LCase(tLine), "form=") <> 0 Then
        Set newForm = frmMain.lstForms.ListItems.Add(, , "")
        newForm.SubItems(1) = Mid(tLine, InStr(1, LCase(tLine), "form=") + 5)
        newForm.SubItems(2) = tPath & "Forms\"
        newForm.Tag = RelativePath(tPath2, newForm.SubItems(2) & Mid(newForm.SubItems(1), InStrRev(newForm.SubItems(1), "\") + 1))
    End If
    
    Line Input #1, tLine
Loop
Close #1

If frmMain.lstForms.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstForms.ListItems.Count

    If Mid(frmMain.lstForms.ListItems(s).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstForms.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstForms.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        tPath = frmMain.Dir1.Path
    Else
        tPath = frmMain.Dir1.Path & "\"
    End If
    
    
    tForm = tPath & (Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1))
    Done = False
    Close #1
    Open tForm For Input As #1
    Input #1, tLine
    Do While Done = False
        If InStr(1, LCase(tLine), "vb.form") <> 0 Then
                frmMain.lstForms.ListItems(s).Text = Trim$(Mid(tLine, InStr(1, LCase(tLine), "vb.form") + 7))
                Done = True
        End If
        If InStr(1, LCase(tLine), "vb.mdiform") <> 0 Then
                frmMain.lstForms.ListItems(s).Text = Trim$(Mid(tLine, InStr(1, LCase(tLine), "vb.mdiform") + 10))
                Done = True
        End If
        
        Line Input #1, tLine
    Loop
    Close #1
Next s

End Sub

Public Sub LoadModules()
Dim tLine
Dim tPos As Integer
Dim newForm As ListItem
Dim tPath As String, tPath2 As String

tPath = Left(theFile, InStrRev(theFile, "\"))

tPath2 = Left(theFile, InStrRev(theFile, "\") - 1)

If Right(tPath2, 1) = ":" Then
    tPath2 = tPath2 & "\"
Else
    'Nothing
End If

Close #1
Open theFile For Input As #1
Input #1, tLine
Do Until EOF(1)
    If InStr(1, LCase(tLine), "module=") <> 0 Then
        Set newForm = frmMain.lstModules.ListItems.Add(, , Mid(tLine, InStr(1, LCase(tLine), "module=") + 7, (InStr(1, LCase(tLine), ";")) - (InStr(1, LCase(tLine), "module=") + 7)))
        newForm.SubItems(1) = Mid(tLine, InStr(1, LCase(tLine), ";") + 2)
        newForm.SubItems(2) = tPath & "Modules\"
        newForm.Tag = RelativePath(tPath2, newForm.SubItems(2) & Mid(newForm.SubItems(1), InStrRev(newForm.SubItems(1), "\") + 1))
    End If
    
    Line Input #1, tLine
Loop
Close #1

End Sub

Public Sub LoadClasses()
Dim tLine
Dim tPos As Integer
Dim newForm As ListItem
Dim tPath As String, tPath2 As String

tPath = Left(theFile, InStrRev(theFile, "\"))
tPath2 = Left(theFile, InStrRev(theFile, "\") - 1)

If Right(tPath2, 1) = ":" Then
    tPath2 = tPath2 & "\"
Else
    'Nothing
End If

Close #1
Open theFile For Input As #1
Input #1, tLine
Do Until EOF(1)
    If InStr(1, LCase(tLine), "class=") <> 0 Then
        Set newForm = frmMain.lstClasses.ListItems.Add(, , Mid(tLine, InStr(1, LCase(tLine), "class=") + 6, (InStr(1, LCase(tLine), ";")) - (InStr(1, LCase(tLine), "class=") + 6)))
        newForm.SubItems(1) = Mid(tLine, InStr(1, LCase(tLine), ";") + 2)
        newForm.SubItems(2) = tPath & "Classes\"
        newForm.Tag = RelativePath(tPath2, newForm.SubItems(2) & Mid(newForm.SubItems(1), InStrRev(newForm.SubItems(1), "\") + 1))
    End If
    
    Line Input #1, tLine
Loop
Close #1

End Sub

Public Sub LoadControls()
Dim tLine
Dim newForm As ListItem
Dim tPath As String, tPath2 As String

tPath = Left(theFile, InStrRev(theFile, "\"))

tPath2 = Left(theFile, InStrRev(theFile, "\") - 1)

If Right(tPath2, 1) = ":" Then
    tPath2 = tPath2 & "\"
Else
    'Nothing
End If

Close #1
Open theFile For Input As #1
Input #1, tLine
Do Until EOF(1)
    If InStr(1, LCase(tLine), "usercontrol=") <> 0 Then
        Set newForm = frmMain.lstControls.ListItems.Add(, , Mid(tLine, InStr(1, LCase(tLine), "usercontrol=") + 12))
        newForm.SubItems(1) = Mid(tLine, InStr(1, LCase(tLine), "usercontrol=") + 12)
        newForm.SubItems(2) = tPath & "User Controls\"
        newForm.Tag = RelativePath(tPath2, newForm.SubItems(2) & Mid(newForm.SubItems(1), InStrRev(newForm.SubItems(1), "\") + 1))
    End If
    
    Line Input #1, tLine
Loop
Close #1

If frmMain.lstControls.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstControls.ListItems.Count

    If Mid(frmMain.lstControls.ListItems(s).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstControls.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstControls.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        tPath = frmMain.Dir1.Path
    Else
        tPath = frmMain.Dir1.Path & "\"
    End If
    
    
    tForm = tPath & (Mid(frmMain.lstControls.ListItems(s).SubItems(1), InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\") + 1))
    Done = False
    Close #1
    Open tForm For Input As #1
    Input #1, tLine
    Do While Done = False
        If InStr(1, LCase(tLine), "vb.usercontrol") <> 0 Then
                frmMain.lstControls.ListItems(s).Text = Trim$(Mid(tLine, InStr(1, LCase(tLine), "vb.usercontrol") + 14))
                Done = True
        End If
        
        Line Input #1, tLine
    Loop
    Close #1
Next s


End Sub
