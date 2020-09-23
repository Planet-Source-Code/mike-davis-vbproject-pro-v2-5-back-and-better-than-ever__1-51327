Attribute VB_Name = "modSave"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: modSave
'     Module File: Modules\modSave.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public Sub MoveForms()
Dim s As Integer
Dim tLine
Dim FRXFile
Dim tPath As String

If frmMain.lstForms.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstForms.ListItems.Count
    frmMain.Caption = "VBProject Pro v2.5 - Status: Moving Form " & s & " of " & frmMain.lstForms.ListItems.Count
    If FileExist(frmMain.lstForms.ListItems(s).SubItems(2)) = False Then
        BuildPath (frmMain.lstForms.ListItems(s).SubItems(2))
    End If
    
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
    
    Call MoveFile(tPath & (Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1)), frmMain.lstForms.ListItems(s).SubItems(2) & Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1))

    FRXFile = Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1)
    
    If FileExist(tPath & Left(FRXFile, Len(FRXFile) - 3) & "frx") = True Then
        Call MoveFile(tPath & Left(FRXFile, Len(FRXFile) - 3) & "frx", frmMain.lstForms.ListItems(s).SubItems(2) & Left(FRXFile, Len(FRXFile) - 3) & "frx")
    End If
    Call UpdateForms(frmMain.lstForms.ListItems(s).SubItems(2) & Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1), s)
    frmMain.pb1.Value = frmMain.pb1.Value + 1
Next s

End Sub
Public Sub MoveModules()
Dim s As Integer
Dim tLine
Dim tPath As String

If frmMain.lstModules.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstModules.ListItems.Count
    frmMain.Caption = "VBProject Pro v2.5 - Status:  Moving Module " & s & " of " & frmMain.lstModules.ListItems.Count
    If FileExist(frmMain.lstModules.ListItems(s).SubItems(2)) = False Then
        BuildPath (frmMain.lstModules.ListItems(s).SubItems(2))
    End If
    
    If Mid(frmMain.lstModules.ListItems(s).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstModules.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstModules.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        tPath = frmMain.Dir1.Path
    Else
        tPath = frmMain.Dir1.Path & "\"
    End If
    
    Call MoveFile(tPath & (Mid(frmMain.lstModules.ListItems(s).SubItems(1), InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\") + 1)), frmMain.lstModules.ListItems(s).SubItems(2) & Mid(frmMain.lstModules.ListItems(s).SubItems(1), InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\") + 1))
    Call UpdateModules(frmMain.lstModules.ListItems(s).SubItems(2) & Mid(frmMain.lstModules.ListItems(s).SubItems(1), InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\") + 1), s)
    frmMain.pb1.Value = frmMain.pb1.Value + 1
Next s

End Sub
Public Sub MoveClasses()
Dim s As Integer
Dim tLine
Dim tPath As String

If frmMain.lstClasses.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstClasses.ListItems.Count
    frmMain.Caption = "VBProject Pro v2.5 - Status:  Moving Class " & s & " of " & frmMain.lstClasses.ListItems.Count
    If FileExist(frmMain.lstClasses.ListItems(s).SubItems(2)) = False Then
        BuildPath (frmMain.lstClasses.ListItems(s).SubItems(2))
    End If
    
    If Mid(frmMain.lstClasses.ListItems(s).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstClasses.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstClasses.ListItems(s).SubItems(1), 1, InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        tPath = frmMain.Dir1.Path
    Else
        tPath = frmMain.Dir1.Path & "\"
    End If
    
    Call MoveFile(tPath & (Mid(frmMain.lstClasses.ListItems(s).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\") + 1)), frmMain.lstClasses.ListItems(s).SubItems(2) & Mid(frmMain.lstClasses.ListItems(s).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\") + 1))
    Call UpdateClasses(frmMain.lstClasses.ListItems(s).SubItems(2) & Mid(frmMain.lstClasses.ListItems(s).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\") + 1), s)
    frmMain.pb1.Value = frmMain.pb1.Value + 1
Next s

End Sub
Public Sub MoveControls()
Dim s As Integer
Dim tLine
Dim tPath As String

If frmMain.lstControls.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstControls.ListItems.Count
    frmMain.Caption = "VBProject Pro v2.5 - Status:  Moving User Control " & s & " of " & frmMain.lstControls.ListItems.Count
    If FileExist(frmMain.lstControls.ListItems(s).SubItems(2)) = False Then
        BuildPath (frmMain.lstControls.ListItems(s).SubItems(2))
    End If

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
    
    Call MoveFile(tPath & (Mid(frmMain.lstControls.ListItems(s).SubItems(1), InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\") + 1)), frmMain.lstControls.ListItems(s).SubItems(2) & Mid(frmMain.lstControls.ListItems(s).SubItems(1), InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\") + 1))
    Call UpdateControls(frmMain.lstControls.ListItems(s).SubItems(2) & Mid(frmMain.lstControls.ListItems(s).SubItems(1), InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\") + 1), s)
    frmMain.pb1.Value = frmMain.pb1.Value + 1
Next s

End Sub
Public Sub MoveCustom()
Dim tExtent As String
Dim theExtent As String
Dim tPos As Integer
Dim tPath As String

If frmMain.lstCustom.ListItems.Count <= 0 Then Exit Sub

For s = 1 To frmMain.lstCustom.ListItems.Count
    frmMain.Caption = "VBProject Pro v2.5 - Status:  Performing CUSTOM Rule " & s & " of " & frmMain.lstCustom.ListItems.Count
    tExtent = frmMain.lstCustom.ListItems(s).Text
    theExtent = ""
    tPos = 1
    Do While tPos <> 0
        If InStr(tPos, tExtent, ",") <> 0 Then
            If theExtent = "" Then
                theExtent = "*." & Mid(tExtent, tPos, InStr(tPos + 1, tExtent, ",") - tPos)
            Else
                theExtent = theExtent & ";*." & Mid(tExtent, tPos, InStr(tPos + 1, tExtent, ",") - tPos)
            End If
            tPos = InStr(tPos, tExtent, ",") + 1
        Else
            If theExtent = "" Then
                theExtent = "*." & Mid(tExtent, tPos)
            Else
                theExtent = theExtent & ";*." & Mid(tExtent, tPos)
            End If
            tPos = 0
        End If
    Loop
    frmMain.File1.Pattern = theExtent
    frmMain.File1.Refresh
    
    If frmMain.File1.ListCount > 0 Then
        For x = 0 To frmMain.File1.ListCount - 1
            If Right(frmMain.File1.Path, 1) = "\" Then
                tPath = frmMain.File1.Path
            Else
                tPath = frmMain.File1.Path & "\"
            End If
            
            If FileExist(frmMain.lstCustom.ListItems(s).SubItems(2)) = False Then
                BuildPath (frmMain.lstCustom.ListItems(s).SubItems(2))
            End If
            
            Call MoveFile(tPath & frmMain.File1.List(x), frmMain.lstCustom.ListItems(s).SubItems(2) & frmMain.File1.List(x))
        Next x
    End If
    frmMain.pb1.Value = frmMain.pb1.Value + 1
Next s
    
End Sub

Public Sub CreateVBP()
Dim tLine
Dim tEntry
Close #1
Open theFile For Input As #1

Do While Not EOF(1)

    If InStr(1, LCase(tLine), "iconform=") <> 0 Then
        frmMain.txtPFile.Text = frmMain.txtPFile.Text & tLine & vbCrLf
        Line Input #1, tLine
    ElseIf InStr(1, LCase(tLine), "form=") <> 0 Then
            
        For s = 1 To frmMain.lstForms.ListItems.Count
            If InStrRev(LCase(tLine), "\") <> 0 Then
                tEntry = Mid(tLine, InStrRev(tLine, "\") + 1)
            Else
                tEntry = Mid(tLine, InStrRev(tLine, "=") + 1)
            End If
            
            If Right(LCase(tLine), Len(tEntry)) = LCase(LCase(Mid(frmMain.lstForms.ListItems(s).SubItems(1), InStrRev(frmMain.lstForms.ListItems(s).SubItems(1), "\") + 1))) Then
                frmMain.txtPFile.Text = frmMain.txtPFile.Text & "Form=" & frmMain.lstForms.ListItems(s).Tag & vbCrLf
            Else
                'Do Nothing
            End If
        Next s
        Line Input #1, tLine
    ElseIf InStr(1, LCase(tLine), "module=") <> 0 Then
        For s = 1 To frmMain.lstModules.ListItems.Count
            If InStrRev(LCase(tLine), "\") <> 0 Then
                tEntry = Mid(tLine, InStrRev(tLine, "\") + 1)
            Else
                tEntry = Mid(tLine, InStrRev(tLine, ";") + 2)
            End If
            
            If Right(LCase(tLine), Len(tEntry)) = LCase(LCase(Mid(frmMain.lstModules.ListItems(s).SubItems(1), InStrRev(frmMain.lstModules.ListItems(s).SubItems(1), "\") + 1))) Then
                frmMain.txtPFile.Text = frmMain.txtPFile.Text & "Module=" & frmMain.lstModules.ListItems(s).Text & "; " & frmMain.lstModules.ListItems(s).Tag & vbCrLf
            Else
                'Do Nothing
            End If
        Next s
        Line Input #1, tLine
    ElseIf InStr(1, LCase(tLine), "class=") <> 0 Then
        For s = 1 To frmMain.lstClasses.ListItems.Count
            If InStrRev(LCase(tLine), "\") <> 0 Then
                tEntry = Mid(tLine, InStrRev(tLine, "\") + 1)
            Else
                tEntry = Mid(tLine, InStrRev(tLine, ";") + 2)
            End If
            
            If Right(LCase(tLine), Len(tEntry)) = LCase(LCase(Mid(frmMain.lstClasses.ListItems(s).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(s).SubItems(1), "\") + 1))) Then
                frmMain.txtPFile.Text = frmMain.txtPFile.Text & "Class=" & frmMain.lstClasses.ListItems(s).Text & "; " & frmMain.lstClasses.ListItems(s).Tag & vbCrLf
            Else
                'Do Nothing
            End If
        Next s
        Line Input #1, tLine
    ElseIf InStr(1, LCase(tLine), "usercontrol=") <> 0 Then
        For s = 1 To frmMain.lstControls.ListItems.Count
            If InStrRev(LCase(tLine), "\") <> 0 Then
                tEntry = Mid(tLine, InStrRev(tLine, "\") + 1)
            Else
                tEntry = Mid(tLine, InStrRev(tLine, "=") + 1)
            End If
            
            If Right(LCase(tLine), Len(tEntry)) = LCase(LCase(Mid(frmMain.lstControls.ListItems(s).SubItems(1), InStrRev(frmMain.lstControls.ListItems(s).SubItems(1), "\") + 1))) Then
                frmMain.txtPFile.Text = frmMain.txtPFile.Text & "UserControl=" & frmMain.lstControls.ListItems(s).Tag & vbCrLf
            Else
                'Do Nothing
            End If
        Next s
        Line Input #1, tLine
    Else
        frmMain.txtPFile.Text = frmMain.txtPFile.Text & tLine & vbCrLf
        Line Input #1, tLine
    End If
Loop
frmMain.txtPFile.Text = frmMain.txtPFile.Text & tLine

Close #1

Kill theFile

Open theFile For Output As #1
Print #1, frmMain.txtPFile.Text
Close #1

frmMain.pb1.Value = frmMain.pb1.Value + 1

End Sub




