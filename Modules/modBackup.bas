Attribute VB_Name = "modBackup"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: modBackup
'     Module File: Modules\modBackup.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public Sub BackupFiles()
Dim oPath As String, BackupPath As String, newFile As String, x As Integer, cFile As String, newFile2 As String, cFile2 As String

oPath = Left(theFile, InStrRev(theFile, "\"))
BackupPath = oPath & "BACKUP\"

BackForms:
If frmMain.lstForms.ListItems.Count <= 0 Then GoTo BackMods

    For x = 1 To frmMain.lstForms.ListItems.Count
    
    If Mid(frmMain.lstForms.ListItems(x).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstForms.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstForms.ListItems(x).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstForms.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstForms.ListItems(x).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        cFile = frmMain.Dir1.Path & Mid(frmMain.lstForms.ListItems(x).SubItems(1), InStrRev(frmMain.lstForms.ListItems(x).SubItems(1), "\") + 1)
    Else
        cFile = frmMain.Dir1.Path & "\" & Mid(frmMain.lstForms.ListItems(x).SubItems(1), InStrRev(frmMain.lstForms.ListItems(x).SubItems(1), "\") + 1)
    End If

    newFile = BackupPath & Mid(cFile, InStr(1, cFile, oPath) + Len(oPath))
    newFile2 = Mid(newFile, 1, Len(newFile) - 3) & "frx"
    cFile2 = Mid(cFile, 1, Len(cFile) - 3) & "frx"
    
    If FileExist(Left(newFile, InStrRev(newFile, "\"))) = False Then
        BuildPath (Left(newFile, InStrRev(newFile, "\")))
    End If
    
    Call CopyFile(cFile, newFile, False)
    
    If FileExist(cFile2) = True Then
        Call CopyFile(cFile2, newFile2, False)
    End If

    Next x


BackMods:
If frmMain.lstModules.ListItems.Count <= 0 Then GoTo BackClasses

    For x = 1 To frmMain.lstModules.ListItems.Count
    
    If Mid(frmMain.lstModules.ListItems(x).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstModules.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstModules.ListItems(x).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstModules.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstModules.ListItems(x).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        cFile = frmMain.Dir1.Path & Mid(frmMain.lstModules.ListItems(x).SubItems(1), InStrRev(frmMain.lstModules.ListItems(x).SubItems(1), "\") + 1)
    Else
        cFile = frmMain.Dir1.Path & "\" & Mid(frmMain.lstModules.ListItems(x).SubItems(1), InStrRev(frmMain.lstModules.ListItems(x).SubItems(1), "\") + 1)
    End If

    newFile = BackupPath & Mid(cFile, InStr(1, cFile, oPath) + Len(oPath))
       
    If FileExist(Left(newFile, InStrRev(newFile, "\"))) = False Then
        BuildPath (Left(newFile, InStrRev(newFile, "\")))
    End If
    
    Call CopyFile(cFile, newFile, False)

    Next x


BackClasses:
If frmMain.lstClasses.ListItems.Count <= 0 Then GoTo BackControls

    For x = 1 To frmMain.lstClasses.ListItems.Count
    
    If Mid(frmMain.lstClasses.ListItems(x).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstClasses.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstClasses.ListItems(x).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstClasses.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstClasses.ListItems(x).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        cFile = frmMain.Dir1.Path & Mid(frmMain.lstClasses.ListItems(x).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(x).SubItems(1), "\") + 1)
    Else
        cFile = frmMain.Dir1.Path & "\" & Mid(frmMain.lstClasses.ListItems(x).SubItems(1), InStrRev(frmMain.lstClasses.ListItems(x).SubItems(1), "\") + 1)
    End If

    newFile = BackupPath & Mid(cFile, InStr(1, cFile, oPath) + Len(oPath))
       
    If FileExist(Left(newFile, InStrRev(newFile, "\"))) = False Then
        BuildPath (Left(newFile, InStrRev(newFile, "\")))
    End If
    
    Call CopyFile(cFile, newFile, False)

    Next x

BackControls:
If frmMain.lstControls.ListItems.Count <= 0 Then GoTo BackCustom

    For x = 1 To frmMain.lstControls.ListItems.Count
    
    If Mid(frmMain.lstControls.ListItems(x).SubItems(1), 2, 1) = ":" Then
        frmMain.Dir1.Path = Mid(frmMain.lstControls.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstControls.ListItems(x).SubItems(1), "\"))
    Else
        frmMain.Dir1.Path = Left(theFile, InStrRev(theFile, "\")) & Mid(frmMain.lstControls.ListItems(x).SubItems(1), 1, InStrRev(frmMain.lstControls.ListItems(x).SubItems(1), "\"))
    End If
    
    If Right(frmMain.Dir1.Path, 1) = "\" Then
        cFile = frmMain.Dir1.Path & Mid(frmMain.lstControls.ListItems(x).SubItems(1), InStrRev(frmMain.lstControls.ListItems(x).SubItems(1), "\") + 1)
    Else
        cFile = frmMain.Dir1.Path & "\" & Mid(frmMain.lstControls.ListItems(x).SubItems(1), InStrRev(frmMain.lstControls.ListItems(x).SubItems(1), "\") + 1)
    End If

    newFile = BackupPath & Mid(cFile, InStr(1, cFile, oPath) + Len(oPath))
    
    If FileExist(Left(newFile, InStrRev(newFile, "\"))) = False Then
        BuildPath (Left(newFile, InStrRev(newFile, "\")))
    End If
    
    Call CopyFile(cFile, newFile, False)

    Next x
    
BackCustom:
Dim tExtent As String
Dim theExtent As String
Dim tPos As Integer

If frmMain.lstCustom.ListItems.Count <= 0 Then GoTo BackProject
    
    For x = 1 To frmMain.lstCustom.ListItems.Count
        tExtent = frmMain.lstCustom.ListItems(x).Text
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
            For s = 0 To frmMain.File1.ListCount - 1
                Call CopyFile(oPath & frmMain.File1.List(s), BackupPath & frmMain.File1.List(s), False)
            Next s
        End If

    Next x

BackProject:

Call CopyFile(theFile, BackupPath & Mid(theFile, InStrRev(theFile, "\") + 1), False)

If FileExist(Left(Mid(theFile, InStrRev(theFile, "\") + 1), Len(Mid(theFile, InStrRev(theFile, "\") + 1)) - 3) & "vbw") = True Then
    Call CopyFile(theFile, BackupPath & Left(Mid(theFile, InStrRev(theFile, "\") + 1), Len(Mid(theFile, InStrRev(theFile, "\") + 1)) - 3) & "vbw", False)
End If

frmMain.File1.Pattern = "*.scc"
frmMain.File1.Refresh

If frmMain.File1.ListCount > 0 Then
    For s = 0 To frmMain.File1.ListCount - 1
        Call CopyFile(oPath & frmMain.File1.List(s), BackupPath & frmMain.File1.List(s), False)
    Next s
End If

End Sub

