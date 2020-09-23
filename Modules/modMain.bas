

Attribute VB_Name = "modMain"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: modMain
'     Module File: Modules\modMain.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Sub InitMainForm()
Call EnhListView_Add_FullRowSelect(frmMain.lstForms, False)
Call EnhListView_Add_GridLines(frmMain.lstForms, False)
Call EnhListView_Add_FullRowSelect(frmMain.lstModules, False)
Call EnhListView_Add_GridLines(frmMain.lstModules, False)
Call EnhListView_Add_FullRowSelect(frmMain.lstClasses, False)
Call EnhListView_Add_GridLines(frmMain.lstClasses, False)
Call EnhListView_Add_FullRowSelect(frmMain.lstControls, False)
Call EnhListView_Add_GridLines(frmMain.lstControls, False)
Call EnhListView_Add_FullRowSelect(frmMain.lstCustom, False)
Call EnhListView_Add_GridLines(frmMain.lstCustom, False)
End Sub

Public Function FileExist(strFile As String) As Boolean
    If PathFileExists(strFile) = 1 Then
        FileExist = True
    ElseIf PathFileExists(strFile) = 0 Then
        FileExist = False
    End If
End Function

'ALL CREDITS FOR THIS FUNCTION BELONG TO Abdalla Mahmoud

Function BuildPath(ByVal Path As String) As Boolean
    On Error Resume Next
    Dim Fnd As Long
    Dim Tmp As String
    Dim FileSystemObj As Object
    Set FileSystemObj = CreateObject("Scripting.FileSystemObject")
    If FileSystemObj.DriveExists(FileSystemObj.GetDriveName(Path)) = False Then Exit Function
    Path = Path & IIf(Right(Path, 1) = "\", vbNullString, "\")
    Fnd = InStr(Path, "\")


    Do While Fnd
        Tmp = Tmp & Left(Path, Fnd)
        Path = Mid(Path, Fnd + 1)
        MkDir Tmp
        If FileSystemObj.DriveExists(Tmp) = False And FileSystemObj.FolderExists(Tmp) = False Then Exit Function
        Fnd = InStr(Path, "\")
    Loop
    BuildPath = True
End Function

