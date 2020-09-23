

Attribute VB_Name = "modRelativePath"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: Get Relative Path For VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Manas Tungare
'    ----------------------------------------------------------
'     Module Name: modRelativePath
'     Module File: Modules\modRelativePath.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     All Credits For This Module Belong To Manas Tungare
'    ----------------------------------------------------------
'   ============================================================

Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
    Private Const MAX_PATH As Long = 260
    Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
    Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Function RelativePath(ByVal parent_path As String, ByVal child_path As String) As String
    Dim out_str As String
    Dim par_str As String
    Dim child_str As String
    out_str = String(MAX_PATH, 0)
    par_str = parent_path + String(100, 0)
    child_str = child_path + String(100, 0)
    PathRelativePathTo out_str, par_str, FILE_ATTRIBUTE_DIRECTORY, child_str, FILE_ATTRIBUTE_NORMAL
    out_str = StripTerminator(out_str)
    RelativePath = out_str
End Function
Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function
