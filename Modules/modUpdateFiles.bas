Attribute VB_Name = "modUpdateFiles"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Mike Davis
'    ----------------------------------------------------------
'     Module Name: modUpdateFiles
'     Module File: Modules\modUpdateFiles.bas
'     Module Type: Module
'    ----------------------------------------------------------
'     Copyright © 2003 R.W.A.C. Software
'    ----------------------------------------------------------
'   ============================================================

Public Sub UpdateForms(tFile2 As String, tIndex As Integer)
Dim Found2 As Boolean, Found As Boolean
frmMain.Caption = "VBProject Pro v2.5 - Status:  Updating Form " & tIndex & " of " & frmMain.lstForms.ListItems.Count
Close #1
frmMain.txtMove.Text = ""
Open tFile2 For Input As #1

Do While Not EOF(1)
    If InStr(1, LCase(tLine), "vb.form") <> 0 And Found = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Begin VB.Form " & frmMain.lstForms.ListItems(tIndex).Text & vbCrLf
        Line Input #1, tLine
        Found = True
    ElseIf InStr(1, LCase(tLine), "vb.mdiform") <> 0 And Found = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Begin VB.MDIForm " & frmMain.lstForms.ListItems(tIndex).Text & vbCrLf
        Line Input #1, tLine
        Found = True
    ElseIf InStr(1, LCase(tLine), "attribute vb_name") <> 0 And Found2 = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Attribute VB_Name = " & """" & frmMain.lstForms.ListItems(tIndex).Text & """" & vbCrLf
        Line Input #1, tLine
        Found2 = True
    Else
        frmMain.txtMove.Text = frmMain.txtMove.Text & tLine & vbCrLf
        Line Input #1, tLine
    End If
Loop
frmMain.txtMove.Text = frmMain.txtMove.Text & tLine
Close #1

Kill tFile2

Open tFile2 For Output As #1
Print #1, frmMain.txtMove.Text
Close #1

End Sub

Public Sub UpdateModules(tFile2 As String, tIndex As Integer)
Dim Found As Boolean
frmMain.Caption = "VBProject Pro v2.5 - Status:  Updating Module " & tIndex & " of " & frmMain.lstModules.ListItems.Count
Close #1
frmMain.txtMove.Text = ""
Open tFile2 For Input As #1

Do While Not EOF(1)
    If InStr(1, LCase(tLine), "attribute vb_name") <> 0 And Found = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Attribute VB_Name = " & """" & frmMain.lstModules.ListItems(tIndex).Text & """" & vbCrLf
        Line Input #1, tLine
        Found = True
    Else
        frmMain.txtMove.Text = frmMain.txtMove.Text & tLine & vbCrLf
        Line Input #1, tLine
    End If
Loop
frmMain.txtMove.Text = frmMain.txtMove.Text & tLine
Close #1

Kill tFile2

Open tFile2 For Output As #1
Print #1, frmMain.txtMove.Text
Close #1

End Sub

Public Sub UpdateClasses(tFile2 As String, tIndex As Integer)
Dim Found As Boolean
frmMain.Caption = "VBProject Pro v2.5 - Status:  Updating Class " & tIndex & " of " & frmMain.lstClasses.ListItems.Count
Close #1
frmMain.txtMove.Text = ""
Open tFile2 For Input As #1

Do While Not EOF(1)
    If InStr(1, LCase(tLine), "attribute vb_name") <> 0 And Found = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Attribute VB_Name = " & """" & frmMain.lstClasses.ListItems(tIndex).Text & """" & vbCrLf
        Line Input #1, tLine
        Found = True
    Else
        frmMain.txtMove.Text = frmMain.txtMove.Text & tLine & vbCrLf
        Line Input #1, tLine
    End If
Loop
frmMain.txtMove.Text = frmMain.txtMove.Text & tLine
Close #1

Kill tFile2

Open tFile2 For Output As #1
Print #1, frmMain.txtMove.Text
Close #1

End Sub

Public Sub UpdateControls(tFile2 As String, tIndex As Integer)
Dim Found2 As Boolean, Found As Boolean
frmMain.Caption = "VBProject Pro v2.5 - Status:  Updating User Control " & tIndex & " of " & frmMain.lstControls.ListItems.Count
Close #1
frmMain.txtMove.Text = ""
Open tFile2 For Input As #1

Do While Not EOF(1)
    If InStr(1, LCase(tLine), "vb.usercontrol") <> 0 And Found = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Begin VB.UserControl " & frmMain.lstControls.ListItems(tIndex).Text & vbCrLf
        Line Input #1, tLine
        Found = True
    ElseIf InStr(1, LCase(tLine), "attribute vb_name") <> 0 And Found2 = False Then
        frmMain.txtMove.Text = frmMain.txtMove.Text & "Attribute VB_Name = " & """" & frmMain.lstControls.ListItems(tIndex).Text & """" & vbCrLf
        Line Input #1, tLine
        Found2 = True
    Else
        frmMain.txtMove.Text = frmMain.txtMove.Text & tLine & vbCrLf
        Line Input #1, tLine
    End If
Loop
frmMain.txtMove.Text = frmMain.txtMove.Text & tLine
Close #1

Kill tFile2

Open tFile2 For Output As #1
Print #1, frmMain.txtMove.Text
Close #1

End Sub

