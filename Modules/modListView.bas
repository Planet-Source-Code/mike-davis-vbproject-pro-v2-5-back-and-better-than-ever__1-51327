Attribute VB_Name = "modListView"
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: VBProject Pro
'                       Visual Basic Project Manager
'     Developer/Programmer: Unknown
'    ----------------------------------------------------------
'     Module Name: modListView
'     Module File: Modules\modListView.bas
'     Module Type: Module
'    ----------------------------------------------------------
'   ============================================================

Option Explicit

Private Const LVIS_STATEIMAGEMASK As Long = &HF000

Private Type LVITEM
    mask         As Long
    iItem        As Long
    iSubItem     As Long
    state        As Long
    stateMask    As Long
    pszText      As String
    cchTextMax   As Long
    iImage       As Long
    lParam       As Long
    iIndent      As Long
End Type

Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_HEADERDRAGDROP = &H10
Private Const LVS_EX_TRACKSELECT = &H8
Private Const LVS_EX_ONECLICKACTIVATE = &H40
Private Const LVS_EX_TWOCLICKACTIVATE = &H80
Private Const LVS_EX_SUBITEMIMAGES = &H2

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const LVIF_STATE = &H8
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)

Private Const HDS_BUTTONS = &H2
Private Const GWL_STYLE = (-16)

Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Declare Function SendMessageAny _
                        Lib "user32" _
                        Alias "SendMessageA" _
                        (ByVal HWND As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        lParam As Any) _
                        As Long

Private Declare Function SendMessageLong Lib _
                        "user32" Alias _
                        "SendMessageA" _
                        (ByVal HWND As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) _
                        As Long
                        
Private Declare Function GetWindowLong _
                        Lib "user32" _
                        Alias "GetWindowLongA" _
                        (ByVal HWND As Long, _
                        ByVal nIndex As Long) _
                        As Long
                        
Private Declare Function SetWindowLong _
                        Lib "user32" _
                        Alias "SetWindowLongA" _
                        (ByVal HWND As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) _
                        As Long
                        
Private Declare Function SetWindowPos _
                        Lib "user32" _
                        (ByVal HWND As Long, _
                        ByVal hWndInsertAfter As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal cx As Long, _
                        ByVal cy As Long, _
                        ByVal wFlags As Long) _
                        As Long
'=======================================================================

'=======================================================================
Public LengthPerCharacter As Long
'=======================================================================

'=======================================================================
' Description: Enables Full Row Select in a ListView
'=======================================================================
Public Function EnhListView_Add_FullRowSelect( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_FullRowSelect
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_FullRowSelect = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.HWND, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_FULLROWSELECT
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.HWND, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_FullRowSelect:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_FullRowSelect = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_FullRowSelect"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_FullRowSelect" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================

'=======================================================================
' Description: Disables Full Row Select in a ListView
'=======================================================================
'=======================================================================
' Description: Enables GridLines in a ListView
'=======================================================================
Public Function EnhListView_Add_GridLines( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_Add_GridLines

    '________________________________________________________________________
    ' set function return to true
    EnhListView_Add_GridLines = True
    
    '________________________________________________________________________
    ' setup variables
    Dim rStyle  As Long
    Dim r       As Long
    
    '________________________________________________________________________
    ' get the current styles
    rStyle = SendMessageLong(lstListViewName.HWND, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    '________________________________________________________________________
    ' add the selected style to the current styles
    rStyle = rStyle Or LVS_EX_GRIDLINES
    
    '________________________________________________________________________
    ' update the listview styles
    SendMessageLong lstListViewName.HWND, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_Add_GridLines:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_Add_GridLines = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_Add_GridLines"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_Add_GridLines" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function
'=======================================================================
