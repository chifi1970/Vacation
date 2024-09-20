Attribute VB_Name = "MdlSuperMenu"
'Option Explicit

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Const WM_DESTROY As Long = &H2
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum ShowMenuType
    Show_Grid = 0
    Show_Color = 1
    Show_FontList = 2
    Show_Size = 3
    show_Smyles = 4
End Enum

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000

Public sMenuType As ShowMenuType

Public TablaX As Long
Public TablaY As Long
Dim Cancelar As Boolean

Dim m_hwndMenu As Long
Dim PrevProc As Long
Dim ControlPrevProc As Long
Dim m_CotrolHwnd As Long
Dim m_hwndOwner As Long

Public Function ShowDialogColor(ByVal hwndOwner As Long) As Long

    Dim CustomColors() As Byte
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    cc.lStructSize = Len(cc)
    cc.hwndOwner = hwndOwner
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 2
    
    If CHOOSECOLOR(cc) <> 0 Then
        ShowDialogColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowDialogColor = -1
    End If
End Function

Public Function ShowMenuGrid(ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    
    If IsWindow(m_hwndMenu) Then Exit Function
    
    TablaX = 0: TablaY = 0
    
    FrmSuperMenu.ShowPopMenu Show_Grid, hwndOwner, X, Y
    
    Call LoopToEvent(hwndOwner)

    If FrmSuperMenu.TablaX Or FrmSuperMenu.TablaY Then
        ShowMenuGrid = True
        TablaX = FrmSuperMenu.TablaX: TablaY = FrmSuperMenu.TablaY
    End If
    
    Unload FrmSuperMenu
    m_hwndMenu = 0
End Function

Public Function ShowMenuFontSize(ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long) As Long

    If IsWindow(m_hwndMenu) Then Exit Function
    
    FrmSuperMenu.ShowPopMenu Show_Size, hwndOwner, X, Y
    
    Call LoopToEvent(hwndOwner)
    
    ShowMenuFontSize = FrmSuperMenu.ResultSize
    Unload FrmSuperMenu
    m_hwndMenu = 0
End Function

Public Function ShowMenuSmyles(ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long, Imagelist As Object) As Long

    If IsWindow(m_hwndMenu) Then Exit Function
    
    FrmSuperMenu.ShowPopMenu show_Smyles, hwndOwner, X, Y, Imagelist
    
    Call LoopToEvent(hwndOwner)
    
    ShowMenuSmyles = FrmSuperMenu.ResultSmyle
    Unload FrmSuperMenu
    m_hwndMenu = 0
    
End Function

Public Function ShowMenuFontList(ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long) As String
Cancelar = False
    If IsWindow(m_hwndMenu) Then Exit Function
    
    FrmSuperMenu.ShowPopMenu Show_FontList, hwndOwner, X, Y
    
    Call LoopToEvent(hwndOwner)

If Cancelar Then
    Unload FrmSuperMenu
    Exit Function
End If
    ShowMenuFontList = FrmSuperMenu.ResultFontName
    Unload FrmSuperMenu
    m_hwndMenu = 0
End Function

Public Function ShowMenuPaleteColor(ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim ResColor As Long
    
    If IsWindow(m_hwndMenu) Then ShowMenuPaleteColor = -1: Exit Function
    
    FrmSuperMenu.ShowPopMenu Show_Color, hwndOwner, X, Y
    
    Call LoopToEvent(hwndOwner)
    
    If FrmSuperMenu.MoreColor Then

        ResColor = ShowDialogColor(hwndOwner)
        
        If ResColor <> -1 Then
            ShowMenuPaleteColor = ResColor
        Else
            ShowMenuPaleteColor = -1
        End If
    Else
        ShowMenuPaleteColor = FrmSuperMenu.ResultColor
    End If

        Unload FrmSuperMenu
        m_hwndMenu = 0
End Function

Public Sub Unhook()
    SetWindowLong m_hwndOwner, GWL_WNDPROC, PrevProc
    PrevProc = 0
    SetWindowLong m_CotrolHwnd, GWL_WNDPROC, ControlPrevProc
    ControlPrevProc = 0
    'Unload FrmSuperMenu
End Sub


Private Sub LoopToEvent(ByVal hwndOwner As Long)
    'UnHook

    If PrevProc Then Call SetWindowLong(m_hwndOwner, GWL_WNDPROC, PrevProc)
    If ControlPrevProc Then Call SetWindowLong(m_CotrolHwnd, GWL_WNDPROC, ControlPrevProc)
    
    m_hwndMenu = FrmSuperMenu.hwnd
    m_hwndOwner = hwndOwner
    m_CotrolHwnd = GetFocus
   
    SetWindowLong FrmSuperMenu.hwnd, GWL_STYLE, GetWindowLong(FrmSuperMenu.hwnd, GWL_STYLE) Or WS_CHILD

    PrevProc = SetWindowLong(hwndOwner, GWL_WNDPROC, AddressOf WindowProc)
    ControlPrevProc = SetWindowLong(m_CotrolHwnd, GWL_WNDPROC, AddressOf ControlProc)
    
    Do While FrmSuperMenu.Visible = True
        DoEvents
    Loop
    
    
    
    
    SetWindowLong hwndOwner, GWL_WNDPROC, PrevProc
    PrevProc = 0
    SetWindowLong m_CotrolHwnd, GWL_WNDPROC, ControlPrevProc
    ControlPrevProc = 0
    
End Sub


Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 'Debug.Print uMsg, wParam, lParam

If uMsg = WM_SYSCOMMAND And wParam = SC_CLOSE Then
    FrmSuperMenu.Visible = False
    Exit Function
End If

If uMsg = 307 Or uMsg = 273 Then
'Debug.Print uMsg
Else
WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
If uMsg <> 123 And _
    uMsg <> 124 And _
    uMsg <> 125 And _
    uMsg <> 60 And _
    uMsg <> 174 And _
    uMsg <> 132 And _
    uMsg <> 512 And _
    uMsg <> 127 And _
    uMsg <> 70 And _
    uMsg <> 32 And _
    uMsg <> 160 And _
    uMsg <> 674 And _
    uMsg <> 134 And _
    uMsg <> 514 And _
    uMsg <> 533 And _
    uMsg <> 517 And _
    uMsg <> 13 And _
    uMsg <> 14 Then
'Debug.Print uMsg
  FrmSuperMenu.Visible = False

End If
End If


End Function

Public Function ControlProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Select Case uMsg
Case 8
    FrmSuperMenu.Visible = False
Case 256
    
    If wParam = 40 Then FrmSuperMenu.SelectIndex 1
    If wParam = 38 Then FrmSuperMenu.SelectIndex -1
    If wParam = 13 Then FrmSuperMenu.Acept
    If wParam = 27 Then FrmSuperMenu.Visible = False
    'If wParam = 8 Then FrmSuperMenu.Visible = False
Case Else
ControlProc = CallWindowProc(ControlPrevProc, hwnd, uMsg, wParam, lParam)
End Select

End Function


