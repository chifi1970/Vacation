VERSION 5.00
Begin VB.UserControl ButtonOffice 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   109
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ButtonOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal lngHandler As Long, ByVal lngIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal lngHandler As Long, ByVal lngIndex As Long, ByVal lngNewClassLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Const DSS_MONO As Long = &H80
Private Const DSS_NORMAL As Long = &H0
Private Const DSS_DISABLED As Long = &H20
Private Const DST_BITMAP As Long = &H4
Private Const DST_ICON As Long = &H3
Private Const DST_COMPLEX As Long = &H0
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long


Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Type tRGB
    R                   As Long
    G                   As Long
    B                   As Long
End Type

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Enum ButonState
    BT_Normal = 0
    BT_Hot = 1
    BT_Presed = 2
End Enum

Public Enum ButonStyle
    BT_2000 = 0
    BT_2003 = 1
End Enum

Private Const DT_LEFT As Long = &H0

Const DT_CENTER = &H1
Const DT_CALCRECT = &H400
Private Const CS_DROPSHADOW     As Long = &H20000
Private Const GCL_STYLE         As Long = -26
Private m_Caption As String
Private m_Picture As StdPicture
Private m_Style As ButonStyle
Private m_Enabled As Boolean
Private m_BackColor As OLE_COLOR
Private BTState As ButonState


Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    m_Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
    
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawButton
End Property


Private Function GetRGB(ByVal iColor As Long) As tRGB
    GetRGB.B = ((iColor And &HFF0000) / 65536)
    GetRGB.G = ((iColor And &HFF00FF00) / 256&)
    GetRGB.R = iColor Mod 256
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Style = .ReadProperty("Style", 0)
        m_Enabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", Empty)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        UserControl.BackColor = .ReadProperty("BackColor", &HD8E9EC)   'vbButtonFace
    End With
    
End Sub

Private Sub UserControl_Resize()
Call DrawButton
End Sub

Private Sub UserControl_Show()
Call DrawButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Style", m_Style, 0)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("Caption", m_Caption, Empty)
        Call .WriteProperty("Picture", m_Picture, Nothing)
        Call .WriteProperty("BackColor", UserControl.BackColor, &HD8E9EC)   'vbButtonFace
    End With
End Sub

Private Function DrawGradient(lLightColor As OLE_COLOR, lDarkColor As OLE_COLOR)

'On Error GoTo ErrHandler
    
    Dim i As Long
    Dim RGB1 As tRGB
    Dim RGB2 As tRGB
    Dim Rx As Long, Gx As Long, Bx As Long
    Dim Rs As Long, Gs As Long, Bs As Long
    
    OleTranslateColor lLightColor, 0, VarPtr(lLightColor)
    OleTranslateColor lDarkColor, 0, VarPtr(lDarkColor)

    RGB1 = GetRGB(lLightColor)
    RGB2 = GetRGB(lDarkColor)

    Rx = RGB1.R: Gx = RGB1.G: Bx = RGB1.B
    Rs = (RGB1.R - RGB2.R) / (UserControl.ScaleHeight - 1)
    Gs = (RGB1.G - RGB2.G) / (UserControl.ScaleHeight - 1)
    Bs = (RGB1.B - RGB2.B) / (UserControl.ScaleHeight - 1)

    For i = 0 To UserControl.ScaleHeight - 1
        UserControl.Line (0, i)-(UserControl.ScaleWidth, i), RGB(Rx, Gx, Bx)
        Rx = Rx - Rs
        Gx = Gx - Gs
        Bx = Bx - Bs
    Next i
        
'ErrHandler:
    Exit Function
End Function


Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    DrawButton
End Property

Public Property Get Style() As ButonStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ButonStyle)
    m_Style = New_Style
    PropertyChanged "Style"
    DrawButton
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawButton
End Property


Private Sub Timer1_Timer()

If Not IsMouseInButton And BTState <> BT_Presed Then
    Timer1.Interval = 0
    BTState = BT_Normal
    DrawButton
End If
End Sub

Private Sub UserControl_Initialize()



m_Caption = "Boton1"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
If Button = 1 Then
    BTState = BT_Presed
    DrawButton
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)

If Button = 0 And BTState <> BT_Hot Then
    Timer1.Interval = 10
    BTState = BT_Hot
    DrawButton
End If
If Button = 1 Then
    Timer1.Interval = 0
    If IsMouseInButton Then
         BTState = BT_Presed
         DrawButton
     Else
         BTState = BT_Hot
         DrawButton
     End If
End If
End Sub

Private Function IsMouseInButton() As Boolean
Dim pt As POINTAPI
GetCursorPos pt
IsMouseInButton = WindowFromPoint(pt.X, pt.Y) = UserControl.hwnd
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)

If Button = 1 Then
   If IsMouseInButton Then
        BTState = BT_Hot
        DrawButton
        RaiseEvent Click
    Else
        BTState = BT_Normal
        DrawButton
        
    End If
End If
End Sub

Private Sub DrawButton()

Dim tRec As RECT
Dim MidH As Long
Dim MidW As Long
Dim PicW As Long
Dim PicH As Long
Dim BorderColor As Long
Dim BackColor As Long


DrawText UserControl.hDC, m_Caption, Len(m_Caption), tRec, DT_CALCRECT

    
    If m_Style = BT_2003 Then
        UserControl.FillStyle = 1
        Select Case BTState
            Case BT_Normal
                Cls
            Case BT_Hot
                DrawGradient &HD6FBFF, &H94D3FF
                UserControl.ForeColor = &H800000
                Rectangle hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
            Case BT_Presed
                DrawGradient &H4A8CFE, &H8ED2FF
                UserControl.ForeColor = 8388608
                Rectangle hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End Select
    Else
        UserControl.FillStyle = 0
        Select Case BTState
            Case BT_Normal
                Cls
            Case BT_Hot
                UserControl.ForeColor = &HC56A31
                UserControl.FillColor = &HEED2C1
                Rectangle hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        
            Case BT_Presed
                UserControl.ForeColor = &H6F4B4B
                UserControl.FillColor = &HE2B598
                Rectangle hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End Select
    End If
    

    
    MidW = UserControl.ScaleWidth
    MidH = (UserControl.ScaleHeight / 2)

    
    If Not m_Picture Is Nothing Then
        PicW = ScaleX(m_Picture.Width, vbHimetric, vbPixels)
        PicH = ScaleX(m_Picture.Height, vbHimetric, vbPixels)
        If BTState = BT_Hot Then
            DrawButtonImage m_Picture, 6, MidH - (PicH / 2), PicW, PicH, True, True
            DrawButtonImage m_Picture, 4, MidH - (PicH / 2) - 2, PicW, PicH, False, True
        Else
            DrawButtonImage m_Picture, 6, MidH - (PicH / 2), PicW, PicH, False, True
        End If
    End If

    SetRect tRec, PicW + 10, MidH - tRec.Bottom / 2, UserControl.ScaleWidth, UserControl.ScaleHeight
    UserControl.ForeColor = vbBlack
    DrawText UserControl.hDC, m_Caption, Len(m_Caption), tRec, 0

    UserControl.Refresh
End Sub


Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore         As UcsRgbQuad
    Dim clrBack         As UcsRgbQuad
    
    OleTranslateColor clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor clrSecond, 0, VarPtr(clrBack)
    With clrFore
        .R = (.R * lAlpha + clrBack.R * (255 - lAlpha)) / 255
        .G = (.G * lAlpha + clrBack.G * (255 - lAlpha)) / 255
        .B = (.B * lAlpha + clrBack.B * (255 - lAlpha)) / 255
    End With
    CopyMemory VarPtr(pvAlphaBlend), VarPtr(clrFore), 4
End Function

Private Sub DrawButtonImage(ByRef m_Picture As StdPicture, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal Width As Long, _
                            ByVal Height As Long, _
                            ByVal bShadow As Boolean, _
                            ByVal Enabled As Boolean)

  Dim lFlags As Long
  Dim hBrush As Long

    On Local Error Resume Next
    Select Case m_Picture.Type
     Case vbPicTypeBitmap
        lFlags = DST_BITMAP
     Case vbPicTypeIcon
        lFlags = DST_ICON
     Case Else
        lFlags = DST_COMPLEX
    End Select
    If bShadow Then
        If m_Style Then
            hBrush = CreateSolidBrush(pvAlphaBlend(vbHighlight, vbButtonShadow, 10))
         Else 'M_OFFICEXPSTYLE = FALSE/0
            hBrush = CreateSolidBrush(pvAlphaBlend(vbHighlight, vbButtonShadow, 10))
        End If
    End If
    If Enabled Then
        DrawState UserControl.hDC, IIf(bShadow, hBrush, 0), 0, m_Picture.handle, 0, X, Y, Width, Height, lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
     Else 'ENABLED = FALSE/0
        DrawState UserControl.hDC, IIf(bShadow, hBrush, 0), 0, m_Picture.handle, 0, X, Y, Width, Height, lFlags Or DSS_DISABLED
    End If
    If bShadow Then
        DeleteObject hBrush
    End If

End Sub



