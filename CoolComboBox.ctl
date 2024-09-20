VERSION 5.00
Begin VB.UserControl CoolComboBox 
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   ScaleHeight     =   315
   ScaleWidth      =   2955
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "CoolComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Const DT_CENTER = &H1
Const DT_CALCRECT = &H400

Public Enum ComboState
    CB_Normal = 0
    CB_Hot = 1
    CB_Presed = 2
End Enum

Public Enum ComboStyle
    Word2000 = 0
    VBNet = 1
End Enum


Event DropDown()

Dim Rec(8) As RECT
Dim iFont(8) As String
Dim AlphaResalteColor As Long

Dim m_Style As ComboStyle
Dim m_CbState As ComboState
Dim m_Text As String


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


Public Property Get State() As ComboState
    State = m_CbState
End Property

Public Property Let State(ByVal New_CbState As ComboState)
    m_CbState = New_CbState
    Draw New_CbState
End Property

Public Property Get Style() As ComboStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ComboStyle)
    m_Style = New_Style
    Draw m_CbState
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    Draw m_CbState
End Property


Private Sub Draw(ByVal CbState As ComboState)


Dim tRec As RECT
Dim MidH As Long
Dim MidW As Long

Dim BorderColor As Long
Dim BackColor As Long
Dim SpinBorderColor As Long
Dim SpinBackColor As Long
Dim ResalteColor As Long

'Debug.Print CbState

DrawText UserControl.hDC, m_Text, Len(m_Text), tRec, DT_CALCRECT

ResalteColor = IIf(m_Style = 0, vbHighlight, &H99FF&)

If m_CbState = CB_Presed Then CbState = CB_Presed

Select Case CbState
    Case CB_Normal
        BorderColor = UserControl.BackColor
        SpinBackColor = IIf(m_Style = 0, &H8000000F, &HF8E4D8)
        SpinBorderColor = UserControl.BackColor
    Case CB_Hot
        BorderColor = IIf(m_Style = 0, ResalteColor, vbBlack)
        SpinBackColor = pvAlphaBlend(ResalteColor, UserControl.BackColor, 100)
        SpinBorderColor = IIf(m_Style = 0, ResalteColor, vbBlack)
    Case CB_Presed
        BorderColor = IIf(m_Style = 0, ResalteColor, vbBlack)
        SpinBackColor = pvAlphaBlend(ResalteColor, UserControl.BackColor, 150)
        SpinBorderColor = IIf(m_Style = 0, ResalteColor, vbBlack)
End Select

    UserControl.FillColor = UserControl.BackColor
    UserControl.ForeColor = BorderColor
    Rectangle hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    UserControl.FillColor = SpinBackColor
    UserControl.ForeColor = SpinBorderColor
    Rectangle hDC, UserControl.ScaleWidth - 16, 0, UserControl.ScaleWidth, UserControl.ScaleHeight



MidW = UserControl.ScaleWidth - 8
MidH = (UserControl.ScaleHeight / 2)

UserControl.Line (MidW - 3, MidH - 1)-(MidW + 2, MidH - 1), vbBlack
UserControl.Line (MidW - 2, MidH)-(MidW + 1, MidH), vbBlack
UserControl.Line (MidW - 1, MidH + 1)-(MidW, MidH + 1), vbBlack


SetRect tRec, 2, MidH - tRec.Bottom / 2, UserControl.ScaleWidth - 16, UserControl.ScaleHeight
UserControl.ForeColor = vbBlack
DrawText UserControl.hDC, m_Text, Len(m_Text), tRec, 0

UserControl.Refresh


End Sub


Private Sub Timer1_Timer()
Dim pt As POINTAPI
GetCursorPos pt
If WindowFromPoint(pt.X, pt.Y) <> UserControl.hwnd Then
    Timer1.Interval = 0
    Draw CB_Normal
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Draw CB_Presed
    RaiseEvent DropDown
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 10
If Button = 0 Then Draw CB_Hot
'If Button = 1 Then Draw CB_Presed

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draw CB_Hot
End Sub

Private Sub UserControl_Resize()
Draw CB_Normal
End Sub

Private Sub UserControl_Show()
AlphaResalteColor = pvAlphaBlend(vbHighlight, UserControl.BackColor, 100)
UserControl.FillStyle = 0
UserControl.ScaleMode = 3
UserControl.AutoRedraw = True
Draw CB_Normal
End Sub

