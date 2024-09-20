VERSION 5.00
Begin VB.Form FrmSuperMenu 
   BackColor       =   &H00F9FCFC&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FrmSuperMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    a                   As Byte
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

Const DT_CENTER = &H1
Const DT_CALCRECT = &H400
Private Const CS_DROPSHADOW     As Long = &H20000
Private Const GCL_STYLE         As Long = -26

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


Public TablaX           As Long
Public TablaY           As Long
Public ResultSize       As Long
Public ResultFontName   As String
Public MoreColor        As Boolean
Public ResultColor      As Long
Public ResultSmyle      As Long

Dim Rec()               As RECT
Dim iFont()             As String
Dim iColor()            As Long

Dim LastX               As Long
Dim LastY               As Long

Dim m_sMenuType         As ShowMenuType
Dim AlphaResalteColor   As Long
Dim StyleForm           As Long
Dim m_ObjImageList      As Object
Dim Index As Long


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


Private Sub Form_Deactivate()
    Me.Visible = False
End Sub

Public Sub ShowPopMenu(sMenuType As ShowMenuType, ByVal hwndOwner As Long, ByVal X As Long, ByVal Y As Long, Optional Imagelist As Object)

Dim i As Integer
Dim XX As Long
Dim YY As Long

Dim ClientPT As POINTAPI

AlphaResalteColor = pvAlphaBlend(vbHighlight, Me.BackColor, 120)

Me.FillStyle = 0
Me.ScaleMode = 3
Me.AutoRedraw = True
m_sMenuType = sMenuType

ClientToScreen hwndOwner, ClientPT
Index = -1

Select Case m_sMenuType
    Case Show_Grid
        Me.Move ClientPT.X * Screen.TwipsPerPixelX + X, ClientPT.Y * Screen.TwipsPerPixelY + Y, ((30 * 5) + 4) * Screen.TwipsPerPixelX, ((30 * 3) + 50) * Screen.TwipsPerPixelY
        Me.FontName = "MS Sans Serif"
        DrawGrid 0, 0
    Case Show_Color
        Me.Move ClientPT.X * Screen.TwipsPerPixelX + X, ClientPT.Y * Screen.TwipsPerPixelY + Y, 2520, 2205
        ReDim iColor(42)
        ReDim Rec(42)
        MoreColor = False
        ResultColor = -1
        Me.FontSize = 12
        iColor(0) = &H0
        iColor(1) = &H3399
        iColor(2) = &H3333
        iColor(3) = &H3300
        iColor(4) = &H663300
        iColor(5) = &H800000
        iColor(6) = &H993333
        iColor(7) = &H333333
        iColor(8) = &H80
        iColor(9) = &H66FF
        iColor(10) = &H8080&
        iColor(11) = &H808080
        iColor(12) = &H808000
        iColor(13) = &HFF0000
        iColor(14) = &H996666
        iColor(15) = &H808080
        iColor(16) = &HFF
        iColor(17) = &H99FF&
        iColor(18) = &HCC99&
        iColor(19) = &H669933
        iColor(20) = &HCCCC33
        iColor(21) = &HFF6633
        iColor(22) = &H800080
        iColor(23) = &H999999
        iColor(24) = &HFF00FF
        iColor(25) = &HCCFF&
        iColor(26) = &HFFFF&
        iColor(27) = &HFF00&
        iColor(28) = &HFFFF00
        iColor(29) = &HFFCC00
        iColor(30) = &H663399
        iColor(31) = &HC0C0C0
        iColor(32) = &HCC99FF
        iColor(33) = &H99CCFF
        iColor(34) = &H99FFFF
        iColor(35) = &HCCFFCC
        iColor(36) = &HFFFFCC
        iColor(37) = &HFFCC99
        iColor(38) = &HFF99CC
        iColor(39) = &HFFFFFF
        iColor(40) = &H0
        
        SetRect Rec(40), 4, 4, 165, 24
        SetRect Rec(41), 4, 124, 165, 144
        SetRect Rec(42), 7, 7, 21, 21
        
        For i = 0 To 39
            If XX = 8 Then XX = 0: YY = YY + 1
            SetRect Rec(i), (20 * XX) + 4, (20 * YY) + 24, (20 * XX) + 24, (20 * YY) + 20 + 24
            XX = XX + 1
        Next
        
        DrawPaleteColors 0, 0
        
    Case Show_FontList
    
        ReDim Rec(8)
        ReDim iFont(8)
        ResultFontName = ""
        Me.Move ClientPT.X * Screen.TwipsPerPixelX + X, ClientPT.Y * Screen.TwipsPerPixelY + Y, 3525, 4080
        For i = 0 To 8
            SetRect Rec(i), 2, (30 * i) + 2, Me.ScaleWidth - 2, (30 * i) + 30
        Next
        
        iFont(0) = "Arial"
        iFont(1) = "Bookman Old Style"
        iFont(2) = "Courier"
        iFont(3) = "Garamond"
        iFont(4) = "Lucida Console"
        iFont(5) = "Symbol"
        iFont(6) = "Tahoma"
        iFont(7) = "Times New Roman"
        iFont(8) = "Verdana"
        DrawFontList 0, 0

    Case Show_Size
    
        ReDim Rec(6)
        ReDim iFont(6)
        ResultSize = 0
        Me.Move ClientPT.X * Screen.TwipsPerPixelX + X, ClientPT.Y * Screen.TwipsPerPixelY + Y, 3255, 2970
        Me.FontName = "Times New Roman"
        SetRect Rec(0), 2, 0 + 2, Me.ScaleWidth - 2, 14
        SetRect Rec(1), 2, 14 + 2, Me.ScaleWidth - 2, 30
        SetRect Rec(2), 2, 30 + 2, Me.ScaleWidth - 2, 48
        SetRect Rec(3), 2, 48 + 2, Me.ScaleWidth - 2, 72
        SetRect Rec(4), 2, 72 + 2, Me.ScaleWidth - 2, 100
        SetRect Rec(5), 2, 100 + 2, Me.ScaleWidth - 2, 138
        SetRect Rec(6), 2, 138 + 2, Me.ScaleWidth - 2, 196
        iFont(0) = 7
        iFont(1) = 8
        iFont(2) = 10
        iFont(3) = 14
        iFont(4) = 18
        iFont(5) = 24
        iFont(6) = 38
        
        DrawFontSize 0, 0
        
    Case show_Smyles
        Set m_ObjImageList = Imagelist
        ReDim Rec(m_ObjImageList.ListImages.Count)
        ResultSmyle = 0
        
        For i = 1 To m_ObjImageList.ListImages.Count

            SetRect Rec(i - 1), ((m_ObjImageList.ImageWidth + 6) * XX) + 10, ((m_ObjImageList.ImageHeight + 6) * YY) + 10, _
            ((m_ObjImageList.ImageWidth + 6) * XX) + (19 + 6) + 10, ((m_ObjImageList.ImageHeight + 6) * YY) + (19 + 6) + 10

            XX = XX + 1
            If XX = 5 Then
                XX = 0
        
                If i <> m_ObjImageList.ListImages.Count Then YY = YY + 1
            End If
        Next
        
        Me.Move ClientPT.X * Screen.TwipsPerPixelX + X, ClientPT.Y * Screen.TwipsPerPixelY + Y, _
            (((m_ObjImageList.ImageWidth + 6) * 5) + 20) * Screen.TwipsPerPixelX, _
            (((m_ObjImageList.ImageHeight + 6) * (YY + 1)) + 20) * Screen.TwipsPerPixelY
            
        
        DrawSmyles 0, 0
End Select

SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'Me.Show

End Sub

Private Sub DrawSmyles(ByVal X As Long, ByVal Y As Long)
Dim i As Integer
Dim XX As Long
Dim YY As Long


    Me.FillColor = Me.BackColor
    Me.ForeColor = &H738F8F
    Rectangle hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Me.FillColor = AlphaResalteColor
    Me.ForeColor = vbHighlight

    For i = 1 To m_ObjImageList.ListImages.Count
    
        SetRect Rec(i - 1), ((m_ObjImageList.ImageWidth + 6) * XX) + 10, ((m_ObjImageList.ImageHeight + 6) * YY) + 10, _
        ((m_ObjImageList.ImageWidth + 6) * XX) + (m_ObjImageList.ImageWidth + 6) + 10, ((m_ObjImageList.ImageHeight + 6) * YY) + (m_ObjImageList.ImageHeight + 6) + 10
    
        If PtInRect(Rec(i - 1), X, Y) Then
            Rectangle Me.hDC, Rec(i - 1).Left, Rec(i - 1).top, Rec(i - 1).Right, Rec(i - 1).Bottom
        End If
    
        'm_ObjImageList.ListImages(i).Draw Me.hdc, Rec(i - 1).Left + 3, Rec(i - 1).Top + 3, 1
        ImageList_Draw m_ObjImageList.hImageList, i - 1, Me.hDC, Rec(i - 1).Left + 3, Rec(i - 1).top + 3, 1
    
        XX = XX + 1
        If XX = 5 Then
            XX = 0
            YY = YY + 1
        End If
    Next
    
    Me.Refresh


End Sub




Private Sub DrawPaleteColors(ByVal X As Long, ByVal Y As Long)
    Dim i As Integer

    Me.ForeColor = &H738F8F
    Me.FillColor = Me.BackColor
    
    Rectangle hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    
    For i = 0 To 41
    
        If PtInRect(Rec(i), X, Y) Or Index = i Then
            Me.ForeColor = vbHighlight
            Me.FillColor = AlphaResalteColor
            Rectangle hDC, Rec(i).Left, Rec(i).top, Rec(i).Right, Rec(i).Bottom
            Me.ForeColor = &H738F8F
            Index = i
        End If
        
        
        If i < 40 Then
        Me.FillColor = iColor(i)
        Rectangle hDC, Rec(i).Left + 3, Rec(i).top + 3, Rec(i).Right - 3, Rec(i).Bottom - 3
        End If
    
    Next

    If PtInRect(Rec(40), X, Y) Or Index = 40 Then
        Me.ForeColor = vbHighlightText
    Else
        Me.ForeColor = vbWindowText
    End If
    
    DrawText hDC, "Automático", 10, Rec(40), DT_CENTER
    
    If PtInRect(Rec(41), X, Y) Or Index = 41 Then
        Me.ForeColor = vbHighlightText
    Else
        Me.ForeColor = vbWindowText
    End If
    
    DrawText hDC, "Más Colores...", 14, Rec(41), DT_CENTER
    
    
    Me.FillColor = iColor(0)
    Me.ForeColor = &H738F8F
    
    Rectangle hDC, Rec(42).Left, Rec(42).top, Rec(42).Right, Rec(42).Bottom
    
    Me.Refresh


End Sub


Private Sub DrawFontList(ByVal X As Long, ByVal Y As Long)
    Dim i As Integer
    Dim j As Integer
    Dim Label As String
    Dim tRec As RECT
    
    
        Me.FillColor = Me.BackColor
        Me.ForeColor = &H738F8F
    Rectangle hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    For i = 0 To 8
        If PtInRect(Rec(i), X, Y) Or Index = i Then
            Me.FillColor = AlphaResalteColor
            Me.ForeColor = vbHighlight
            Index = i
        Else
            Me.FillColor = Me.BackColor
            Me.ForeColor = &H738F8F
        End If
        
        On Error Resume Next
        Me.FontName = iFont(i)
        On Error GoTo 0
        
        Me.FontSize = 16
        
        Label = iFont(i)
        
        SetRect tRec, 0, 0, 0, 0
        DrawText Me.hDC, Label, Len(Label), tRec, DT_CALCRECT
        Rectangle hDC, Rec(i).Left, Rec(i).top, Rec(i).Right, Rec(i).Bottom
        Dim medio As Single
        
        medio = Rec(i).top + ((Rec(i).Bottom - Rec(i).top) / 2) - (tRec.Bottom / 2)
        
        SetRect tRec, Rec(i).Left, medio, Rec(i).Right, medio + tRec.Bottom
        
        If PtInRect(Rec(i), X, Y) Or Index = i Then
            Me.ForeColor = vbHighlightText
        Else
            Me.ForeColor = vbWindowText
        End If
        
        
        DrawText Me.hDC, Label, Len(Label), tRec, DT_CENTER
    Next
    
    
    
    Me.Refresh


End Sub

Private Sub DrawFontSize(ByVal X As Long, ByVal Y As Long)
    Dim i As Integer
    Dim j As Integer
    Dim Label As String
    
    
    
        Me.FillColor = Me.BackColor
        Me.ForeColor = &H738F8F
        
    Rectangle hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    For i = 0 To 6
        If PtInRect(Rec(i), X, Y) Or Index = i Then
            Me.FillColor = AlphaResalteColor
            Me.ForeColor = vbHighlight
            Index = i
        Else
            Me.FillColor = Me.BackColor
            Me.ForeColor = &H738F8F
        End If
        Rectangle hDC, Rec(i).Left, Rec(i).top, Rec(i).Right, Rec(i).Bottom
        Label = "Tamaño " & i + 1
        Me.FontSize = iFont(i)
        
        If PtInRect(Rec(i), X, Y) Or Index = i Then
            Me.ForeColor = vbHighlightText
        Else
            Me.ForeColor = vbWindowText
        End If
        
        
        DrawText Me.hDC, Label, Len(Label), Rec(i), DT_CENTER
    Next
    
    
    
    Me.Refresh


End Sub



Private Sub Form_Load()
    StyleForm = GetClassLong(Me.hwnd, GCL_STYLE)
    SetClassLong Me.hwnd, GCL_STYLE, StyleForm Or CS_DROPSHADOW
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetClassLong Me.hwnd, GCL_STYLE, StyleForm
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Select Case m_sMenuType
    
        Case Show_Grid
            If Y > Me.ScaleHeight - 20 Then Me.Visible = False
        Case Show_Color
        
        Case Show_FontList
            
        Case Show_Size
    
    
    End Select
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XWidth As Long
Dim YHeight As Long
Dim CurrX As Long
Dim CurrY As Long



Select Case m_sMenuType

    Case Show_Grid
        If Button = 1 Then
            Timer1.Interval = 0
            If X > Me.ScaleWidth Then
                XWidth = X / 30
                Me.Width = ((XWidth * 30) + 4) * Screen.TwipsPerPixelX
            End If
            If Y > Me.ScaleHeight - 15 Then
                YHeight = Y / 30
                Me.Height = ((YHeight * 30) + 50) * Screen.TwipsPerPixelY
            End If
        Else
        Timer1.Interval = 100
        End If
    
        If Y < Me.ScaleHeight - 20 Then DrawGrid X, Y
        'DoEvents
        
    Case Show_Color
        Timer1.Interval = 100
        DrawPaleteColors X, Y
    Case Show_FontList
        DrawFontList X, Y
        Timer1.Interval = 100
    Case Show_Size
        DrawFontSize X, Y
        Timer1.Interval = 100
    Case show_Smyles
        DrawSmyles X, Y
        Timer1.Interval = 100
End Select



End Sub


Private Sub DrawGrid(ByVal X As Long, ByVal Y As Long)
    Dim i As Integer
    Dim j As Integer
    Dim Rec As RECT
    Dim Label As String
    
    TablaX = 0
    TablaY = 0
    
    Me.FillColor = Me.BackColor
    Me.ForeColor = &H738F8F
    
    Rectangle hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    For i = 0 To Me.ScaleWidth / 30
        For j = 0 To (Me.ScaleHeight - 50) / 30
        
            
            If (X > (i * 30) + 2) And (Y > (j * 30) + 2) Then
                Me.ForeColor = vbHighlight
                Me.FillColor = AlphaResalteColor
                TablaX = i + 1
                TablaY = j + 1
            Else
                Me.ForeColor = &H738F8F
                Me.FillColor = Me.BackColor
            End If
            
            Rectangle hDC, (i * 30) + 4, (j * 30) + 4, (i * 30) + 30, (j * 30) + 30
        Next
    Next
    
    If TablaX = 0 And TablaY = 0 Then
        Label = "Cancelar"
        
    Else
        Label = TablaX & " por " & TablaY & " Tabla"
    End If
    
    SetRect Rec, 0, Me.ScaleHeight - 18, Me.ScaleWidth, Me.ScaleHeight
    Me.ForeColor = vbWindowText
    DrawText Me.hDC, Label, Len(Label), Rec, DT_CENTER
    
    Me.Refresh


End Sub

Private Sub Form_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Button = 1 Then
    Select Case m_sMenuType
    
        Case Show_Grid
            Me.Visible = False
        Case Show_Color
            For i = 0 To 41
                If PtInRect(Rec(i), X, Y) Then
                    If i = 41 Then
                        MoreColor = True
                    Else
                        ResultColor = iColor(i)
                    End If
                    Me.Visible = False
                    Exit For
                End If
            Next
    
        Case Show_FontList
            For i = 0 To 8
                If PtInRect(Rec(i), X, Y) Then
                    ResultFontName = iFont(i)
                    Me.Visible = False
                    Exit For
                End If
            Next
    
        Case Show_Size
            For i = 0 To 6
                If PtInRect(Rec(i), X, Y) Then
                    ResultSize = i + 1
                    Me.Visible = False
                    Exit For
                End If
            Next
        Case show_Smyles
            For i = 0 To UBound(Rec)
                If PtInRect(Rec(i), X, Y) Then
                    ResultSmyle = i + 1
                    Me.Visible = False
                    Exit For
                End If
            Next
        
        
        
    End Select
End If
End Sub




Private Sub Timer1_Timer()
    Dim pt As POINTAPI
    Dim ClientPT As POINTAPI
    Dim InForm As Boolean
    'Index = -1
    GetCursorPos pt
    InForm = WindowFromPoint(pt.X, pt.Y) = Me.hwnd
    
    Select Case m_sMenuType
    
        Case Show_Grid
            ClientToScreen Me.hwnd, ClientPT
            
            pt.X = pt.X - ClientPT.X
            pt.Y = pt.Y - ClientPT.Y
            
            If pt.X < 0 Or pt.Y < 0 Or pt.X > Me.ScaleWidth Or pt.Y > Me.ScaleHeight - 20 Then
                DrawGrid 0, 0
            End If
        Case Show_Color
            If Not InForm Then Index = -1: DrawPaleteColors 0, 0
        Case Show_FontList
            If Not InForm Then Index = -1: DrawFontList 0, 0
        Case Show_Size
            If Not InForm Then Index = -1: DrawFontSize 0, 0
        Case show_Smyles
            If Not InForm Then Index = -1: DrawSmyles 0, 0
    End Select
Timer1.Interval = 0
End Sub

Public Sub SelectIndex(ByVal New_Index As Long)


Index = Index + New_Index

If Index > UBound(Rec) Then Index = 0
If Index < 0 Then Index = UBound(Rec)

Select Case m_sMenuType
Case Show_Grid
Case Show_Color
DrawPaleteColors 0, 0
Case Show_FontList
DrawFontList 0, 0
Case Show_Size
DrawFontSize 0, 0

End Select
End Sub



Public Sub Acept()
Dim i As Integer

Select Case m_sMenuType

    'Case Show_Grid
    '    Me.Visible = False
    Case Show_Color
        If Index > -1 Then
            If Index = 41 Then
                MoreColor = True
            Else
                ResultColor = iColor(Index)
            End If
            Me.Visible = False
        End If
    Case Show_FontList
        If Index > -1 Then
                ResultFontName = iFont(Index)
        End If
        
        Me.Visible = False
    Case Show_Size
        If Index > -1 Then
                ResultSize = Index + 1
        End If
        
        Me.Visible = False
End Select
End Sub

