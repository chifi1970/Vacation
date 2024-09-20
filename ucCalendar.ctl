VERSION 5.00
Begin VB.UserControl ucCalendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   770
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1320
      Top             =   1440
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   6720
      Max             =   48
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ucCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
''-----------------------------------------------
'AddEvents: funcion principal para agregar eventos al calendario, parametros requerido;  Sujeto (titulo descriptivo),  fecha y Hora de inicio Fecha y Hora finalizacion, valor de retorno una key del evento.
'CenterCalenarInNow: mueve el scroll a la hora actual.
'Clear: elimina todo los eventos.
'DateValue: asigna o retorna la fecha actual del calendario.
'DayHaveEvents: consulta si hay eventos en un dia especifico.
'DropDownColor: cuando hay muchos eventos en un dia y en el modo de vista Mes hay mas de los que se pueden mostrar, se muestra una barra desplegable la cual podemos cambiar el color con esta propiedad.
'EventsCount: cantidad de eventos agregados.
'EventsRoundCorner: propiedad booleana para mostrar o no esquinas redondeadas en enventos y botones.
'FirstDayOfWeek: aqui podesmos asignar que dia queremos que se muestre como primer dia de la semana, por defecto usa el del sistema.
'GetAllEvents: obtiene una coleccion de las keys de los eventos agregados.
'GetEventData: obtiene los datos de un evento, su primer parametro es la key del evento la cual podemos obenenrla con GetAllEvents o por algun evento, el resto de los parametros son valores de retorno.
'GetEventsFromDay retorna una coleccion de keys de los eventos de un dia en especifico,
'GetSelectionRangeDate: funcion para obtener el rango de fechas selecionada.
'HeaderColor: color de la cabecera y parte de la tematica.
'HiddeEvent: oculta el evento, util para filtrar.
'LinesColor: color de las lineas.
'Redraw: habilita o desabilita el repintando del calendario, esto sirve para acelerar la carga de eventos.
'RemoveEvent: elimina un evento.
'Refresh: refresca el repintado calendario.
'SelectedEvent: retorna el evento selecionado.
'SelectionColor: Color de la seleción.
'SetStrLanguage: aqui podemos pasar la traducion de las palabras utilizadas.
'ShowToolTipEvents: si se quiere mostrar o no la ventana tooltip, esta se puede remplazar por otra personalizada y con informacion mas detallada. vease los eventos EventMouseEnter y EventMouseLeave
'Update: es mas completo que refresh, este vuelve a reordenar por fecha y alfaveticamente los evento, recalcula la posicion y por ulitmo repinta todo.
'UpdateEventData funcion para actualizar los datos de un evento. deve pasarse la key del evento que queresmos modificar.
'UserCanChangeDate: habilita o desabilita para que el usuario pueda cambiar la pagina actual del calendario.
'UserCanChangeEvents: habilita o desabilita si el usuario puede cambiar los eventos (mediante estiramiento o arrastre)
'UserCanChangeViewMode: oculta todos los botones en la parte superior de la parte derecha. de esta forma el usuario no puede cambiar el modo de vista o bien el progrmador toma el control de que vista quiere mostrar.
'UserCanScrollMonth: en el modo de vista mes, se puede scrollear infinitamente si necesidad de cambiar de pagina,esto solo si se habilita esta opcion.
'ViewMode: cambia por codigo en el modo de vista (Dia, Semama, Mes, Año)
'----------------------------------------------------------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mEndCap As Long) As Long
Private Declare Function GdipDrawClosedCurve Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTF, ByVal mCount As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mHbmReturn As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mStartCap As Long) As Long
Private Declare Function GdipFillEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "GdiPlus.dll" (ByVal mHatchStyle As HatchStyle, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long

Private Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
    
Public Enum HatchStyle
    HatchStyleHorizontal = &H0
    HatchStyleVertical = &H1
    HatchStyleForwardDiagonal = &H2
    HatchStyleBackwardDiagonal = &H3
    HatchStyleCross = &H4
    HatchStyleDiagonalCross = &H5
    HatchStyle05Percent = &H6
    HatchStyle10Percent = &H7
    HatchStyle20Percent = &H8
    HatchStyle25Percent = &H9
    HatchStyle30Percent = &HA
    HatchStyle40Percent = &HB
    HatchStyle50Percent = &HC
    HatchStyle60Percent = &HD
    HatchStyle70Percent = &HE
    HatchStyle75Percent = &HF
    HatchStyle80Percent = &H10
    HatchStyle90Percent = &H11
    HatchStyleLightDownwardDiagonal = &H12
    HatchStyleLightUpwardDiagonal = &H13
    HatchStyleDarkDownwardDiagonal = &H14
    HatchStyleDarkUpwardDiagonal = &H15
    HatchStyleWideDownwardDiagonal = &H16
    HatchStyleWideUpwardDiagonal = &H17
    HatchStyleLightVertical = &H18
    HatchStyleLightHorizontal = &H19
    HatchStyleNarrowVertical = &H1A
    HatchStyleNarrowHorizontal = &H1B
    HatchStyleDarkVertical = &H1C
    HatchStyleDarkHorizontal = &H1D
    HatchStyleDashedDownwardDiagonal = &H1E
    HatchStyleDashedUpwardDiagonal = &H1F
    HatchStyleDashedHorizontal = &H20
    HatchStyleDashedVertical = &H21
    HatchStyleSmallConfetti = &H22
    HatchStyleLargeConfetti = &H23
    HatchStyleZigZag = &H24
    HatchStyleWave = &H25
    HatchStyleDiagonalBrick = &H26
    HatchStyleHorizontalBrick = &H27
    HatchStyleWeave = &H28
    HatchStylePlaid = &H29
    HatchStyleDivot = &H2A
    HatchStyleDottedGrid = &H2B
    HatchStyleDottedDiamond = &H2C
    HatchStyleShingle = &H2D
    HatchStyleTrellis = &H2E
    HatchStyleSphere = &H2F
    HatchStyleSmallGrid = &H30
    HatchStyleSmallCheckerBoard = &H31
    HatchStyleLargeCheckerBoard = &H32
    HatchStyleOutlinedDiamond = &H33
    HatchStyleSolidDiamond = &H34
    HatchStyleTotal = &H35
    HatchStyleLargeGrid = &H4
    HatchStyleMin = &H0
    HatchStyleMax = &H34
End Enum

Private Enum ButonState
    Normal = 0
    Hot = 1
    Pressed = 2
    Selected = 3
    disabled = 4
End Enum

Public Enum eEventShowAs
    [ESA_Busy]
    [ESA_Free]
    [ESA_Out of office]
    [ESA_Tentative]
    [ESA_Working elsewhere]
End Enum

Public Enum EnuViewMode
    vm_Year
    vm_Month
    vm_Week
    vm_Day
End Enum

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type RECTF
    Left As Single
    top As Single
    Width As Single
    Height As Single
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
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

Private Type TOOLINFO
    cbSize              As Long
    uFlags              As Long
    hwnd                As Long
    uId                 As Long
    RECT                As RECT
    hInst               As Long
    lpszText            As String
    lParam              As Long
End Type

Private Type CalEvents
    key As Long
    StartTime As Date
    EndTime As Date
    Subject As String
    office As String
    body As String
    Rects() As RECT
    RectsCount As Long
    ForeColor As Long
    AllDayEvent As Boolean
    More24Hours As Boolean
    Image As Long
    Hidden As Boolean
    Tag As Boolean
    IsSerie As Boolean
    IsPrivate As Boolean
    NotifyIcon As Boolean
    EventShowAs As eEventShowAs
    idvacaciones As Integer
    email1 As String
    email2 As String
End Type

Private Type CalGrid
    EventsCount As Long
    HaveHideEvents As Boolean
End Type

Private Type CalButtons
    Caption As String
    RECT As RECT
    State As ButonState
End Type

Private Type ColGrid
    Rows() As Long
End Type

Private Const IDC_HAND                  As Long = 32649
Private Const GWL_WNDPROC               As Long = -4
Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_USER                   As Long = &H400&
Private Const WM_DESTROY                As Long = &H2
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CombineModeExclude        As Long = &H4
Private Const TME_LEAVE                 As Long = &H2&
Private Const WS_EX_TOPMOST             As Long = &H8&
Private Const TTM_TRACKACTIVATE         As Long = (WM_USER + 17)
Private Const TTM_ADDTOOLW              As Long = WM_USER + 50& ' Add a new tooltip.
Private Const TTM_DELTOOLW              As Long = WM_USER + 51& ' Delete an existing tooltip
Private Const TTM_UPDATETIPTEXTW        As Long = WM_USER + 57& ' Update text in an existing tooltip.
Private Const TTM_SETTITLEW             As Long = WM_USER + 33& ' Set title above the tooltip; 100 chars max
'Private Const TTM_SETDELAYTIME          As Long = WM_USER + 3& ' Sets one of Reshow, Autopoop or Initial times (milliseconds), -1 use system defaults
'Private Const TTDT_RESHOW           As Long = 1& ' Milliseconds for subsequent tooltips to appear as pointer moves from one tool to another, -1 to reset to default
'Private Const TTDT_AUTOPOP          As Long = 2& ' MS for tool to show if pointer is stationary in a tooltip, -1 to reset to default
'Private Const TTDT_INITIAL          As Long = 3& ' MS until a tooltip is displayed after pointer is stationary within a tool's bounding rectangle, -1 to reset to default
Private Const TTS_ALWAYSTIP             As Long = &H1
Private Const TOOLTIPS_CLASS            As String = "tooltips_class32"
Private Const TTF_IDISHWND              As Long = &H1
Private Const LineCapRound              As Long = &H2
Private Const LineCapArrowAnchor        As Long = &H14
Private Const DashStyleDot              As Long = &H2
Private Const StringFormatFlagsNoWrap   As Long = &H1000
Private Const StringTrimmingEllipsisCharacter As Long = &H3


Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
'Public Event PrePaint(hdc As Long, X As Long, Y As Long)
'Public Event PostPaint(ByVal hdc As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event EVENTCLICK(ByVal EventKey As Long, Button As Integer)
Public Event EventMouseEnter(ByVal EventKey As Long)
Public Event EventMouseLeave(ByVal EventKey As Long)
Public Event PreEventChangeDate(ByVal EventKey As Long, ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
Public Event EventChangeDate(ByVal EventKey As Long, ByVal StartDate As Date, ByVal EndDate As Date, ByVal AllDay As Boolean)
Public Event PreDateChange(NewDate As Date, Cancel As Boolean)
Public Event DateChange(NewDate As Date)
Public Event DateBackColor(CellDate As Date, Color As OLE_COLOR, eHatchStyle As HatchStyle)
Public Event DropDownViewMore(CellDate As Date, CancelViewModeDay As Boolean)
Public Event DragNewEvent(ByVal EventKey As Long, ByVal SourceKey As Long)

Dim m_Font As StdFont
Dim m_ForeColor As OLE_COLOR
Dim m_LinesColor As OLE_COLOR
Dim m_HeaderColor As OLE_COLOR
Dim m_SelectionColor As OLE_COLOR
Dim m_DropDownColor As OLE_COLOR
Dim m_ForeColorAlpha As Long
Dim m_MousePointerHands As Long
Dim m_UserCanChangeEvents As Boolean
Dim m_UserCanChangeViewMode As Boolean
Dim m_UserCanChangeDate As Boolean
Dim m_UserCanScrollMonth As Boolean
Dim m_FirstDayOfWeek As VbDayOfWeek
Dim m_ShowToolTipEvents As Boolean
Dim m_Redraw As Boolean
Dim m_TopHeaderHeight As Long
Dim m_ColumnHeaderHeight As Long
Dim m_EventsRoundCorner As Boolean

Dim mTodayHeight As Long

Dim hCur As Long
Dim nScale As Single
Dim xDate As Date
Dim HeaderFont As StdFont
Dim RowHeight As Single, RowWidth As Single
Dim SelStart As Long
Dim SelEnd As Long
Dim StartDay As Long
Dim EndDay As Long
Dim PenWidth  As Single
Dim MarginText As Single
Dim EventHeight As Single
Dim ButtonsSize As Single
Dim FirstDate As Date
Dim LastDate As Date
Dim RowHW As Long
Dim mSelectedEvent As Long
Dim bMouseDownInCal As Boolean
Dim isMouseEnter As Boolean
Dim PrevWndProc As Long
Dim bvASM(43) As Byte
Dim tEvents() As CalEvents
Dim tGrid() As CalGrid
Dim tButtons(6) As CalButtons
Dim eViewMode As EnuViewMode
Dim mStrStarts As String
Dim mStrEnds As String
Dim mStrAllDay As String
Dim DayIndex As Integer
Dim mEventsCount As Long
Dim m_WeekCounts As Long
Dim WeekScroll As Long
Dim mStartDrag As Boolean
Dim mStartSize As Boolean
Dim mSizeDirection As Long
Dim mDragKey As Long
Dim PointMdown As POINTAPI
Dim mEvDragDaysCount As Long
Dim mEvDragFromDayNro As Long
Dim GridWeek() As ColGrid
Dim mHotEvent As Long
Dim YearCalRects() As RECT
Dim YearCalIndex As Long
Dim TI As TOOLINFO
Dim m_hwndTT As Long
Dim GdipToken As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProcA(PrevWndProc, hwnd, msg, wParam, lParam)
    Const LinesPerScroll = 3
    
    If msg = WM_DESTROY Then
        Call StopSubclassing(hwnd)
        
    ElseIf msg = WM_MOUSEWHEEL Then
        If eViewMode = vm_Month Then
            If m_UserCanChangeDate = False Or m_UserCanScrollMonth = False Then Exit Function
            WeekScroll = WeekScroll + IIf(wParam < 0, 1, -1)
            ProcessEvents
            Refresh
        Else
            If wParam < 0 Then
                If VScroll1.Value + LinesPerScroll > VScroll1.max Then
                    VScroll1.Value = VScroll1.max
                Else
                    VScroll1.Value = VScroll1.Value + LinesPerScroll
                End If
            Else
                If VScroll1.Value - LinesPerScroll < VScroll1.min Then
                    VScroll1.Value = VScroll1.min
                Else
                    VScroll1.Value = VScroll1.Value - LinesPerScroll
                End If
            End If
        End If
        
    ElseIf msg = WM_MOUSELEAVE Then
        Dim i As Long
        If mHotEvent <> -1 Then
            RaiseEvent EventMouseLeave(mHotEvent)
            If m_hwndTT Then SendMessageW m_hwndTT, TTM_TRACKACTIVATE, False, TI
            mHotEvent = -1
        End If
        isMouseEnter = False
        If YearCalIndex <> -1 Then
            YearCalIndex = -1
            Refresh
        End If
        
        For i = 0 To UBound(tButtons)
            If tButtons(i).State <> Normal Then
                tButtons(i).State = Normal
                Refresh
            End If
        Next

        RaiseEvent MouseLeave
    End If
End Function
 
Private Sub SetSubclassing(Obj As Object, hwnd As Long, Optional nOrdinal As Long = 1)
    Dim WindowProcAddress As Long
    Dim pObj As Long
    Dim pVar As Long
    Dim i As Long
 
    If PrevWndProc <> 0 Then Exit Sub
    
    For i = 0 To 40
        bvASM(i) = Choose(i + 1, &H55, &H8B, &HEC, &H83, &HC4, &HFC, &H8D, &H45, &HFC, &H50, &HFF, &H75, &H14, _
                                 &HFF, &H75, &H10, &HFF, &H75, &HC, &HFF, &H75, &H8, &H68, &H0, &H0, &H0, &H0, _
                                 &HB8, &H0, &H0, &H0, &H0, &HFF, &HD0, &H8B, &H45, &HFC, &HC9, &HC2, &H10, &H0)
    Next i

    pObj = ObjPtr(Obj)
    Call CopyMemory(pVar, ByVal pObj, 4)
    Call CopyMemory(WindowProcAddress, ByVal (pVar + (nOrdinal - 1) * 4 + &H7A4), 4)
    Call CopyMemory(bvASM(23), pObj, 4)
    Call CopyMemory(bvASM(28), WindowProcAddress, 4)
    PrevWndProc = SetWindowLongA(hwnd, GWL_WNDPROC, VarPtr(bvASM(0)))
End Sub

Private Sub StopSubclassing(hwnd)
    If PrevWndProc Then
        Call SetWindowLongA(hwnd, GWL_WNDPROC, PrevWndProc)
        PrevWndProc = 0
    End If
End Sub

Public Sub CenterCalenarInNow()
    Dim Value As Long
    VScroll1.Visible = False
    Dim VisibleHours As Long
    Dim HoursNow As Integer
    VisibleHours = VScroll1.Height \ RowHeight
    HoursNow = VBA.Hour(Now) * 2
    If HoursNow >= 12 Then
        Value = HoursNow - VisibleHours / 3
    End If
    If Value > VScroll1.max Then Value = VScroll1.max
    If Value < VScroll1.min Then Value = VScroll1.min
    VScroll1.Value = Value
    VScroll1.Visible = True
End Sub

Public Function GetAllEvents() As Collection
    Dim i As Long
    Set GetAllEvents = New Collection
    Form1.List1.Clear
    For i = 0 To mEventsCount - 1
        GetAllEvents.Add tEvents(i).key
        llavero(i + 1) = tEvents(i).key
        Form1.List1.AddItem Format(tEvents(i).StartTime, "mm/dd/yyyy") + Space(1) + STR(tEvents(i).key)
    Next
End Function

Public Property Let HiddeEvent(key As Long, Value As Boolean)
    Dim Index As Long
    Index = GetEventIndexByKey(key)
    If Index > -1 Then
        tEvents(Index).Hidden = Value
    End If
End Property

Public Function GetEventsFromDay(ByVal TheDay As Date) As Collection
    Dim StartDay As Date
    Dim EndDay As Date
    Dim i As Long
    
    Set GetEventsFromDay = New Collection
    
    StartDay = VBA.DateValue(TheDay)
    EndDay = DateAdd("s", -1, StartDay + 1)

    For i = 0 To mEventsCount - 1
        If (StartDay >= tEvents(i).StartTime And StartDay <= tEvents(i).EndTime) Or _
              (EndDay >= tEvents(i).StartTime And EndDay <= tEvents(i).EndTime) Or _
              tEvents(i).StartTime >= StartDay And tEvents(i).StartTime < EndDay Then
              
            GetEventsFromDay.Add tEvents(i).key
            Indice_del_evento = tEvents(i).key
        End If
    Next
    
    
    
End Function

Public Function DayHaveEvents(ByVal TheDay As Date) As Boolean
    Dim StartDay As Date
    Dim EndDay As Date
    Dim i As Long

    StartDay = VBA.DateValue(TheDay)
    EndDay = DateAdd("s", -1, StartDay + 1)

    For i = 0 To mEventsCount - 1
        If (StartDay >= tEvents(i).StartTime And StartDay <= tEvents(i).EndTime) Or _
              (EndDay >= tEvents(i).StartTime And EndDay <= tEvents(i).EndTime) Or _
              tEvents(i).StartTime >= StartDay And tEvents(i).StartTime < EndDay Then
              
            DayHaveEvents = True
            
            Exit Function
        End If
    Next
    
    
End Function

Public Sub SetStrLanguage(StrYear As String, StrMonth As String, StrWeek As String, StrDay As String, StrToday As String, StrStarts As String, StrEnds As String, StrAllDay As String)
    tButtons(2).Caption = StrYear
    tButtons(3).Caption = StrMonth
    tButtons(4).Caption = StrWeek
    tButtons(5).Caption = StrDay
    tButtons(6).Caption = StrToday
    mStrStarts = StrStarts
    mStrEnds = StrEnds
    mStrAllDay = StrAllDay
    Refresh
End Sub

Public Sub Clear()
    ReDim tEvents(0)
    mEventsCount = 0
    SelStart = -1
    SelEnd = -1
    mSelectedEvent = -1
    mHotEvent = -1
    mTodayHeight = RowHeight * 2
    Refresh
End Sub

Public Property Get DateValue() As Date
    DateValue = xDate
End Property

Public Property Let DateValue(ByVal NewValue As Date)
    xDate = VBA.DateValue(NewValue)
    ProcessEvents
    Refresh
End Property

Public Sub Update()
    QSortEvents 0, mEventsCount - 1
    ProcessEvents
    Refresh
End Sub

Public Property Get EventsCount()
    EventsCount = mEventsCount
End Property

Public Property Get SelectedEvent() As Long
    SelectedEvent = mSelectedEvent
End Property

Private Sub QSortEvents(ByVal First As Long, ByVal Last As Long)
   On Error Resume Next
    Dim Low As Long, High As Long
    Dim MidEvent As CalEvents
    Dim TempEvent As CalEvents

    Low = First
    High = Last
    If mEventsCount = 0 Then Exit Sub
    MidEvent = tEvents((First + Last) \ 2)

    Do
        While (MidEvent.StartTime = tEvents(Low).StartTime And _
              (MidEvent.EndTime - MidEvent.StartTime) < (tEvents(Low).EndTime - tEvents(Low).StartTime)) Or _
              (MidEvent.StartTime > tEvents(Low).StartTime) Or _
              (MidEvent.StartTime = tEvents(Low).StartTime And _
              (MidEvent.EndTime = tEvents(Low).EndTime) And _
              StrComp(MidEvent.Subject, tEvents(Low).Subject, vbTextCompare))

            Low = Low + 1
        Wend

        While (tEvents(High).StartTime = MidEvent.StartTime And _
              (tEvents(High).EndTime - tEvents(High).StartTime) < (MidEvent.EndTime - MidEvent.StartTime)) Or _
              (tEvents(High).StartTime > MidEvent.StartTime) Or _
              (tEvents(High).StartTime = MidEvent.StartTime And _
              (tEvents(High).EndTime = MidEvent.EndTime) And _
              StrComp(tEvents(High).Subject, MidEvent.Subject, vbTextCompare))
              
            High = High - 1
        Wend

        If Low <= High Then
            TempEvent = tEvents(Low)
            tEvents(Low) = tEvents(High)
            tEvents(High) = TempEvent
            
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
        
    If First < High Then QSortEvents First, High
    If Low < Last Then QSortEvents Low, Last
End Sub

Public Function RemoveEvent(ByVal EventKey As Long) As Boolean
    Dim i As Long, j As Long, lCount As Long
    
    If mEventsCount = 0 Then Exit Function
    
    lCount = mEventsCount - 1
    If lCount = 0 Then
        Erase tEvents
        mEventsCount = mEventsCount - 1
        RemoveEvent = mEventsCount
        ProcessEvents
        Refresh
        Exit Function
    End If

    For i = 0 To lCount
        If tEvents(i).key = EventKey Then
            For j = i To lCount - 1
                tEvents(j) = tEvents(j + 1)
            Next
            ReDim Preserve tEvents(lCount - 1)
            mEventsCount = mEventsCount - 1
            RemoveEvent = mEventsCount
            ProcessEvents
            Refresh
            Exit Function
        End If
    Next
    mEventsCount = -1
End Function

Public Function AddEvents(ByVal Subject As String, _
                            ByVal StartTime As Date, _
                            ByVal EndTime As Date, _
                            ByVal Color As Long, _
                            Optional ByVal AllDayEvent As Boolean, _
                            Optional ByVal body As String, _
                            Optional ByVal Tag As Variant, _
                            Optional ByVal IsSerie As Boolean, _
                            Optional ByVal NotifyIcon As Boolean, _
                            Optional ByVal IsPrivate As Boolean, _
                            Optional ByVal EventShowAs As eEventShowAs, _
                            Optional ByVal office As String, _
                            Optional ByVal idvacaciones As Integer, _
                            Optional ByVal email1 As String, _
                            Optional ByVal email2 As String) As Long
                            
    
    On Error Resume Next
    
          
        
    
    If StartTime > EndTime Then
        '--- SDO Add 2022-05-15 -------------------------------------
        If EndTime < CDate("1910-01-01") Then
            EndTime = StartTime
        End If
        '-------------------------------------------------------------
        If StartTime > EndTime Then
            ' AQUI77
            If Form2.op_dias(1).Value = True Then
              MsgBox "Start time cannot be greater than end time", 16, "Attention"
              'Err.Raise 100, , "Start time cannot be greater than end time"
              Exit Function
            End If
        End If
    End If
    
    ReDim Preserve tEvents(mEventsCount)
    
    
    
    
    
    With tEvents(mEventsCount)
        .ForeColor = colorx
        .StartTime = StartTime
        .EndTime = EndTime
         
        .Subject = Subject
         If Subject = "RESERVED" Then
             .office = " "
         Else
         
            .office = ubicacion_de_trabajo$
         End If
    
        .idvacaciones = ID_vacaciones$
        .AllDayEvent = AllDayEvent
        .More24Hours = DateDiff("h", StartTime, EndTime) > 24
        .body = nota$
        .key = RndKey
        If Not IsMissing(Tag) Then .Tag = Tag
        .IsSerie = IsSerie
        .NotifyIcon = NotifyIcon
        .IsPrivate = IsPrivate
        .EventShowAs = EventShowAs
        .email1 = correo_agente$
        .email2 = correo_manager$
        AddEvents = .key
    End With

    mEventsCount = mEventsCount + 1
    If m_Redraw = False Then Exit Function
    QSortEvents 0, mEventsCount - 1
    ProcessEvents
    Refresh
End Function

Public Function UpdateEventData(EventKey As Long, _
                Optional ByVal Subject As String, _
                Optional ByVal office As String, _
                Optional ByVal idvacaciones As Integer, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date, _
                Optional ByVal Color As Long, _
                Optional ByVal AllDayEvent As Boolean, _
                Optional ByVal body As String, _
                Optional Tag As Variant, _
                Optional IsSerie As Boolean, _
                Optional NotifyIcon As Boolean, _
                Optional IsPrivate As Boolean, _
                Optional ByVal EventShowAs As eEventShowAs, _
                Optional ByVal email1 As String, _
                Optional ByVal email2 As String) As Boolean
                
    On Error Resume Next
                
    Dim i As Long
    
    If StartTime > EndTime Then
        MsgBox "Start time cannot be greater than end time", 16, "Attention"
        Exit Function
    End If
    
    For i = 0 To mEventsCount - 1
        If tEvents(i).key = EventKey Then
            With tEvents(i)
                If Not IsMissing(colorx) Then .ForeColor = colorx
                If Not IsMissing(StartTime) Then .StartTime = StartTime
                If Not IsMissing(EndTime) Then .EndTime = EndTime
                If Not IsMissing(Subject) Then .Subject = Subject
                If Not IsMissing(office) Then .office = ubicacion_de_trabajo$
                If Not IsMissing(idvacaciones) Then .idvacaciones = ID_vacaciones$
                If Not IsMissing(AllDayEvent) Then .AllDayEvent = AllDayEvent
                If Not IsMissing(body) Then .body = nota$
                If Not IsMissing(Tag) Then .Tag = Tag
                If Not IsMissing(IsSerie) Then .IsSerie = IsSerie
                If Not IsMissing(NotifyIcon) Then .NotifyIcon = NotifyIcon
                If Not IsMissing(IsPrivate) Then .IsPrivate = IsPrivate
                If Not IsMissing(EventShowAs) Then .EventShowAs = EventShowAs
                If Not IsMissing(email1) Then .email1 = email1
                If Not IsMissing(email2) Then .email2 = email2
                
                 .More24Hours = DateDiff("h", .StartTime, .EndTime) > 24
                
                UpdateEventData = True
                If m_Redraw Then
                    QSortEvents 0, mEventsCount - 1
                    ProcessEvents
                    Refresh
                End If
                UpdateEventData = True
                Exit Function
            End With
        End If
    Next
End Function

Public Function GetEventData(EventKey As Long, _
                Optional Subject As String, _
                Optional StartDate As Date, _
                Optional EndDate As Date, _
                Optional Color As Long, _
                Optional AllDayEvent As Boolean, _
                Optional body As String, _
                Optional Tag As Variant, _
                Optional IsSerie As Boolean, _
                Optional NotifyIcon As Boolean, _
                Optional IsPrivate As Boolean, _
                Optional EventShowAs As eEventShowAs, _
                Optional office As String, _
                Optional idvacaciones As Integer, _
                Optional email1 As String, _
                Optional email2 As String)
                               '
    On Error Resume Next

    Dim i As Long
    
    For i = 0 To mEventsCount - 1
        If tEvents(i).key = EventKey Then
            If tEvents(i).Subject = "RESERVED" Then
              Exit For
            End If
            
        
            With tEvents(i)
                
                Color = .ForeColor
                StartDate = .StartTime
                EndDate = .EndTime
                Subject = .Subject
                office = .office
                AllDayEvent = .AllDayEvent
                body = .body
                Tag = .Tag
                IsSerie = .IsSerie
                NotifyIcon = .NotifyIcon
                IsPrivate = .IsPrivate
                EventShowAs = .EventShowAs
                idvacaciones = .idvacaciones
                email1 = .email1
                email2 = .email2
            End With
            
            
            usuario$ = Subject
            
            reservado = IsPrivate
            
            ID_vacaciones$ = tEvents(i).idvacaciones
            
            correo_agente$ = tEvents(i).email1
            
            correo_manager$ = tEvents(i).email2
            

            GetEventData = True
            
            If IsPrivate = True Then
                Exit Function
            Else
                       
                
               valido1 = 777
            
               If Format(StartDate, "mm/dd/yyyy") <> Format(EndDate, "mm/dd/yyyy") Then
                 Form2.marco_dias(1).Visible = True
                 Form2.op_dias(1).Value = True
               Else
                 Form2.marco_dias(1).Visible = False
                 Form2.op_dias(1).Value = False
               End If
            
               valido1 = 0
    
            End If
            
            
            Exit Function
        End If
    Next
End Function

Public Function GetSelectionRangeDate(StartDate As Date, EndDate As Date) As Boolean
    Dim FDOM As Date
    Dim d As Integer
    Dim Date1 As Date, Date2 As Date
    
    If (SelEnd = -1 And SelStart = -1) Then
        
      
          Exit Function
      
        
    End If

    Select Case eViewMode
        Case vm_Month

            FDOM = DateSerial(Year(xDate), Month(xDate), 1)
            d = Weekday(FDOM, m_FirstDayOfWeek) - VScroll1.Value * 7
            Date1 = DateAdd("d", -d + SelStart + 1 + 7 * WeekScroll, FDOM)
            Date2 = DateAdd("d", -d + SelEnd + 1 + 7 * WeekScroll, FDOM)
            If SelStart > SelEnd Then
                StartDate = Date2
                EndDate = Date1
            Else
                StartDate = Date1
                EndDate = Date2
            End If
        Case vm_Day
            Dim CurDate As Date
            Dim min As Long

            CurDate = DateSerial(Year(xDate), Month(xDate), Day(xDate))

            If SelEnd < 0 Then
                If SelStart + SelEnd < 0 Then SelEnd = -SelStart
                min = (SelStart + SelEnd) * 30
                StartDate = DateAdd("n", min, CurDate)
                min = min + (Abs(SelEnd) + 1) * 30
                EndDate = DateAdd("n", min, CurDate)
            Else
                min = SelStart * 30
                StartDate = DateAdd("n", min, CurDate)
                min = min + (SelEnd + 1) * 30
                EndDate = DateAdd("n", min, CurDate)
            End If
        Case vm_Week
            StartDate = DateSerial(Year(xDate), Month(xDate), Day(xDate))
            d = Weekday(xDate, m_FirstDayOfWeek)
            
            '////ALL DAY
            If SelStart <= -1 Then
                StartDate = DateAdd("D", DayIndex - d + 1, StartDate)
                EndDate = StartDate
                GetSelectionRangeDate = True
                Exit Function
            End If
            If StartDay < EndDay Then
                StartDate = DateAdd("D", StartDay - d + 1, StartDate)
                EndDate = DateAdd("D", EndDay - StartDay, StartDate)
            Else
                StartDate = DateAdd("D", EndDay - d + 1, StartDate)
                EndDate = DateAdd("D", StartDay - EndDay, StartDate)
            End If
            
            If SelEnd < 0 Then
                If SelStart + SelEnd < 0 Then SelEnd = -SelStart
                min = (SelStart + SelEnd) * 30
                StartDate = DateAdd("n", min, StartDate)
                min = min + (Abs(SelEnd) + 1) * 30
                EndDate = DateAdd("n", min, EndDate)
            Else
                min = SelStart * 30
                StartDate = DateAdd("n", min, StartDate)
                min = min + (SelEnd + 1) * 30
                EndDate = DateAdd("n", min, EndDate)
            End If
    End Select
    GetSelectionRangeDate = True
End Function

'*1
Private Sub DrawDay()

    Exit Sub

    Dim hGraphics As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim lTop As Single, lLeft As Single
    Dim i As Long, j As Long
    Dim Days As Date
    Dim lForeColor As Long
    Dim MinutesNow As Long
    Dim lBackColor As Long
    Dim eHatchStyle As HatchStyle
    
    UserControl.Cls
    GdipCreateFromHDC UserControl.hDC, hGraphics
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    'Color---------------------------
    Days = xDate
    lLeft = RowHW
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight - (VScroll1.Value * RowHeight)
    For j = 0 To 47

        If j = 47 Then
            Days = DateAdd("s", 1799, Days) '23:59:59
        Else
            Days = DateAdd("n", 30, Days)
        End If
                
        If lTop >= m_TopHeaderHeight + m_ColumnHeaderHeight Then
            lBackColor = UserControl.BackColor
            eHatchStyle = -1
            
            RaiseEvent DateBackColor(Days, lBackColor, eHatchStyle)
            
            If lBackColor <> UserControl.BackColor Then
                If eHatchStyle <> -1 Then
                    lForeColor = RGBtoARGB(lBackColor, 100)
                    lBackColor = RGBtoARGB(lBackColor, 50)
                    FillRectangleEx hGraphics, lForeColor, lBackColor, eHatchStyle, lLeft, lTop, UserControl.ScaleWidth - RowHW, RowHeight
                Else
                    GdipCreateSolidFill RGBtoARGB(lBackColor, 100), hBrush
                    GdipFillRectangleI hGraphics, hBrush, lLeft, lTop, UserControl.ScaleWidth - RowHW, RowHeight
                    GdipDeleteBrush hBrush
                End If
            End If
        End If
        lTop = lTop + RowHeight
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    lLeft = lLeft + RowWidth
    '-----------------------------------------------------
    
    '----Selection
    If SelStart > -1 Then
        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
        lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + PenWidth - (VScroll1.Value * RowHeight)
        If SelEnd < 0 Then
            If SelStart + SelEnd < 0 Then SelEnd = -SelStart
            GdipFillRectangleI hGraphics, hBrush, RowHW, lTop + (SelStart + SelEnd) * RowHeight, UserControl.ScaleWidth, (Abs(SelEnd) + 1) * RowHeight
        Else
            GdipFillRectangleI hGraphics, hBrush, RowHW, lTop + SelStart * RowHeight, UserControl.ScaleWidth, (SelEnd + 1) * RowHeight
        End If
        GdipDeleteBrush hBrush
    ElseIf SelEnd > -1 Then
        lTop = m_TopHeaderHeight + m_ColumnHeaderHeight - (VScroll1.Value * RowHeight)
        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
        GdipFillRectangleI hGraphics, hBrush, RowHW, lTop, UserControl.ScaleWidth, mTodayHeight
        GdipDeleteBrush hBrush
    End If
    '-------------
    
    'Line Minute 60
    GdipCreatePen1 RGBtoARGB(m_LinesColor, 100), PenWidth, UnitPixel, hPen
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + PenWidth - (VScroll1.Value * RowHeight)
    For i = 0 To 24
        GdipDrawLineI hGraphics, hPen, RowHW / 1.2, lTop, UserControl.ScaleWidth, lTop
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    GdipDrawLineI hGraphics, hPen, RowHW, m_TopHeaderHeight + m_ColumnHeaderHeight, RowHW, UserControl.ScaleHeight
    GdipDeletePen hPen
   
    'Line Minute 30
    GdipCreatePen1 RGBtoARGB(m_LinesColor, 50), PenWidth, UnitPixel, hPen
    GdipSetPenDashStyle hPen, DashStyleDot
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + RowHeight + PenWidth - (VScroll1.Value * RowHeight)
    For i = 0 To 23
        GdipDrawLineI hGraphics, hPen, RowHW, lTop, UserControl.ScaleWidth, lTop
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    GdipDeletePen hPen

    'Minute Text
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight - (VScroll1.Value * RowHeight)
    DrawText hGraphics, mStrAllDay, 0, lTop, RowHW, mTodayHeight, UserControl.Font, m_ForeColorAlpha, StringAlignmentCenter, StringAlignmentNear, True
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight - RowHeight * 0.5 + 1 - (VScroll1.Value * RowHeight)
    For i = 0 To 23
        DrawText hGraphics, Format(i, "00") & ":00", MarginText, lTop, RowHW, RowHeight, UserControl.Font, m_ForeColorAlpha, StringAlignmentNear, StringAlignmentCenter
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    
    '***********
    DrawEvents hGraphics
    '**************
    
    'Minute Line
    GdipCreateSolidFill RGBtoARGB(vbRed, 100), hBrush
    MinutesNow = (Hour(Now) * 60) + Minute(Now)
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + (MinutesNow * (RowHeight * 48) / 1440) - (VScroll1.Value * RowHeight)
    GdipFillRectangleI hGraphics, hBrush, RowHW, lTop, UserControl.ScaleWidth - RowHW, MarginText / 2
    GdipFillEllipse hGraphics, hBrush, RowHW - MarginText * 2, lTop - MarginText, MarginText * 2.5, MarginText * 2.5
    GdipDeleteBrush hBrush
    
    'Back Header
    GdipCreateSolidFill RGBtoARGB(UserControl.BackColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, m_TopHeaderHeight + m_ColumnHeaderHeight / 2
    GdipDeleteBrush hBrush
    
    GdipCreateSolidFill RGBtoARGB(m_HeaderColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, RowHW, m_TopHeaderHeight, UserControl.ScaleWidth - RowHW, m_ColumnHeaderHeight + PenWidth / 2
    GdipDeleteBrush hBrush
    
    'Text ColumnHeader
    lForeColor = RGBtoARGB(IIf(IsDarkColor(m_HeaderColor), vbWhite, vbBlack), 100)
    DrawText hGraphics, UCase(Format(xDate, "dd dddd")), RowHW + MarginText, m_TopHeaderHeight, UserControl.ScaleWidth - RowHW, m_ColumnHeaderHeight, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentCenter

    'Text Header
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    DrawText hGraphics, StrConv(Format(xDate, "DD MMMM YYYY"), vbProperCase), ButtonsSize * 2 + MarginText * 3, 0, UserControl.ScaleWidth, m_TopHeaderHeight, HeaderFont, lForeColor, StringAlignmentNear, StringAlignmentCenter, False

    DrawButtons hGraphics
    
    GdipDeleteGraphics hGraphics
    UserControl.Refresh
    
End Sub

Private Sub DrawButtons(hGraphics As Long)
    Dim hBrush As Long, i As Long
    Dim lBtnBackColor As Long
    
    For i = 0 To 1
        With tButtons(i).RECT
            Select Case tButtons(i).State
                Case Pressed
                    lBtnBackColor = m_HeaderColor
                Case Hot
                    lBtnBackColor = ShiftColor(m_HeaderColor, UserControl.BackColor, 127)
            End Select
            If tButtons(i).State <> Normal And m_UserCanChangeDate Then
                GdipCreateSolidFill RGBtoARGB(lBtnBackColor, 100), hBrush
                GdipFillEllipse hGraphics, hBrush, .Left, .top, .Right - .Left, .Bottom - .top
                GdipDeleteBrush hBrush
            End If
            
            DrawButtonArrow hGraphics, .Left, .top, .Right - .Left, .Bottom - .top, i = 0
            
        End With
    Next

    If m_UserCanChangeViewMode = False Then Exit Sub
    
    For i = 2 To 6
        With tButtons(i).RECT
            lBtnBackColor = IIf(eViewMode = i - 2 Or tButtons(i).State = Pressed, m_HeaderColor, UserControl.BackColor)
            If tButtons(i).State = Pressed Then
                lBtnBackColor = RGBtoARGB(lBtnBackColor, 100)
            Else
                lBtnBackColor = RGBtoARGB(lBtnBackColor, 50)
            End If
            RoundRect hGraphics, .Left, .top, .Right - .Left, .Bottom - .top, lBtnBackColor, RGBtoARGB(m_LinesColor, 100), , IIf(m_EventsRoundCorner, 6, 0)
            DrawText hGraphics, tButtons(i).Caption, .Left, .top, .Right - .Left, .Bottom - .top, UserControl.Font, RGBtoARGB(m_ForeColor, 100), StringAlignmentCenter, StringAlignmentCenter
        End With
    Next
End Sub

'*1
Private Sub DrawWeek()
    
   

    Dim hGraphics As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim lTop As Single, lLeft As Single
    Dim i As Long, j As Long
    Dim d As Integer
    Dim Days As Date
    Dim lForeColor As Long
    Dim MinutesNow As Long
    Dim SD As Long, ED As Long
    Dim lBackColor As Long
    Dim eHatchStyle As HatchStyle
    
    RowHeight = 20 * nScale
    RowHW = 50 * nScale
    RowWidth = (UserControl.ScaleWidth - RowHW - VScroll1.Width) / 7
    EventHeight = RowHeight
    
    d = Weekday(xDate, m_FirstDayOfWeek)

    UserControl.Cls
    GdipCreateFromHDC UserControl.hDC, hGraphics
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

    'Color---------------------------
    lLeft = RowHW
    For i = 0 To 6
        Days = DateAdd("d", -d + i + 1, xDate)
        
        lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight - (VScroll1.Value * RowHeight)
        For j = 0 To 47
 
            If j = 47 Then
                Days = DateAdd("s", 1799, Days) '23:59:59
            Else
                Days = DateAdd("n", 30, Days)
            End If
                
             If lTop >= m_TopHeaderHeight + m_ColumnHeaderHeight - RowHeight Then
                lBackColor = UserControl.BackColor
                eHatchStyle = -1
                
                RaiseEvent DateBackColor(Days, lBackColor, eHatchStyle)
                
                If lBackColor <> UserControl.BackColor Then
                    If eHatchStyle <> -1 Then
                        lForeColor = RGBtoARGB(lBackColor, 100)
                        lBackColor = RGBtoARGB(lBackColor, 50)
                        FillRectangleEx hGraphics, lForeColor, lBackColor, eHatchStyle, lLeft, lTop, RowWidth, RowHeight
                    Else
                        GdipCreateSolidFill RGBtoARGB(lBackColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, lLeft, lTop, RowWidth, RowHeight
                        GdipDeleteBrush hBrush
                    End If
                End If
            End If
            lTop = lTop + RowHeight
            If lTop > UserControl.ScaleHeight Then Exit For
        Next
        lLeft = lLeft + RowWidth
    Next
    '-----------------------------------------------------
    '----Selection
    If SelStart > -1 Then
        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
        lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + PenWidth - (VScroll1.Value * RowHeight)

        If EndDay - StartDay > 0 Then
            SD = StartDay ' + 1
            ED = EndDay '- 1
        Else
            SD = EndDay
            ED = StartDay
        End If
 
        If EndDay - StartDay <> 0 And EndDay <> -1 Then
            
            If SelStart = 0 Or SelStart + SelEnd = 0 Then
                GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * SD, lTop - mTodayHeight, RowWidth + 1, (48) * RowHeight + mTodayHeight
            Else
                If EndDay - StartDay < 0 Then
                    If SelStart + SelEnd < 0 Then SelEnd = -SelStart
                    GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * SD, lTop + (SelStart + SelEnd) * RowHeight, RowWidth + 1, (48 - (SelStart + SelEnd)) * RowHeight
                Else
                    GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * SD, lTop + SelStart * RowHeight, RowWidth + 1, (48 - SelStart) * RowHeight
                End If
            End If

            For i = SD + 1 To ED - 1
                GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * i, lTop - mTodayHeight, RowWidth + 1, 48 * RowHeight + mTodayHeight
            Next

            If SelEnd = 47 Or SelStart = 47 Then
                GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * ED, lTop - mTodayHeight, RowWidth + 1, (48) * RowHeight + mTodayHeight
            Else
                If EndDay - StartDay < 0 Then
                    If SelStart + SelEnd < 0 Then SelEnd = -SelStart
                    GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * ED, lTop, RowWidth, (SelStart + 1) * RowHeight
                Else
                    GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * ED, lTop, RowWidth + 1, (SelStart + SelEnd + 1) * RowHeight
                End If
            End If
        Else
            If SelEnd < 0 Then
                If SelStart + SelEnd < 0 Then SelEnd = -SelStart
                GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * DayIndex, lTop + (SelStart + SelEnd) * RowHeight, RowWidth, (Abs(SelEnd) + 1) * RowHeight
            Else
                GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * DayIndex, lTop + SelStart * RowHeight, RowWidth, (SelEnd + 1) * RowHeight
            End If
        End If
        
        GdipDeleteBrush hBrush
        
    ElseIf SelEnd > -1 Then

        lTop = m_TopHeaderHeight + m_ColumnHeaderHeight - (VScroll1.Value * RowHeight)
        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
        GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * DayIndex, lTop, RowWidth, mTodayHeight + 1
        GdipDeleteBrush hBrush
    End If
    '-------------
    
    'Line Minute 60
    GdipCreatePen1 RGBtoARGB(m_LinesColor, 100), PenWidth, UnitPixel, hPen
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + PenWidth - (VScroll1.Value * RowHeight)
    For i = 0 To 24
        If lTop >= 0 Then
            GdipDrawLineI hGraphics, hPen, RowHW / 1.2, lTop, UserControl.ScaleWidth, lTop
        End If
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    GdipDrawLineI hGraphics, hPen, RowHW, m_TopHeaderHeight + m_ColumnHeaderHeight, RowHW, UserControl.ScaleHeight
    GdipDeletePen hPen
    
    For i = 1 To 7
        GdipCreatePen1 RGBtoARGB(m_LinesColor, 100), PenWidth, UnitPixel, hPen
        GdipDrawLineI hGraphics, hPen, RowHW + RowWidth * i, m_TopHeaderHeight + m_ColumnHeaderHeight, RowHW + RowWidth * i, UserControl.ScaleHeight
        GdipDeletePen hPen
    Next
   
    'Line Minute 30
    GdipCreatePen1 RGBtoARGB(m_LinesColor, 50), PenWidth, UnitPixel, hPen
    GdipSetPenDashStyle hPen, DashStyleDot
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + RowHeight + PenWidth - (VScroll1.Value * RowHeight)
    For i = 0 To 23
        If lTop >= 0 Then
            GdipDrawLineI hGraphics, hPen, RowHW, lTop, UserControl.ScaleWidth, lTop
        End If
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    GdipDeletePen hPen

    'Minute Text
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight - (VScroll1.Value * RowHeight)
    DrawText hGraphics, mStrAllDay, 0, lTop, RowHW, mTodayHeight, UserControl.Font, m_ForeColorAlpha, StringAlignmentFar, StringAlignmentNear, True
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight - RowHeight * 0.5 + 1 - (VScroll1.Value * RowHeight)
    For i = 0 To 23
        If lTop >= 0 Then
            DrawText hGraphics, Format(i, "00") & ":00", MarginText, lTop, RowHW, RowHeight, UserControl.Font, m_ForeColorAlpha, StringAlignmentNear, StringAlignmentCenter
        End If
        lTop = lTop + RowHeight * 2
        If lTop > UserControl.ScaleHeight Then Exit For
    Next
    
    '***********
    DrawEvents hGraphics
    '**************
    
    'Back Header
    GdipCreateSolidFill RGBtoARGB(UserControl.BackColor, 100), hBrush
     GdipCreateSolidFill RGBtoARGB(UserControl.BackColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, m_TopHeaderHeight + m_ColumnHeaderHeight / 2
    GdipDeleteBrush hBrush
    
    'ColumnHeader
    GdipCreateSolidFill RGBtoARGB(m_HeaderColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, RowHW, m_TopHeaderHeight, UserControl.ScaleWidth - RowHW, m_ColumnHeaderHeight + PenWidth / 2
    GdipDeleteBrush hBrush
    
    'Text ColumnHeader
    For i = 0 To 7
        Days = DateAdd("d", -d + i + 1, xDate)

        If Days = Date Then
            GdipCreateSolidFill RGBtoARGB(ShiftColor(m_HeaderColor, UserControl.BackColor, 127), 100), hBrush
            GdipFillRectangleI hGraphics, hBrush, RowHW + RowWidth * i, m_TopHeaderHeight, RowWidth, m_ColumnHeaderHeight
            GdipDeleteBrush hBrush
        End If

        lForeColor = RGBtoARGB(IIf(IsDarkColor(m_HeaderColor), vbWhite, vbBlack), 100)
        DrawText hGraphics, UCase(Format(Days, "dd dddd")), RowHW + RowWidth * i + MarginText, m_TopHeaderHeight, RowWidth - MarginText, m_ColumnHeaderHeight, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentCenter
        
        'GdipCreatePen1 RGBtoARGB(m_LinesColor, 100), PenWidth, UnitPixel, hPen
        'GdipDrawLineI hGraphics, hPen, RowHW + RowWidth * i, m_TopHeaderHeight, RowHW + RowWidth * i, m_TopHeaderHeight + m_ColumnHeaderHeight
        'GdipDeletePen hPen
    Next
    
    'Minute Line
    MinutesNow = (Hour(Now) * 60) + Minute(Now)
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight + (MinutesNow * (RowHeight * 48) / 1440) - (VScroll1.Value * RowHeight)
    If lTop > m_TopHeaderHeight + m_ColumnHeaderHeight Then
        GdipCreateSolidFill RGBtoARGB(vbRed, 100), hBrush
        GdipFillRectangleI hGraphics, hBrush, RowHW, lTop, UserControl.ScaleWidth - RowHW, MarginText / 2
        GdipFillEllipse hGraphics, hBrush, RowHW - MarginText * 2, lTop - MarginText, MarginText * 2.5, MarginText * 2.5
        GdipDeleteBrush hBrush
    End If

    DrawButtons hGraphics
    
    'Text Header
    Dim sText As String
    Dim FDay As Date, LDay As Date
        
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    FDay = xDate - d + 1
    LDay = xDate - d + 7
    sText = Day(FDay)
    If Month(FDay) <> Month(LDay) Then sText = StrConv(Format(FDay, "D MMMM"), vbProperCase)
    sText = sText & " - " & Day(LDay) & " " & StrConv(Format(LDay, "MMMM YYYY"), vbProperCase)
    DrawText hGraphics, sText, ButtonsSize * 2 + MarginText * 3, 0, UserControl.ScaleWidth, m_TopHeaderHeight, HeaderFont, lForeColor, StringAlignmentNear, StringAlignmentCenter, False

    GdipDeleteGraphics hGraphics
    UserControl.Refresh
End Sub

Private Sub DrawIconRepeat(hGraphics As Long, Left As Long, top As Long, Size As Long, ByVal Color As Long)
    Dim hPen As Long
    GdipCreatePen1 Color, nScale, &H2, hPen
    GdipSetPenEndCap hPen, LineCapArrowAnchor
    GdipDrawArc hGraphics, hPen, Left, top, Size, Size, -20, -150
    GdipDrawArc hGraphics, hPen, Left, top, Size, Size, 160, -150
    GdipDeletePen hPen
End Sub

Private Sub DrawIconBell(hGraphics As Long, Left As Long, top As Long, BoxSize As Long, ByVal Color As Long)
    Dim hPen As Long
    Dim PF(8) As POINTF
    
    GdipCreatePen1 Color, 1 * nScale, &H2, hPen
    
    PF(0).X = Left + 0.484 * BoxSize: PF(0).Y = top + 0.136 * BoxSize
    PF(1).X = Left + 0.76 * BoxSize: PF(1).Y = top + 0.384 * BoxSize
    PF(2).X = Left + 0.76 * BoxSize: PF(2).Y = top + 0.6 * BoxSize
    PF(3).X = Left + 0.852 * BoxSize: PF(3).Y = top + 0.76 * BoxSize
    PF(4).X = Left + 0.852 * BoxSize: PF(4).Y = top + 0.832 * BoxSize
    PF(5).X = Left + 0.128 * BoxSize: PF(5).Y = top + 0.832 * BoxSize
    PF(6).X = Left + 0.128 * BoxSize: PF(6).Y = top + 0.78 * BoxSize
    PF(7).X = Left + 0.216 * BoxSize: PF(7).Y = top + 0.6 * BoxSize
    PF(8).X = Left + 0.216 * BoxSize: PF(8).Y = top + 0.384 * BoxSize
    GdipDrawClosedCurve hGraphics, hPen, PF(0), 9
    PF(0).X = Left + 0.392 * BoxSize: PF(0).Y = top + 0.896 * BoxSize
    PF(1).X = Left + 0.488 * BoxSize: PF(1).Y = top + 0.976 * BoxSize
    PF(2).X = Left + 0.584 * BoxSize: PF(2).Y = top + 0.896 * BoxSize
    GdipDrawClosedCurve hGraphics, hPen, PF(0), 3
    PF(0).X = Left + 0.484 * BoxSize: PF(0).Y = top + 0.016 * BoxSize
    PF(1).X = Left + 0.56 * BoxSize: PF(1).Y = top + 0.084 * BoxSize
    PF(2).X = Left + 0.48 * BoxSize: PF(2).Y = top + 0.144 * BoxSize
    PF(3).X = Left + 0.4 * BoxSize: PF(3).Y = top + 0.084 * BoxSize
    GdipDrawClosedCurve hGraphics, hPen, PF(0), 4
    
    GdipDeletePen hPen
End Sub

Private Function DrawIconPadlock(ByVal hGraphics As Long, ByVal X As Long, ByVal Y As Long, BoxSize As Long, ByVal Color As Long)
    Dim hPen As Long

    GdipSetClipRectI hGraphics, X + BoxSize / 4, Y + BoxSize * 0.4, BoxSize / 2, BoxSize * 0.6, CombineModeExclude
    RoundRect hGraphics, X + BoxSize / 3.5, Y, BoxSize / 2.5, BoxSize, , Color, 1, 100

    GdipResetClip hGraphics
    RoundRect hGraphics, X + BoxSize / 6, Y + BoxSize * 0.4, BoxSize / 1.5, BoxSize * 0.6, , Color, 1, BoxSize / 10
    GdipCreatePen1 Color, nScale, &H2, hPen
    GdipDrawLineI hGraphics, hPen, X + BoxSize / 2, Y + BoxSize * 0.65, X + BoxSize / 2, Y + BoxSize * 0.8
    GdipDeletePen hPen
End Function

Private Sub DrawYear()
    Dim hGraphics As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim sDay As String
    Dim i As Long, j As Long, c As Long
    Dim d As Integer
    Dim FDOM As Date
    Dim Days As Date
    Dim lForeColor As Long
    Dim X As Long, Y As Long
    Dim CalSize As Long
    Dim lLeft  As Long, lTop As Long
    Dim lWidth As Long, lHeight As Long
    Dim PartWidth As Single
    Dim MT As Long
    Dim TW As Long, TH As Long
    Dim FontMonth As StdFont
    Dim Cols As Long
    Dim HH As Long

    Set FontMonth = New StdFont
    FontMonth.Name = UserControl.Name
    FontMonth.Size = FontMonth.Size * 1.5

    CalSize = 200 * nScale
    ReDim YearCalRects(11)
    
    HH = m_TopHeaderHeight
    
    lWidth = UserControl.ScaleWidth - MarginText * 4
    If VScroll1.Visible Then lWidth = lWidth - VScroll1.Width

    lHeight = UserControl.ScaleHeight - MarginText * 4 - HH
    MT = MarginText * 2
    Cols = lWidth \ CalSize
    
    If Cols > 6 Then
        Cols = 6
    ElseIf Cols = 5 Then
        Cols = 4
    ElseIf Cols = 0 Then
        Cols = 1
    End If
    
    PartWidth = lWidth / Cols
    
    TW = (CalSize - MT) \ 7
    TH = TextHeight("A") * 1.4
    
    If (12 \ Cols) * CalSize > lHeight Then
        lTop = HH + MT
    Else
        lTop = HH + lHeight / 2 - ((12 \ Cols) * CalSize) / 2
    End If
    
    For Y = 0 To 12 \ Cols - 1
        lLeft = MT
        For X = 1 To Cols
       
            With YearCalRects(i)
                .Left = lLeft + PartWidth / 2 - CalSize / 2
                .top = lTop - VScroll1.Value
                .Right = .Left + CalSize
                .Bottom = lTop + CalSize - VScroll1.Value
            End With

            i = i + 1
            lLeft = lLeft + PartWidth
        Next
        lTop = lTop + CalSize
    Next
    
    If VScroll1.Visible = False Then
        If lTop > UserControl.ScaleHeight Then
            VScroll1.Visible = True
            If VScroll1.Visible = True Then
                DrawYear
                Exit Sub
            End If
        End If
    ElseIf VScroll1.Visible = True Then
        If lTop < UserControl.ScaleHeight Then
            VScroll1.Visible = False
            DrawYear
            Exit Sub
        End If
    End If
    
    VScroll1.max = lTop - lHeight - HH
            
    UserControl.Cls
    GdipCreateFromHDC UserControl.hDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    lForeColor = RGBtoARGB(m_ForeColor, 100)


    If YearCalIndex <> -1 Then
'        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
'        With YearCalRects(YearCalIndex)
'            GdipFillRectangleI hGraphics, hBrush, .Left, .Top, .Right - .Left, .Bottom - .Top
'        End With
'        GdipDeleteBrush hBrush
        With YearCalRects(YearCalIndex)
            RoundRect hGraphics, .Left, .top, .Right - .Left, .Bottom - .top, RGBtoARGB(m_SelectionColor, 70), , , IIf(m_EventsRoundCorner, 10, 0)
        End With
    End If

    For i = 0 To 11
        FDOM = DateSerial(Year(xDate), i + 1, 1)
        d = Weekday(FDOM, m_FirstDayOfWeek)
        With YearCalRects(i)
            DrawText hGraphics, UCase(Format(FDOM, "mmmm")), .Left + MT, .top + MT, CalSize, CalSize, FontMonth, lForeColor, StringAlignmentNear, StringAlignmentNear
            lTop = .top + FontMonth.Size * 2 * nScale + MarginText
            lLeft = .Left + MT
            For j = 0 To 6
                Days = PvFirstDayOfWeek(VBA.DateSerial(Year(xDate), i + 1, 1)) + j
                sDay = Left$(UCase(Format(Days, "DDD")), 2) & Space(1)
                DrawText hGraphics, sDay, lLeft, lTop, TW, TW, UserControl.Font, lForeColor, StringAlignmentCenter, StringAlignmentNear
                lLeft = lLeft + TW
            Next
            
            lTop = lTop + TH
            
            c = 0
            For Y = 1 To 6
                lLeft = .Left + MT
                For X = 1 To 7
                    c = c + 1
                    Days = FDOM - d + c 'DateAdd("d", -D + C, FDOM)

                    If Days = Date Then
'                        GdipCreatePen1 RGBtoARGB(m_HeaderColor, 90), 2 * nScale, &H2, hPen
'                        GdipDrawRectangleI hGraphics, hPen, lLeft, lTop - TW / 8, TW, TW
'                        GdipDeletePen hPen
                        RoundRect hGraphics, lLeft, lTop, TW, TW, , lForeColor, 2, IIf(m_EventsRoundCorner, 100, 0)
                    End If
        
                    If Month(Days) = i + 1 Then
                        If DayHaveEvents(Days) Then
                            UserControl.Font.Bold = True
                            'GdipCreatePen1 RGBtoARGB(m_HeaderColor, 50), nScale * 2, &H2, hPen
                            'GdipDrawLineI hGraphics, hPen, lLeft, lTop + TH / 1.2, lLeft + TW, lTop + TH / 1.2
                            'GdipDeletePen hPen
                            RoundRect hGraphics, lLeft, lTop, TW, TW, RGBtoARGB(m_HeaderColor, 50), , , IIf(m_EventsRoundCorner, 100, 0)
                            
'                            GdipCreateSolidFill RGBtoARGB(m_HeaderColor, 50), hBrush
'                            GdipFillRectangleI hGraphics, hBrush, lLeft, lTop - TW / 8, TW, TW
'                            GdipDeleteBrush hBrush
                        End If

                        DrawText hGraphics, Day(Days), lLeft, lTop, TW, TW, UserControl.Font, lForeColor, StringAlignmentCenter, StringAlignmentCenter
                        UserControl.Font.Bold = False
                    End If
                    lLeft = lLeft + TW
                Next
                lTop = lTop + TH
            Next
        End With
    Next
    
    'Back Header
    GdipCreateSolidFill RGBtoARGB(UserControl.BackColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, m_TopHeaderHeight + m_ColumnHeaderHeight / 2
    GdipDeleteBrush hBrush
    
    DrawButtons hGraphics
    
    'Text Header
    DrawText hGraphics, Year(xDate), ButtonsSize * 2 + MarginText * 3, 0, UserControl.ScaleWidth, m_TopHeaderHeight, HeaderFont, lForeColor, StringAlignmentNear, StringAlignmentCenter, False

    GdipDeleteGraphics hGraphics
    UserControl.Refresh
End Sub

'*1
Private Sub DrawMonth()
    Dim hGraphics As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim lTop As Long, lLeft As Long
    Dim sDay As String
    Dim i As Long, j As Long, N As Long
    Dim d As Integer
    Dim FDOM As Date
    Dim Days As Date
    Dim lForeColor As Long
    Dim eHatchStyle As HatchStyle
    Dim lBackColor As Long
    
    RowWidth = UserControl.ScaleWidth / 7
    RowHeight = (UserControl.ScaleHeight - m_ColumnHeaderHeight - m_TopHeaderHeight) / m_WeekCounts
    
    UserControl.Cls
    GdipCreateFromHDC UserControl.hDC, hGraphics
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    'ColumnHeader-------------------
    GdipCreateSolidFill RGBtoARGB(m_HeaderColor, 100), hBrush
    GdipFillRectangleI hGraphics, hBrush, 0, m_TopHeaderHeight, UserControl.ScaleWidth, m_ColumnHeaderHeight
    GdipDeleteBrush hBrush
    
    If IsDateInRange(Date, FirstDate, LastDate) Then
        d = Weekday(Date, m_FirstDayOfWeek) - 1
        GdipCreateSolidFill RGBtoARGB(ShiftColor(m_HeaderColor, UserControl.BackColor, 127), 100), hBrush
        GdipFillRectangleI hGraphics, hBrush, RowWidth * d, m_TopHeaderHeight, RowWidth, m_ColumnHeaderHeight
        GdipDeleteBrush hBrush
    End If
    '--------------------------------
    'Color---------------------------
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight
    FDOM = DateSerial(Year(xDate), Month(xDate), 1)
    d = Weekday(FDOM, m_FirstDayOfWeek) '- VScroll1.Value * 7
    FDOM = DateAdd("d", 7 * WeekScroll, FDOM)
    N = 0
    For i = 0 To m_WeekCounts - 1
        lLeft = 0
        For j = 1 To 7
            
            N = N + 1
            
            lBackColor = UserControl.BackColor
            eHatchStyle = -1
            Days = DateAdd("d", -d + N, FDOM)
            RaiseEvent DateBackColor(Days, lBackColor, eHatchStyle)
            If lBackColor <> UserControl.BackColor Then
                If eHatchStyle <> -1 Then
                    lForeColor = RGBtoARGB(lBackColor, 100)
                    lBackColor = RGBtoARGB(lBackColor, 50)
                    FillRectangleEx hGraphics, lForeColor, lBackColor, eHatchStyle, lLeft, lTop, RowWidth, RowHeight
                Else
                    GdipCreateSolidFill RGBtoARGB(lBackColor, 100), hBrush
                    GdipFillRectangleI hGraphics, hBrush, lLeft, lTop, RowWidth, RowHeight
                    GdipDeleteBrush hBrush
                End If
            End If
            lLeft = lLeft + RowWidth
        Next
        lTop = lTop + RowHeight
    Next
    '--------------------------------
    'Selecion------------------------
    If SelStart > -1 Then
        GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
        For i = SelStart To SelEnd Step IIf(SelStart < SelEnd, 1, -1)
            lTop = m_ColumnHeaderHeight + m_TopHeaderHeight + (i \ 7) * RowHeight
            lLeft = (i Mod 7) * RowWidth
            GdipFillRectangleI hGraphics, hBrush, lLeft, lTop, RowWidth, RowHeight
        Next
        GdipDeleteBrush hBrush
    End If
    '------------------
    'Lineas------------
    GdipCreatePen1 RGBtoARGB(m_LinesColor, 100), PenWidth, UnitPixel, hPen
    lTop = m_TopHeaderHeight
    lTop = lTop + m_ColumnHeaderHeight
    For i = 0 To 5
        GdipDrawLineI hGraphics, hPen, 0, lTop, UserControl.ScaleWidth, lTop
        lTop = lTop + RowHeight
    Next
    GdipDrawLineI hGraphics, hPen, 0, lTop - PenWidth, UserControl.ScaleWidth, lTop - PenWidth
    
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight
    lLeft = 0

    For i = 1 To 7
        GdipDrawLineI hGraphics, hPen, lLeft, lTop, lLeft, UserControl.ScaleHeight
        lLeft = lLeft + RowWidth
    Next
    GdipDrawLineI hGraphics, hPen, lLeft - PenWidth, lTop, lLeft - PenWidth, UserControl.ScaleHeight
    GdipDeletePen hPen
    '----------------
    
    lLeft = 0
    For i = 0 To 6
        sDay = StrConv(Format(DateAdd("d", i, PvFirstDayOfWeek(Date)), "DDDD"), vbProperCase)
        If i = d And IsDateInRange(Date, FirstDate, LastDate) Then
            lForeColor = RGBtoARGB(IIf(IsDarkColor(m_HeaderColor), vbWhite, vbBlack), 100)
            UserControl.FontBold = True
        Else
            lForeColor = RGBtoARGB(m_ForeColor, 100)
            UserControl.FontBold = False
        End If
        
        DrawText hGraphics, sDay, lLeft + MarginText, m_TopHeaderHeight, RowWidth, m_ColumnHeaderHeight, UserControl.Font, lForeColor, StringAlignmentCenter, StringAlignmentCenter, False
        lLeft = lLeft + RowWidth
    Next
    
    FDOM = DateSerial(Year(xDate), Month(xDate), 1)
    d = Weekday(FDOM, m_FirstDayOfWeek)
    FDOM = DateAdd("d", 7 * WeekScroll, FDOM)
    
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight
    lLeft = 0
    
    For i = 1 To 7 * m_WeekCounts
        Days = DateAdd("d", -d + i, FDOM)

        If Month(FDOM) = Month(Days) Then
            If Days = Date Then
                GdipCreateSolidFill RGBtoARGB(m_HeaderColor, 100), hBrush
                GdipFillRectangleI hGraphics, hBrush, lLeft, lTop, RowWidth, 3 * nScale
                GdipDeleteBrush hBrush
                If IsDarkColor(m_HeaderColor) = IsDarkColor(UserControl.BackColor) Then
                    lForeColor = RGBtoARGB(m_ForeColor, 100)
                Else
                    lForeColor = RGBtoARGB(m_HeaderColor, 60)
                End If
                UserControl.FontBold = True
            Else
                lForeColor = RGBtoARGB(m_ForeColor, 100)
                UserControl.FontBold = False
            End If
        Else
            lForeColor = m_ForeColorAlpha 'RGBtoARGB(m_ForeColor, 50)
        End If
        
        If Day(Days) = 1 Then
            sDay = Format(Days, "dd mmm")
            UserControl.Font.Bold = True
        Else
            sDay = Day(Days)
            UserControl.Font.Bold = False
        End If
        
        If SelStart < SelEnd Then
            If i - 1 >= SelStart And i - 1 <= SelEnd Then
                lForeColor = RGBtoARGB(IIf(IsDarkColor(m_SelectionColor), vbWhite, vbBlack), 100)
            End If
        Else
            If i - 1 >= SelEnd And i - 1 <= SelStart Then
                lForeColor = RGBtoARGB(IIf(IsDarkColor(m_SelectionColor), vbWhite, vbBlack), 100)
            End If
        End If
        
        DrawText hGraphics, sDay, lLeft + MarginText, lTop + MarginText, RowWidth, 16 * nScale, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentCenter, False
        If mEventsCount > 0 Then
            If tGrid(i).HaveHideEvents Then
                DrawDropDown hGraphics, lLeft + PenWidth, lTop + RowHeight - 10 * nScale, RowWidth - PenWidth * 2, 10 * nScale
            End If
        End If
        
        If i Mod 7 = 0 Then
            lTop = lTop + RowHeight
            lLeft = 0
        Else
            lLeft = lLeft + RowWidth
        End If
    
    Next i

    '***********
    DrawEvents hGraphics
    '**************

    DrawButtons hGraphics
    
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    Dim sDisplayDate As String

    Days = DateAdd("d", -d + 7, FDOM)

    sDisplayDate = StrConv(Format(Days, "MMMM YYYY"), vbProperCase)
    DrawText hGraphics, sDisplayDate, ButtonsSize * 2 + MarginText * 3, 0, UserControl.ScaleWidth, m_TopHeaderHeight, HeaderFont, lForeColor, StringAlignmentNear, StringAlignmentCenter, False

    GdipDeleteGraphics hGraphics
    UserControl.Refresh
End Sub

Private Sub FillRectangleEx(hGraphics As Long, Color1 As Long, Color2 As Long, HatchStyle As HatchStyle, ByVal Left As Long, ByVal top As Long, ByVal Width As Long, ByVal Height As Long)
    Dim hBrush As Long
    GdipCreateHatchBrush HatchStyle, Color1, Color2, hBrush
    GdipFillRectangleI hGraphics, hBrush, Left, top, Width, Height
    GdipDeleteBrush hBrush
End Sub

'*2
Private Sub DrawEvents(hGraphics As Long)
    '-----------Eventos
    Dim lCount As Long
    Dim i As Long, j As Long
    Dim lForeColor As Long
    Dim hBrush As Long, hPen As Long
    Dim lTop As Long, lLeft As Long
    Dim tRect As RECT
    Dim Radius As Double
    
    Radius = IIf(m_EventsRoundCorner, 6, 0)
    
    
    
    For i = 0 To mEventsCount - 1
        With tEvents(i)
           If .RectsCount > 0 And .Hidden = False Then
                If IsDarkColor(.ForeColor) Then lForeColor = vbWhite Else lForeColor = vbBlack
                lForeColor = RGBtoARGB(IIf(IsDarkColor(.ForeColor), vbWhite, vbBlack), 100)
                
                
                
'                If tEvents(i).Key = mSelectedEvent Then
'                     GdipCreateSolidFill RGBtoARGB(.ForeColor, 40), hBrush
'                Else
'                     GdipCreateSolidFill RGBtoARGB(.ForeColor, 70), hBrush
'                End If
                
                For j = 0 To UBound(.Rects)
                    tRect = .Rects(j)
    
                    With tRect
                         lLeft = .Left + 10 * nScale
                        If .Bottom > 0 And .top > 0 And lLeft > 0 And .Right > 0 Then
                             
                             lTop = .top - (VScroll1.Value * RowHeight)
    
                             'GdipFillRectangleI hGraphics, hBrush, .Left, lTop, .Right - .Left, .Bottom - .Top
                             
                             If tEvents(i).key = mSelectedEvent Then
                                RoundRect hGraphics, .Left, lTop, .Right - .Left, .Bottom - .top, RGBtoARGB(tEvents(i).ForeColor, 40), RGBtoARGB(ShiftColor(tEvents(i).ForeColor, vbBlack, 200), 100), 2, Radius
                                
                                
                             Else
                                RoundRect hGraphics, .Left, lTop, .Right - .Left, .Bottom - .top, RGBtoARGB(tEvents(i).ForeColor, 70), , , Radius
                             End If
                             
                             Select Case tEvents(i).EventShowAs
                                 Case ESA_Busy
                                     FillRectangleEx hGraphics, RGBtoARGB(tEvents(i).ForeColor, 100), RGBtoARGB(tEvents(i).ForeColor, 100), 0, .Left + nScale, lTop + nScale, 8 * nScale, .Bottom - .top - nScale * 2
                                 Case ESA_Free
                                     FillRectangleEx hGraphics, RGBtoARGB(UserControl.BackColor, 100), RGBtoARGB(UserControl.BackColor, 100), 0, .Left + nScale, lTop + nScale, 8 * nScale, .Bottom - .top - nScale * 2
                                 Case [ESA_Out of office]
                                     FillRectangleEx hGraphics, RGBtoARGB(&H800080, 100), RGBtoARGB(&H800080, 100), 0, .Left + nScale, lTop + nScale, 8 * nScale, .Bottom - .top - nScale * 2
                                 Case [ESA_Working elsewhere]
                                     FillRectangleEx hGraphics, RGBtoARGB(tEvents(i).ForeColor, 100), RGBtoARGB(UserControl.BackColor, 100), 14, .Left + nScale, lTop + nScale, 8 * nScale, .Bottom - .top - nScale * 2
                                 Case ESA_Tentative
                                     FillRectangleEx hGraphics, RGBtoARGB(tEvents(i).ForeColor, 100), RGBtoARGB(UserControl.BackColor, 100), 23, .Left + nScale, lTop + nScale, 8 * nScale, .Bottom - .top - nScale * 2
                             End Select
    
'                             If tEvents(i).Key = mSelectedEvent Then
'                                 GdipCreatePen1 RGBtoARGB(ShiftColor(tEvents(i).ForeColor, vbBlack, 200), 100), 2 * nScale, &H2, hPen
'                                 GdipDrawRectangleI hGraphics, hPen, .Left, lTop, .Right - .Left, .Bottom - .Top
'                                 GdipDeletePen hPen
'                             End If
    
                             UserControl.Font.Size = UserControl.Font.Size - 2
                             Dim TH As Long, StrTime As String, EH As Long
                             EH = (.Bottom - .top)
                             TH = UserControl.TextHeight("A")
                             
                             If EH < TH Or ((tEvents(i).AllDayEvent Or tEvents(i).More24Hours) And eViewMode <> vm_Month) Then
                                 lTop = .top - (VScroll1.Value * RowHeight) + EH / 2 - 8 * nScale
                             End If
                             
                            
                             
                             If tEvents(i).IsSerie Then
                                 DrawIconRepeat hGraphics, lLeft + MarginText / 2, lTop + MarginText, 8 * nScale, lForeColor
                                 lTop = lTop + 12 * nScale
                             End If
                             
                             If tEvents(i).IsPrivate Then
                                 If lTop + (12 * nScale) > .Bottom Then
                                     lTop = lTop - 12 * nScale
                                     lLeft = lLeft + 12 * nScale
                                 End If
                                 DrawIconPadlock hGraphics, lLeft + MarginText / 2, lTop + MarginText, 8 * nScale, lForeColor
                                 lTop = lTop + 12 * nScale
                             End If
                             
                             If tEvents(i).NotifyIcon Then
                                 If lTop + (12 * nScale) > .Bottom Then
                                     lTop = lTop - 12 * nScale
                                     lLeft = lLeft + 12 * nScale
                                 End If
                                 DrawIconBell hGraphics, lLeft + MarginText / 2, lTop + MarginText, 8 * nScale, lForeColor
                                 lTop = lTop + 12 * nScale
                             End If
                             
                             If tEvents(i).NotifyIcon Or tEvents(i).IsPrivate Or tEvents(i).IsSerie Then
                                 lLeft = lLeft + (12 * nScale)
                             End If
                             
                             
                             If EH < TH Or (tEvents(i).AllDayEvent Or tEvents(i).More24Hours) And eViewMode <> vm_Month Then
                                 lTop = .top - (VScroll1.Value * RowHeight) + RowHeight / 2 - 8 * nScale
                             Else
                                 lTop = .top - (VScroll1.Value * RowHeight)
                             End If
                             
                             
                             
                             
                             
                             If EH < TH Then
                                 DrawText hGraphics, tEvents(i).Subject, lLeft, lTop, .Right - lLeft, .Bottom - .top, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentCenter, False
                             Else
                                 
                                 
                              
                                 UserControl.Font.Bold = True
                                 DrawText hGraphics, tEvents(i).Subject, lLeft, lTop, .Right - lLeft, .Bottom - .top, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentNear, False
                                 UserControl.Font.Bold = False
                                 
                             End If
                             
                             If EH > TH * 2 Then
                                 lTop = lTop + TH
                                 StrTime = Format(tEvents(i).StartTime, "hh:nn") & " - " & Format(tEvents(i).EndTime, "hh:nn")
                                 
                                 If Format(tEvents(i).StartTime, "mm/dd") = "12/25" Or Format(tEvents(i).StartTime, "mm/dd") = "01/01" Or Format(tEvents(i).StartTime, "mm/dd") = "11/" + Format(dia_thanks, "00") Or Format(tEvents(i).StartTime, "mm/dd") = "11/" + Format(dia_thanks2, "00") Then
                                 
                                     DrawText hGraphics, "", lLeft, lTop + 3, .Right - lLeft, .Bottom - .top, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentNear, False
                                     
                                 Else
                                 
                                     DrawText hGraphics, tEvents(i).office, lLeft, lTop + 3, .Right - lLeft, .Bottom - .top, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentNear, False
                                 ' DrawText hGraphics, StrTime, lLeft, lTop, .Right - lLeft, .Bottom - .Top, UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentNear, False
                                 End If
                             End If
                             
                             If EH > TH * 3 And Len(tEvents(i).body) Then
                                 lTop = lTop + TH + MarginText
                                 DrawText hGraphics, tEvents(i).body, lLeft, lTop, .Right - lLeft, .Bottom - lTop - MarginText - (VScroll1.Value * RowHeight), UserControl.Font, lForeColor, StringAlignmentNear, StringAlignmentNear, True
                             End If
                                   
                             UserControl.Font.Size = UserControl.Font.Size + 2
                             
                        End If
                    End With
                Next
                
'                GdipDeleteBrush hBrush
            End If
        End With
    Next
End Sub

Private Function IsDateInRange(ByVal TheDate As Date, ByVal MinDate As Date, ByVal MaxDate As Date) As Boolean
    TheDate = VBA.DateSerial(Year(TheDate), Month(TheDate), Day(TheDate))
    MinDate = VBA.DateSerial(Year(MinDate), Month(MinDate), Day(MinDate))
    MaxDate = VBA.DateSerial(Year(MaxDate), Month(MaxDate), Day(MaxDate))
    
    If TheDate >= MinDate And TheDate <= MaxDate Then
        IsDateInRange = True
    End If
End Function

Private Function GetNroOfEventsAllDayForDay(TheDay As Date) As Long
    Dim i As Long, lCount As Long

    For i = 0 To mEventsCount - 1
        With tEvents(i)
           If IsDateInRange(TheDay, .StartTime, .EndTime) Then
                If .AllDayEvent Or .More24Hours Then '.StartTime <= TheDay And .EndTime >= DateAdd("d", 1, TheDay) Then
                    lCount = lCount + 1
                End If
            End If
        End With
    Next
    GetNroOfEventsAllDayForDay = lCount
End Function

Private Function GetEventWidth(N As Long, R As RECT) As Long
    Dim lTop As Long, i  As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim lMax As Long
    
    lTop = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight
    
    lStart = (R.top - lTop) \ EventHeight
    lEnd = (R.Bottom - lTop) \ EventHeight
    If lStart < 0 Then lStart = 0
    If lEnd > 48 Then lEnd = 48

    For i = lStart To lEnd - 1
        If GridWeek(N).Rows(i) > lMax Then
            lMax = GridWeek(N).Rows(i)
        End If
    Next
    
    If lMax > 0 Then
        If eViewMode = vm_Week Then
            R.Right = R.Left + (UserControl.ScaleWidth - RowHW - VScroll1.Width - MarginText * 14) / 7 / lMax
        Else
            R.Right = R.Left + (UserControl.ScaleWidth - RowHW - VScroll1.Width - MarginText * 2) / lMax
        End If
    End If
End Function

'*3
Private Sub ProcessEvents()
    Dim i As Long, j As Long, N As Long, c As Long
    Dim d As Integer
    Dim FDOM As Date
    Dim Days As Date
    Dim NumDays As Long
    Dim lFirst  As Long, X As Single, Y As Single
    Dim lCount As Long
    Dim EventsStartMin As Long, EventsEndMin As Long
    Dim lTodayHeight As Long
    Dim SpaceClick As Long
    Dim ColWidth As Single
    Dim NroEvOfAllDay As Long
    Dim lStart As Long, lEnd As Long

    If mEventsCount = 0 Then Exit Sub
    
    SpaceClick = 8 * nScale
    
    lStart = -1
    lEnd = -1
    
    '////////////WEEK/////////////////
    If eViewMode = vm_Week Then
  
        RowHeight = 20 * nScale
        RowHW = 50 * nScale
        EventHeight = RowHeight
        
        d = Weekday(xDate, m_FirstDayOfWeek)
        ColWidth = (UserControl.ScaleWidth - RowHW - VScroll1.Width) / 7
        
        'Calcula rango de eventos usados en el mes actual
        FirstDate = xDate - d + 1
        LastDate = FirstDate + 7

        For j = 0 To mEventsCount - 1
            With tEvents(j)
               If IsDateInRange(tEvents(j).StartTime, FirstDate, LastDate) Or IsDateInRange(tEvents(j).EndTime, FirstDate, LastDate) Then
                    If lStart = -1 Then lStart = j
               Else
                    If lStart <> -1 And tEvents(j).EndTime > FirstDate Then lEnd = j: Exit For
               End If
            End With
        Next
        
        If lStart = -1 Then lStart = 0
        If lEnd <= 1 Then lEnd = mEventsCount - 1
        
        ReDim tGrid(6)
        ReDim GridWeek(6)
        
        For N = 0 To 6
            Days = FirstDate + N ' DateAdd("d", -D + n, xDate)
            NumDays = GetNroOfEventsAllDayForDay(Days)
            If NumDays > NroEvOfAllDay Then NroEvOfAllDay = NumDays
            ReDim GridWeek(N).Rows(47)
        Next

        mTodayHeight = NroEvOfAllDay * (EventHeight + MarginText) + MarginText
        
        If mTodayHeight < RowHeight * 2 Then mTodayHeight = RowHeight * 2

        For i = 0 To mEventsCount - 1
            Erase tEvents(i).Rects
            tEvents(i).RectsCount = 0
        Next
        
        For N = 0 To 6
            Days = xDate - d + N + 1 ' DateAdd("d", -D + n + 1, xDate)
            For j = lStart To lEnd
                With tEvents(j)
                   If IsDateInRange(Days, .StartTime, .EndTime) And .Hidden = False Then
                        ReDim Preserve .Rects(.RectsCount)
                        If .AllDayEvent Or .More24Hours Then '.StartTime <= Days And .EndTime >= DateAdd("d", 1, Days) Then
                            Y = m_TopHeaderHeight + m_ColumnHeaderHeight + MarginText

                            With .Rects(.RectsCount)
                                .Left = RowHW + ColWidth * N
                                .top = Y + tGrid(N).EventsCount * (EventHeight + MarginText)
                                .Right = .Left + ColWidth
                                .Bottom = .top + EventHeight
                                lTodayHeight = .Bottom - Y + MarginText * 2
                            End With

                            tGrid(N).EventsCount = tGrid(N).EventsCount + 1
                        Else
                            If .StartTime < Days Then
                                EventsStartMin = 0
                            Else
                                EventsStartMin = (Hour(.StartTime) * 60) + Minute(.StartTime)
                            End If

                            If .EndTime >= Days + 1 Then
                                EventsEndMin = 60 * 24
                            Else
                                EventsEndMin = (Hour(.EndTime) * 60) + Minute(.EndTime)
                            End If
                            
                            Y = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight '- (VScroll1.Value * RowHeight)

                            For c = (EventsStartMin * 48 / 1440) To (EventsEndMin * 48 / 1440) - 1
                                GridWeek(N).Rows(c) = GridWeek(N).Rows(c) + 1
                            Next
                            
                            RowWidth = ColWidth
                            
                            With .Rects(.RectsCount)
                                .Left = RowHW + (ColWidth * N) '+ (RowWidth * (tEvents(J).CollPos))
                                .top = Y + (EventsStartMin * (RowHeight * 48) / 1440)
                                .Right = .Left + ColWidth - MarginText
                                .Bottom = Y + (EventsEndMin * (RowHeight * 48) / 1440)
                                If .Bottom - .top < EventHeight / 2 Then
                                  .Bottom = .top + EventHeight / 2
                                End If
                            End With
                        End If
                        .RectsCount = .RectsCount + 1
                   End If
                End With
            Next
        Next
        
        For N = 0 To 6
            Days = xDate - d + N + 1
            Dim R As RECT
            For j = lStart To lEnd
                With tEvents(j)
                    If Not .AllDayEvent And Not .More24Hours Then
                        If IsDateInRange(Days, .StartTime, .EndTime) Then
                            For i = 0 To tEvents(j).RectsCount - 1
                                Call GetEventWidth(N, tEvents(j).Rects(i))
                                AcomodateRect j, tEvents(j).Rects(i)
                            Next
                        End If
                    End If
                End With
            Next
        Next

        VScroll1.max = 49 + mTodayHeight \ RowHeight - (VScroll1.Height) \ RowHeight
        Exit Sub
    End If
    
    '/////////////DAY/////////////////
    If eViewMode = vm_Day Then
  
        RowHeight = 20 * nScale
        RowHW = 50 * nScale
        EventHeight = RowHeight
      
        ReDim tGrid(48)
        ReDim GridWeek(0)
        ReDim GridWeek(0).Rows(47)
        
        NroEvOfAllDay = GetNroOfEventsAllDayForDay(xDate)
        mTodayHeight = NroEvOfAllDay * (EventHeight + MarginText) + MarginText
        If mTodayHeight < RowHeight * 2 Then mTodayHeight = RowHeight * 2

        For j = 0 To mEventsCount - 1
            With tEvents(j)
               Erase .Rects
               .RectsCount = 0
               
               If IsDateInRange(xDate, .StartTime, .EndTime) And .Hidden = False Then

                    ReDim Preserve .Rects(0)
                    .RectsCount = 1
                    
                    If .AllDayEvent Or .More24Hours Then ' (.StartTime <= VBA.DateValue(xDate) And .EndTime >= DateAdd("d", 1, VBA.DateValue(xDate))) Then
                        EventHeight = RowHeight

                        Y = m_TopHeaderHeight + m_ColumnHeaderHeight + MarginText

                        With .Rects(lCount)
                            .Left = RowHW
                            .top = Y + tGrid(0).EventsCount * (EventHeight + MarginText)
                            .Right = UserControl.ScaleWidth
                            .Bottom = .top + EventHeight
                            lTodayHeight = .Bottom - Y + MarginText * 2
                        End With
     
                        tGrid(0).EventsCount = tGrid(0).EventsCount + 1
                    Else
                        If .StartTime < xDate Then
                            EventsStartMin = 0
                        Else
                            EventsStartMin = (Hour(.StartTime) * 60) + Minute(.StartTime)
                        End If
                        
                        If .EndTime >= xDate + 1 Then
                            EventsEndMin = 60 * 24
                        Else
                            EventsEndMin = (Hour(.EndTime) * 60) + Minute(.EndTime)
                        End If

                        Y = m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight '- (VScroll1.Value * RowHeight)
                    
                        For c = (EventsStartMin * 48 / 1440) To (EventsEndMin * 48 / 1440) - 1
                            GridWeek(0).Rows(c) = GridWeek(0).Rows(c) + 1
                        Next
                    
                        RowWidth = (UserControl.ScaleWidth - RowHW - VScroll1.Width - SpaceClick) '/ tEvents(J).CollCount
                        With .Rects(lCount)
                            .Left = RowHW '+ (RowWidth * (tEvents(J).CollPos - 1))
                            .top = Y + (EventsStartMin * (RowHeight * 48) / 1440)
                            .Right = .Left + RowWidth
                            .Bottom = Y + (EventsEndMin * (RowHeight * 48) / 1440)
                            
                            If .Bottom - .top < EventHeight / 2 Then
                              .Bottom = .top + EventHeight / 2
                            End If
                        End With
                    End If
               End If
            End With
        Next j

        For j = 0 To mEventsCount - 1
            With tEvents(j)
                If Not .AllDayEvent And Not .More24Hours Then
                    If IsDateInRange(xDate, .StartTime, .EndTime) Then
                        For i = 0 To tEvents(j).RectsCount - 1
                            Call GetEventWidth(0, tEvents(j).Rects(i))
                            AcomodateRect j, tEvents(j).Rects(i)
                        Next
                    End If
                End If
            End With
        Next

        VScroll1.max = 49 + mTodayHeight \ RowHeight - (VScroll1.Height) \ RowHeight
        Exit Sub
    End If

    '////////MONTH/////////////
    If eViewMode = vm_Month Then
        RowWidth = UserControl.ScaleWidth / 7
        RowHeight = (UserControl.ScaleHeight - m_ColumnHeaderHeight - m_TopHeaderHeight) / m_WeekCounts
        'RowHeight = 150
        EventHeight = 32 * nScale    ' 60
        
        FDOM = DateSerial(Year(xDate), Month(xDate), 1)
        
        d = Weekday(FDOM, m_FirstDayOfWeek)
    
        ReDim tGrid(1 To 7 * m_WeekCounts)

        For i = 0 To mEventsCount - 1
            Erase tEvents(i).Rects
            tEvents(i).RectsCount = 0
        Next

        FirstDate = FDOM - d + 1 + 7 * WeekScroll ' DateAdd("d", -D + 1 + 7 * WeekScroll, FDOM)
        LastDate = FDOM - d + 7 * m_WeekCounts + 7 * WeekScroll 'DateAdd("d", -D + 7 * m_WeekCounts + 7 * WeekScroll, FDOM)
        
        mes_selecto_en_calendario$ = Format(xDate, "mm")

        'Calcula rango de eventos usados en el mes actual
        For j = 0 To mEventsCount - 1
            With tEvents(j)
               If IsDateInRange(tEvents(j).StartTime, FirstDate, LastDate) Or IsDateInRange(tEvents(j).EndTime, FirstDate, LastDate) Then
                    If lStart = -1 Then lStart = j
               Else
                    If lStart <> -1 And tEvents(j).EndTime > FirstDate Then lEnd = j: Exit For
               End If
            End With
        Next
        If lStart = -1 Then lStart = 0
        If lEnd <= 1 Then lEnd = mEventsCount - 1
        
        For i = 1 To 7 * m_WeekCounts
        
            Days = FDOM - d + i + 7 * WeekScroll 'DateAdd("d", -D + i + 7 * WeekScroll, FDOM)
            '-----------Eventos
            lCount = 0
    
            For j = lStart To lEnd
                With tEvents(j)
                   If IsDateInRange(Days, .StartTime, .EndTime) And .Hidden = False Then
                        If .RectsCount = 0 Then
                            Dim t As Long
                            ReDim Preserve .Rects(lCount)
                            .RectsCount = lCount + 1
                            
                            NumDays = DateDiff("d", Days, .EndTime) + 1
                      
                            For N = i To i + NumDays - 1
                                If N > 7 * m_WeekCounts Then Exit For
                                tGrid(N).EventsCount = tGrid(N).EventsCount + 1
                            Next
              
                            Y = m_TopHeaderHeight + m_ColumnHeaderHeight + MarginText + EventHeight + ((i - 1) \ 7) * RowHeight '- RowHeight
    
                            lFirst = (i - 1) Mod 7

                            X = lFirst * RowWidth
                            
                            .Rects(lCount).Right = X
                            t = lFirst
    
                            For N = i To i + NumDays - 1
                                If N > 7 * m_WeekCounts Then Exit For
                                
                                t = t + 1
                                If t = 8 Then
                                    t = 1
                                    Y = Y + RowHeight
                                    X = PenWidth
                                    lCount = lCount + 1
                                    ReDim Preserve .Rects(lCount)
                                    .RectsCount = lCount + 1
                                End If
                                
                                With .Rects(lCount)
                                    .Left = X
                                    .top = Y
                                    .Right = .Right + RowWidth - PenWidth
                                    .Bottom = .top + EventHeight
                                End With
                                   AcomodateRect j, .Rects(lCount)
                                With .Rects(lCount)
                                    If .Bottom > Y + RowHeight - 10 * nScale - MarginText - EventHeight Then
                                        tGrid(N).HaveHideEvents = True
                                        .Bottom = 0
                                    End If
                                End With
                            Next
                        End If
                   End If
                End With
            Next j
        Next i
        
        ' carga fechas de contratacion
        
        carga_aniversarios
        
        
    End If
    
    
    
End Sub

Private Function AcomodateRect(NroEvent As Long, R As RECT)
    Dim i As Long, j As Long, w As Long
    Dim RetRect As RECT
    For i = 0 To NroEvent - 1
        For j = 0 To tEvents(i).RectsCount - 1
            If IntersectRect(RetRect, tEvents(i).Rects(j), R) And tEvents(i).Hidden = False Then
                If eViewMode = vm_Month Then
                    R.top = tEvents(i).Rects(j).Bottom + MarginText
                    R.Bottom = R.top + (EventHeight)
                    AcomodateRect i, R
                Else
                    w = R.Right - R.Left
                    R.Left = tEvents(i).Rects(j).Right
                    R.Right = R.Left + w
                End If
            End If
        Next
    Next
End Function

Private Sub Timer1_Timer()
    If eViewMode = vm_Day Or eViewMode = vm_Week Then
        Refresh
    End If
    
    
    
End Sub
    
Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GdiplusStartupInput
    GdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)

    nScale = GetWindowsDPI
    m_Redraw = True
    SelStart = -1
    SelEnd = -1
    mSelectedEvent = -1
    mHotEvent = -1
    YearCalIndex = -1
    xDate = Date
    PenWidth = 1 * nScale
    m_TopHeaderHeight = 48 * nScale
    m_ColumnHeaderHeight = 24 * nScale
    MarginText = 4 * nScale
    EventHeight = 16 * nScale
    ButtonsSize = 32 * nScale
    RowHeight = 20 * nScale
    mTodayHeight = RowHeight * 2
    m_WeekCounts = 5
    Set HeaderFont = New StdFont
    HeaderFont.Name = "Arial"
    HeaderFont.Size = 18
    RowHW = 50 * nScale
    
    With tButtons(0).RECT
        .Left = MarginText
        .top = m_TopHeaderHeight / 2 - ButtonsSize / 2
        .Right = .Left + ButtonsSize
        .Bottom = .top + ButtonsSize
    End With
    
    With tButtons(1).RECT
        .Left = MarginText * 2 + ButtonsSize
        .top = m_TopHeaderHeight / 2 - ButtonsSize / 2
        .Right = .Left + ButtonsSize
        .Bottom = .top + ButtonsSize
    End With
    
    tButtons(2).Caption = "Year" '"Año"
    tButtons(3).Caption = "Month" '"Mes"
    tButtons(4).Caption = "Week" '"Semana"
    tButtons(5).Caption = "Day" '"Día"
    tButtons(6).Caption = "Today" '"Hoy"
    
    
    mStrStarts = "Start" '"Comienza"
    mStrEnds = "End" '"Finaliza"
    mStrAllDay = "All day" '"Todo el día"
End Sub

Private Function RndKey() As Long
    RndKey = CLng((10000000 - 1 + 1) * Rnd + 1)
End Function


Private Function RoundRect(ByVal hGraphics As Long, ByVal X As Long, ByVal Y As Long, _
                            ByVal Width As Long, ByVal Height As Long, _
                            Optional BackColor As Long, Optional BorderColor As Long, _
                            Optional BordeWidth As Long = 1, Optional Radius As Double = 6) As Boolean

    Dim hBrush As Long
    Dim hPen As Long, mPath As Long
    Dim d As Double

    d = Radius * nScale
    
    If d >= Width Then d = Width - 1
    If d >= Height Then d = Height - 1

    Call GdipCreatePath(&H0, mPath)
    
    If d = 0 Then
        GdipAddPathRectangleI mPath, X, Y, Width, Height
    Else
        GdipAddPathArcI mPath, X, Y, d, d, 180, 90
        GdipAddPathArcI mPath, X + Width - d, Y, d, d, 270, 90
        GdipAddPathArcI mPath, X + Width - d, Y + Height - d, d, d, 0, 90
        GdipAddPathArcI mPath, X, Y + Height - d, d, d, 90, 90
        GdipAddPathLineI mPath, X, Y + d / 2, X, Y + Height - d / 2
    End If
    
    If Not IsMissing(BackColor) Then
        GdipCreateSolidFill BackColor, hBrush
        GdipFillPath hGraphics, hBrush, mPath
        GdipDeleteBrush hBrush
    End If
    
    If Not IsMissing(BorderColor) Then
        GdipCreatePen1 BorderColor, BordeWidth * nScale, &H2, hPen
        GdipDrawPath hGraphics, hPen, mPath
        GdipDeletePen hPen
    End If
    
    Call GdipDeletePath(mPath)
End Function

Private Function DrawDropDown(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    Dim hPen As Long
    Dim hBrush As Long
    Dim lPenColor As Long
    
    lPenColor = IIf(IsDarkColor(m_DropDownColor), vbWhite, vbBlack)
    GdipCreateSolidFill RGBtoARGB(m_DropDownColor, 100), hBrush
    GdipCreatePen1 RGBtoARGB(lPenColor, 100), PenWidth, UnitPixel, hPen
    GdipSetPenStartCap hPen, LineCapRound
    GdipFillRectangleI hGraphics, hBrush, X, Y, Width, Height

    GdipDrawLineI hGraphics, hPen, X + Width / 2, Y + Height / 1.5, X + Width / 2 - Height / 3, Y + Height / 3
    GdipDrawLineI hGraphics, hPen, X + Width / 2, Y + Height / 1.5, X + Width / 2 + Height / 3, Y + Height / 3
    
    GdipDeletePen hPen
    GdipDeleteBrush hBrush
End Function

Private Sub DrawButtonArrow(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional DirectionLeft As Boolean)
    Dim hPen As Long
    Dim PenWidth  As Long
    Dim Opacity As Long
    PenWidth = 2 * nScale

    If m_UserCanChangeDate Then
        Opacity = 100
    Else
        Opacity = 10
    End If
     
    GdipCreatePen1 RGBtoARGB(m_ForeColor, Opacity), PenWidth, UnitPixel, hPen
    If DirectionLeft Then
        GdipSetPenStartCap hPen, LineCapRound
        GdipDrawLineI hGraphics, hPen, X + Width / 2.5, Y + Height / 2, X + Width / 1.6, Y + Height / 3.33
        GdipDrawLineI hGraphics, hPen, X + Width / 2.5, Y + Height / 2, X + Width / 1.6, Y + Height / 1.43
    Else
        GdipSetPenEndCap hPen, LineCapRound
        GdipDrawLineI hGraphics, hPen, X + Width / 2.5, Y + Height / 3.33, X + Width / 1.6, Y + Height / 2
        GdipDrawLineI hGraphics, hPen, X + Width / 2.5, Y + Height / 1.43, X + Width / 1.6, Y + Height / 2
    End If
    
    GdipDeletePen hPen
End Sub

Private Function DrawText(ByVal hGraphics As Long, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As StringAlignment, Optional VAlign As StringAlignment, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RECTF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long

    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
        GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisCharacter
    End If
        
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        
        
    lFontSize = MulDiv(10.2, GetDeviceCaps(hDC, LOGPIXELSY), 72)
    lFontSize = 12
    'lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)

    layoutRect.Left = X: layoutRect.top = Y '* nScale
    layoutRect.Width = Width: layoutRect.Height = Height  '* nScale


    GdipCreateSolidFill ForeColor, hBrush
    
        

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    
    

    
    
    GdipDrawString hGraphics, StrPtr(Text), -1, hFont, layoutRect, hFormat, hBrush
    
   
    
    
    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
End Function

Private Function GetWindowsDPI() As Double
    Dim hDC As Long, LPX  As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Public Property Get ShowToolTipEvents() As Boolean
   ShowToolTipEvents = m_ShowToolTipEvents
End Property

Public Property Let ShowToolTipEvents(ByVal NewValue As Boolean)
    m_ShowToolTipEvents = NewValue
    PropertyChanged "ShowToolTipEvents"
End Property

Public Property Get FirstDayOfWeek() As VbDayOfWeek
    UserCanScrollMonth = m_FirstDayOfWeek
End Property

Public Property Let FirstDayOfWeek(ByVal NewValue As VbDayOfWeek)
    m_FirstDayOfWeek = NewValue
    PropertyChanged "FirstDayOfWeek"
    ProcessEvents
    Refresh
End Property

Public Property Get UserCanScrollMonth() As Boolean
    UserCanScrollMonth = m_UserCanScrollMonth
End Property

Public Property Let UserCanScrollMonth(ByVal NewValue As Boolean)
    m_UserCanScrollMonth = NewValue
    PropertyChanged "UserCanScrollMonth"
End Property

Public Property Get UserCanChangeDate() As Boolean
    UserCanChangeDate = m_UserCanChangeDate
End Property

Public Property Let UserCanChangeDate(ByVal NewValue As Boolean)
    m_UserCanChangeDate = NewValue
    PropertyChanged "UserCanChangeDate"
    Refresh
End Property

Public Property Get UserCanChangeViewMode() As Boolean
    UserCanChangeViewMode = m_UserCanChangeViewMode
End Property

Public Property Let UserCanChangeViewMode(ByVal NewValue As Boolean)
    m_UserCanChangeViewMode = NewValue
    PropertyChanged "UserCanChangeViewMode"
    Refresh
End Property

Public Property Get UserCanChangeEvents() As Boolean
    UserCanChangeEvents = m_UserCanChangeEvents
End Property

Public Property Let UserCanChangeEvents(ByVal NewValue As Boolean)
    m_UserCanChangeEvents = NewValue
    PropertyChanged "UserCanChangeEvents"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get DropDownColor() As OLE_COLOR
    DropDownColor = m_DropDownColor
End Property

Public Property Let DropDownColor(ByVal New_Color As OLE_COLOR)
    m_DropDownColor = New_Color
    PropertyChanged "DropDownColor"
    Refresh
End Property

Public Property Get SelectionColor() As OLE_COLOR
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_Color As OLE_COLOR)
    m_SelectionColor = New_Color
    PropertyChanged "SelectionColor"
    Refresh
End Property

Public Property Get HeaderColor() As OLE_COLOR
    HeaderColor = m_HeaderColor
End Property

Public Property Let HeaderColor(ByVal New_Color As OLE_COLOR)
    m_HeaderColor = New_Color
    PropertyChanged "HeaderColor"
    Refresh
End Property

Public Property Get LinesColor() As OLE_COLOR
    LinesColor = m_LinesColor
End Property

Public Property Let LinesColor(ByVal New_Color As OLE_COLOR)
    m_LinesColor = New_Color
    PropertyChanged "LinesColor"
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_Color As OLE_COLOR)
    m_ForeColor = New_Color
    m_ForeColorAlpha = RGBtoARGB(m_ForeColor, 60)
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    UserControl.BackColor = New_Color
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get EventsRoundCorner() As Boolean
    EventsRoundCorner = m_EventsRoundCorner
End Property

Public Property Let EventsRoundCorner(ByVal New_Value As Boolean)
    m_EventsRoundCorner = New_Value
    PropertyChanged "EventsRoundCorner"
    Refresh
End Property



Public Property Let ViewMode(ByVal NewValue As EnuViewMode)
    eViewMode = NewValue
    Select Case eViewMode
        Case vm_Year
            VScroll1.Visible = False
        Case vm_Month
            VScroll1.Visible = False
            VScroll1.Value = 0
        Case vm_Week
            VScroll1.Visible = True
        Case vm_Day
            VScroll1.Visible = True
    End Select
    PropertyChanged "ViewMode"
    ProcessEvents
    Refresh
End Property

Public Property Get ViewMode() As EnuViewMode
    ViewMode = eViewMode
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal New_Value As Boolean)
    m_Redraw = New_Value
    If m_Redraw = True Then
        QSortEvents 0, mEventsCount - 1
        ProcessEvents
        Refresh
    End If
End Property

Private Sub UserControl_InitProperties()
    Set m_Font = UserControl.Ambient.Font
    m_ForeColor = vbButtonText
    UserControl.BackColor = vbWindowBackground
    
    m_HeaderColor = vbActiveTitleBar
   
    
    m_LinesColor = vb3DLight
    m_SelectionColor = vbHighlight
    m_DropDownColor = vbButtonFace
    m_UserCanChangeEvents = True
    m_UserCanChangeViewMode = True
    m_UserCanChangeDate = True
    UserCanScrollMonth = True
    m_FirstDayOfWeek = vbUseSystemDayOfWeek
    m_ShowToolTipEvents = True
    m_EventsRoundCorner = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If m_UserCanChangeDate = False Or m_UserCanScrollMonth = False Then Exit Sub
    
    If KeyCode = vbKeyDown Then
        WeekScroll = WeekScroll + 1
        ProcessEvents
        Refresh
    ElseIf KeyCode = vbKeyUp Then
        WeekScroll = WeekScroll - 1
        ProcessEvents
        Refresh
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Public Sub Refresh()
    If m_Redraw Then
        Select Case eViewMode
            Case vm_Day: DrawDay
            Case vm_Week: DrawWeek
            Case vm_Month: DrawMonth
            Case vm_Year: DrawYear
        End Select
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_ForeColorAlpha = RGBtoARGB(m_ForeColor, 60)
        UserControl.BackColor = .ReadProperty("BackColor", vbWindowBackground)
        
        m_HeaderColor = .ReadProperty("HeaderColor", &H80000002)          ' vbActiveTitleBar)
        
        m_LinesColor = .ReadProperty("LinesColor", vb3DLight)
        m_SelectionColor = .ReadProperty("SelectionColor", vbHighlight)
        m_DropDownColor = .ReadProperty("DropDownColor", vbButtonFace)
        eViewMode = .ReadProperty("ViewMode", vm_Month)
        m_UserCanChangeEvents = .ReadProperty("UserCanChangeEvents", True)
        m_UserCanChangeViewMode = .ReadProperty("UserCanChangeViewMode", True)
        m_UserCanChangeDate = .ReadProperty("UserCanChangeDate", True)
        m_UserCanScrollMonth = .ReadProperty("UserCanScrollMonth", True)
        m_FirstDayOfWeek = .ReadProperty("FirstDayOfWeek", vbUseSystemDayOfWeek)
        m_ShowToolTipEvents = .ReadProperty("ShowToolTipEvents", True)
        m_EventsRoundCorner = .ReadProperty("EventsRoundCorner", True)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If
    
    End With
    
    If eViewMode = vm_Day Or eViewMode = vm_Week Then
        If App.LogMode Then SetSubclassing Me, UserControl.hwnd
    End If
    If eViewMode = vm_Day Or eViewMode = vm_Week Then
        
        VScroll1.Visible = True
    End If
    
    
    ProcessEvents
    'Refresh
    
    If m_ShowToolTipEvents Then
        m_hwndTT = CreateWindowExW(WS_EX_TOPMOST, StrPtr(TOOLTIPS_CLASS), 0, TTS_ALWAYSTIP, 0, 0, 0, 0, hwnd, 0, App.hInstance, ByVal 0)

        If m_hwndTT Then
            With TI
                .cbSize = Len(TI)
                .uFlags = TTF_IDISHWND
                .uId = hwnd
            End With
 
            SendMessageLongW m_hwndTT, TTM_ADDTOOLW, 0&, VarPtr(TI)

        End If
    End If
End Sub

Private Sub UserControl_Show()
    If eViewMode = vm_Day Or eViewMode = vm_Week Then
        CenterCalenarInNow
        Refresh
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", m_Font, UserControl.Ambient.Font
        .WriteProperty "ForeColor", m_ForeColor, vbButtonText
        .WriteProperty "BackColor", UserControl.BackColor, vbWindowBackground
        .WriteProperty "HeaderColor", m_HeaderColor, vbActiveTitleBar
        .WriteProperty "LinesColor", m_LinesColor, vb3DLight
        .WriteProperty "SelectionColor", m_SelectionColor, vbHighlight
        .WriteProperty "DropDownColor", m_DropDownColor, vbButtonFace
        .WriteProperty "ViewMode", eViewMode, vm_Month
        .WriteProperty "UserCanChangeEvents", m_UserCanChangeEvents, True
        .WriteProperty "UserCanScrollMonth", m_UserCanScrollMonth, True
        .WriteProperty "UserCanChangeViewMode", m_UserCanChangeViewMode, True
        .WriteProperty "UserCanChangeDate", m_UserCanChangeDate, True
        .WriteProperty "EventsRoundCorner", m_EventsRoundCorner, True
        .WriteProperty "FirstDayOfWeek", m_FirstDayOfWeek, vbUseSystemDayOfWeek
        .WriteProperty "ShowToolTipEvents", m_ShowToolTipEvents, True
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    VScroll1.Move UserControl.ScaleWidth - VScroll1.Width, m_TopHeaderHeight + m_ColumnHeaderHeight + PenWidth, VScroll1.Width, UserControl.ScaleHeight - m_TopHeaderHeight - m_ColumnHeaderHeight
    On Error GoTo 0

    Dim i As Long, lLeft As Long, ButtonWidth As Long
    lLeft = UserControl.ScaleWidth - MarginText
    For i = 2 To 6
        ButtonWidth = UserControl.TextWidth(tButtons(i).Caption) + MarginText * 2
        lLeft = lLeft - ButtonWidth - MarginText * 2
        With tButtons(i).RECT
            .Left = lLeft
            .top = m_TopHeaderHeight / 2 - ButtonsSize / 2
            .Right = .Left + ButtonWidth
            .Bottom = .top + ButtonsSize
        End With
    Next
    
    RowHeight = 20 * nScale
    VScroll1.max = 49 + mTodayHeight \ RowHeight - (VScroll1.Height) \ RowHeight
    If m_Redraw = False Then Exit Sub
    ProcessEvents
    Refresh
End Sub

Private Sub UserControl_Terminate()
    If hCur Then DestroyCursor hCur
    StopSubclassing UserControl.hwnd
    
    If m_hwndTT Then
        Call SendMessageW(m_hwndTT, TTM_DELTOOLW, 0, TI)
        Call DestroyWindow(m_hwndTT)
        m_hwndTT = 0
    End If
    
    GdiplusShutdown GdipToken
End Sub

Private Function RGBtoARGB(ByVal rgbColor As Long, ByVal Opacity As Long) As Long 'By LaVople
    If (rgbColor And &H80000000) Then rgbColor = GetSysColor(rgbColor And &HFF&)
    RGBtoARGB = (rgbColor And &HFF00&) Or (rgbColor And &HFF0000) \ &H10000 Or (rgbColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function


Private Sub UserControl_Click()
     RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, j As Long
    Dim Index As Long
    Dim YY As Single
    
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If m_hwndTT Then SendMessageW m_hwndTT, TTM_TRACKACTIVATE, False, TI
    
    If Button <> vbLeftButton Then Exit Sub
    
    bMouseDownInCal = Y > m_TopHeaderHeight + m_ColumnHeaderHeight
    
    'Click en un evento
    If mEventsCount > 0 And Y > m_TopHeaderHeight + m_ColumnHeaderHeight Then
        For i = UBound(tEvents) To 0 Step -1
            With tEvents(i)
                If .RectsCount > 0 Then
                    For j = 0 To UBound(.Rects)
                        If PtInRect(.Rects(j), X, Y + (VScroll1.Value * RowHeight)) Then
                            
                            mSelectedEvent = .key
                            PointMdown.X = X: PointMdown.Y = Y
                            SelStart = -1
                            SelEnd = -1
                            Refresh

                            If m_UserCanChangeEvents = False Then Exit Sub
                            
                            If eViewMode = vm_Day Or eViewMode = vm_Week Then
                                YY = Y + (VScroll1.Value * RowHeight)
                                If YY - .Rects(j).top <= MarginText Then
                                    mDragKey = .key
                                    mStartSize = True
                                    mSizeDirection = 0
                                    MousePointer = vbSizeNS
                                ElseIf .Rects(j).Bottom - YY <= MarginText Then
                                    mDragKey = .key
                                    mStartSize = True
                                    mSizeDirection = 1
                                    MousePointer = vbSizeNS
                                End If
                            Else
                                If X - .Rects(j).Left <= MarginText Then
                                    mDragKey = .key
                                    mStartSize = True
                                    mSizeDirection = 0
                                    MousePointer = vbSizeWE
                                ElseIf .Rects(j).Right - X <= MarginText Then
                                    mDragKey = .key
                                    mStartSize = True
                                    mSizeDirection = 1
                                    MousePointer = vbSizeWE
                                End If
                            End If
                            Exit Sub
                        End If
                    Next
                End If
            End With
        Next
    End If
    mSelectedEvent = -1
    
    If eViewMode = vm_Week And Button = vbLeftButton Then
        YY = Y + (VScroll1.Value * RowHeight)
        
        If YY > m_TopHeaderHeight + m_ColumnHeaderHeight Then
            If X < RowHW Then
                DayIndex = -1
            Else
                DayIndex = (X - RowHW) \ RowWidth
            End If
        End If
        
        If YY > m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight Then
            SelStart = ((YY - m_TopHeaderHeight - m_ColumnHeaderHeight - mTodayHeight) \ RowHeight)
            StartDay = (X - RowHW) \ RowWidth
            EndDay = -1
            If SelStart > 47 Then SelStart = -1
            SelEnd = 0
            Refresh
        ElseIf Y > m_TopHeaderHeight Then
            SelStart = -1
            SelEnd = 0
            Refresh
        End If
    End If
    
    If eViewMode = vm_Day And Button = vbLeftButton Then
        YY = Y + (VScroll1.Value * RowHeight)
        If YY > m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight Then
            SelStart = ((YY - m_TopHeaderHeight - m_ColumnHeaderHeight - mTodayHeight) \ RowHeight)
            If SelStart > 47 Then SelStart = -1
            SelEnd = 0
            Refresh
        ElseIf Y > m_TopHeaderHeight Then
            SelStart = -1
            SelEnd = 0
            Refresh
        End If
    End If
    
    If eViewMode = vm_Month Then
        If Y > m_TopHeaderHeight + m_ColumnHeaderHeight Then
        
            If Button = vbLeftButton Then
                SelStart = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth)
                SelEnd = SelStart
                Refresh
            End If
        
            Index = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth) + 1
            If mEventsCount > 0 Then
                If Index < 0 Or Index > 7 * m_WeekCounts Then Exit Sub
                If tGrid(Index).HaveHideEvents Then
                    YY = (Y - m_TopHeaderHeight - m_ColumnHeaderHeight) Mod RowHeight
                    If YY > RowHeight - 10 * nScale Then
                        
                        SelStart = -1
                        SelEnd = -1
                        Refresh
                        Exit Sub
                    End If
                End If
            End If
    
            SelStart = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth)
            SelEnd = SelStart
            Exit Sub
        End If
    End If
    
    'Buttons
    If Y < m_TopHeaderHeight Then
        For i = 0 To UBound(tButtons)
            If PtInRect(tButtons(i).RECT, X, Y) Then
                tButtons(i).State = Pressed
                Refresh
                Exit For
            End If
        Next
        
        StartDay = -1
        EndDay = -1
        SelStart = -1
        SelEnd = -1
    End If

    Refresh
End Sub

Private Function GetDTFromMousePos(ByVal X As Single, ByVal Y As Single) As Date
    Dim d As Integer
    Dim FDOM As Date
    Dim nDay As Integer
    Dim HH As Long
    Dim ScrollWidth As Long
    HH = m_TopHeaderHeight + m_ColumnHeaderHeight
    If VScroll1.Visible Then ScrollWidth = VScroll1.Width
    Dim Days As Date
    
    Y = Y - HH
    If Y < 0 Then Y = 0
    If X < 0 Then X = 0
    If Y + HH > UserControl.ScaleHeight Then Y = UserControl.ScaleHeight - HH - EventHeight
    If X > UserControl.ScaleWidth - ScrollWidth Then X = UserControl.ScaleWidth - ScrollWidth - EventHeight

    If eViewMode = vm_Month Then
        FDOM = DateSerial(Year(xDate), Month(xDate), 1)
        d = Weekday(FDOM, m_FirstDayOfWeek)
        nDay = (Y \ RowHeight) * 7 + X \ RowWidth
        GetDTFromMousePos = DateAdd("d", -d + 7 * WeekScroll + 1 + nDay, FDOM)
    ElseIf eViewMode = vm_Week Then
        d = Weekday(xDate, m_FirstDayOfWeek)
        Days = DateAdd("d", -d + (X - RowHW) \ RowWidth + 1, xDate)
        GetDTFromMousePos = DateAdd("n", ((Y - mTodayHeight + (VScroll1.Value * RowHeight)) \ RowHeight) * 30, Days)
    ElseIf eViewMode = vm_Day Then
        GetDTFromMousePos = DateAdd("n", ((Y - mTodayHeight + (VScroll1.Value * RowHeight)) \ RowHeight) * 30, xDate)
    End If
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, j As Long
    Dim Index As Long
    Dim YY As Single
    Dim TempDate As Date
    Dim ST As Date, ET As Date, CellDate As Date
    Dim bCancel As Boolean
    Dim AllDay As Boolean
    
    If isMouseEnter = False Then
        Dim TMET As TRACKMOUSEEVENTTYPE
        TMET.cbSize = Len(TMET)
        TMET.hwndTrack = hwnd
        TMET.dwFlags = TME_LEAVE
        TrackMouseEvent TMET
        isMouseEnter = True
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)

    '///////////RESIZE EVENT/////////////
    If mStartSize = True Then

        For i = 0 To UBound(tEvents)
            If tEvents(i).key = mDragKey Then
                Index = i
                Exit For
            End If
        Next
        
        TempDate = GetDTFromMousePos(X, Y)

        If eViewMode = vm_Month Then
            ST = VBA.DateValue(tEvents(Index).StartTime)
            ET = VBA.DateValue(tEvents(Index).EndTime)
        Else
            ST = tEvents(Index).StartTime
            ET = tEvents(Index).EndTime
        End If
        
        
        
        If mSizeDirection = 0 Then
            If TempDate <= ET Then
                RaiseEvent PreEventChangeDate(tEvents(Index).key, TempDate, ET, bCancel)
            End If
        Else
            If TempDate >= ST Then
                If TimeValue(TempDate) = "00:00:00" Then TempDate = DateAdd("s", -1, TempDate) '23:59:59
                RaiseEvent PreEventChangeDate(tEvents(Index).key, ST, TempDate, bCancel)
            End If
        End If
        
        If bCancel Then
            UserControl.MousePointer = vbNoDrop
            Exit Sub
        Else
            If UserControl.MousePointer = vbNoDrop Then
                If eViewMode = vm_Month Then
                    UserControl.MousePointer = vbSizeWE
                Else
                    UserControl.MousePointer = vbSizeNS
                End If
            End If
        End If

        If mSizeDirection = 0 Then
            If TempDate <= ET Then
                tEvents(Index).StartTime = TempDate
            End If
        Else
            If TempDate >= ST Then
                tEvents(Index).EndTime = TempDate
            End If
        End If

        If tEvents(Index).EndTime = tEvents(Index).StartTime Then
            tEvents(Index).EndTime = DateAdd("n", 15, tEvents(Index).StartTime)
        End If
        
        tEvents(Index).More24Hours = DateDiff("h", ST, ET) > 24
        
        ST = tEvents(Index).StartTime
        ET = tEvents(Index).EndTime
        AllDay = tEvents(Index).AllDayEvent
        
        QSortEvents 0, mEventsCount - 1
        ProcessEvents
        Refresh
        
        RaiseEvent EventChangeDate(mDragKey, ST, ET, AllDay)
        Exit Sub
    End If
    
    '///////////DRAG EVENT/////////////
    If mStartDrag = True Then
        For i = 0 To UBound(tEvents)
            If tEvents(i).key = mDragKey Then
                Index = i
                Exit For
            End If
        Next
        
        ST = DateAdd("n", -mEvDragFromDayNro, GetDTFromMousePos(X, Y))
        ET = DateAdd("n", mEvDragDaysCount, ST)
        If TimeValue(ET) = "00:00:00" Then ET = DateAdd("s", -1, ET) '23:59:59
        RaiseEvent PreEventChangeDate(tEvents(Index).key, ST, ET, bCancel)
        
        If bCancel Then
            UserControl.MousePointer = vbNoDrop
            Exit Sub
        Else
            If UserControl.MousePointer = vbNoDrop Then
                UserControl.MousePointer = vbSizeAll
            End If
        End If
    
        If eViewMode = vm_Week Or eViewMode = vm_Day Then
            If Y < m_TopHeaderHeight + m_ColumnHeaderHeight + mTodayHeight - (VScroll1.Value * RowHeight) Then
                tEvents(Index).AllDayEvent = True
                AllDay = tEvents(Index).AllDayEvent
                RaiseEvent EventChangeDate(mDragKey, ST, ET, AllDay)
                QSortEvents 0, mEventsCount - 1
                ProcessEvents
                Refresh
                Exit Sub
            Else
                tEvents(Index).AllDayEvent = False
                tEvents(Index).More24Hours = False
            End If
        End If

        tEvents(Index).StartTime = ST
        tEvents(Index).EndTime = ET
        AllDay = tEvents(Index).AllDayEvent
        QSortEvents 0, mEventsCount - 1
        ProcessEvents
        Refresh
        
        RaiseEvent EventChangeDate(mDragKey, ST, ET, AllDay)
        Exit Sub
        
   End If
       
    '////////BUTTONS///////////
    If Y < m_TopHeaderHeight Then
        For i = 0 To UBound(tButtons)
            If PtInRect(tButtons(i).RECT, X, Y) Then
                If tButtons(i).State = Normal Then
                    tButtons(i).State = IIf(Button = vbLeftButton, Pressed, Hot)
                    Refresh
                End If
            Else
                If tButtons(i).State <> Normal Then
                    tButtons(i).State = Normal
                    Refresh
                End If
            End If
        Next
        If Button <> vbLeftButton Then Exit Sub
    End If
    
    '//////////////WEEK & DAY//////////////
    If eViewMode = vm_Week Or eViewMode = vm_Day Then
        YY = Y + (VScroll1.Value * RowHeight)
        If Y > m_TopHeaderHeight + m_ColumnHeaderHeight Then
            If Button = vbLeftButton And mSelectedEvent = -1 Then
                
                SelEnd = (YY - m_TopHeaderHeight - m_ColumnHeaderHeight - mTodayHeight) \ RowHeight - SelStart
                EndDay = (X - RowHW) \ RowWidth

                If SelStart + SelEnd > 47 Then SelEnd = 47 - SelStart
               
                
                If Y > UserControl.ScaleHeight Then
                    If VScroll1.Value < VScroll1.max Then
                        VScroll1.Value = VScroll1.Value + 1
                    End If
                Else
                    Refresh
                End If
            End If
        Else
            If bMouseDownInCal And Button = vbLeftButton Then
                If SelEnd > 0 Then SelEnd = SelEnd - 1
                If VScroll1.Value > 0 Then
                    VScroll1.Value = VScroll1.Value - 1
                End If
            End If
        End If
    ElseIf eViewMode = vm_Month Then '///////////MONTH//////////
        If Y > m_TopHeaderHeight + m_ColumnHeaderHeight Then
        
            If Button = vbLeftButton And SelStart > -1 Then
                SelEnd = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth)
                If SelEnd > 7 * m_WeekCounts - 1 Then SelEnd = 7 * m_WeekCounts - 1
                Refresh
            End If
            
            Index = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth) + 1
            If Index < 1 Or Index > 7 * m_WeekCounts Then Exit Sub
            
            If mEventsCount > 0 Then
                If tGrid(Index).HaveHideEvents Then
                    YY = (Y - m_TopHeaderHeight - m_ColumnHeaderHeight) Mod RowHeight
                    If YY > RowHeight - 10 * nScale Then
                        MousePointerHands = True
                        Exit Sub
                    End If
                End If
            End If
            
            YY = Y
        End If
    ElseIf eViewMode = vm_Year Then
        For i = 0 To 11
            If PtInRect(YearCalRects(i), X, Y) Then
                If YearCalIndex <> i Then
                    YearCalIndex = i
                    Refresh
                End If
                Exit Sub
            End If
        Next
        
        If YearCalIndex <> -1 Then
            YearCalIndex = -1
            Refresh
        End If
        
        Exit Sub
    End If

    For i = 0 To mEventsCount - 1
        Dim CopyEvent As CalEvents
        
        CopyEvent = tEvents(i)
        With CopyEvent
        For j = 0 To .RectsCount - 1
            If PtInRect(.Rects(j), X, YY) Then
            
                If .key <> mHotEvent Then
                    mHotEvent = .key
                    
                    '//////TOOLTIPTEXT//////////
                    If m_ShowToolTipEvents Then
                        Dim hIcon As Long
                        TI.lpszText = mStrStarts & ": " & .StartTime & vbCrLf & mStrEnds & ": " & .EndTime
                        'SetWindowPos m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
                        SendMessageW m_hwndTT, TTM_TRACKACTIVATE, False, TI
                        SendMessageLongW m_hwndTT, TTM_UPDATETIPTEXTW, 0, VarPtr(TI)
                        
                        hIcon = GetEventIcon(i)
                        ' Se desactivaron estas 2 funciones que siguen para evitar los mensajes popup
                        ' SendMessageLongW m_hwndTT, TTM_SETTITLEW, hIcon, StrPtr(.Subject)
                        ' SendMessageW m_hwndTT, TTM_TRACKACTIVATE, True, TI
                        DestroyIcon hIcon
                    End If
                    
                    RaiseEvent EventMouseEnter(mHotEvent)
                End If

                If mStartDrag = False Then 'And mSelectedEvent <> -1 THEN
                    If Button = vbLeftButton And (Abs(PointMdown.X - X) > 2 * nScale Or Abs(PointMdown.Y - Y) > 2 * nScale) Then
                        If mSelectedEvent <> -1 Then
                            CellDate = GetDTFromMousePos(X, Y)
                            mEvDragFromDayNro = DateDiff("n", .StartTime, CellDate)
                            mStartDrag = True
                            MousePointerHands = False
                            Me.MousePointer = vbSizeAll
                            
                            mEvDragDaysCount = DateDiff("n", .StartTime, .EndTime)
                            If GetKeyState(vbKeyControl) < 0 Then

                                mDragKey = Me.AddEvents(.Subject, .StartTime, .EndTime, _
                                            .ForeColor, .AllDayEvent, .body, .Tag, _
                                            .IsSerie, .NotifyIcon, .IsPrivate, .EventShowAs)
                                RaiseEvent DragNewEvent(mDragKey, .key)
                            Else
                                mDragKey = .key
                            End If
                        End If
                    Else
                       
                        '///SIZE///
                        If eViewMode = vm_Month Then
                            If X - .Rects(j).Left <= MarginText Then
                                MousePointer = vbSizeWE
                            ElseIf .Rects(j).Right - X <= MarginText Then
                                MousePointer = vbSizeWE
                            Else
                                If MousePointer = vbSizeWE Then MousePointer = vbDefault
                                MousePointerHands = True
                            End If
                        Else
                            If YY - .Rects(j).top <= MarginText Then
                                MousePointer = vbSizeNS
                            ElseIf .Rects(j).Bottom - YY <= MarginText Then
                                MousePointer = vbSizeNS
                            Else
                                If MousePointer = vbSizeNS Then MousePointer = vbDefault
                                MousePointerHands = True
                            End If
                        End If
                    End If
                End If
                Exit Sub
            End If
        Next
        End With
    Next

    If mHotEvent <> -1 Then
        RaiseEvent EventMouseLeave(mHotEvent)
        If m_hwndTT Then SendMessageW m_hwndTT, TTM_TRACKACTIVATE, False, TI
        mHotEvent = -1
    End If

    If hCur Then MousePointerHands = False
    If Me.MousePointer <> vbDefault Then
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim i As Long
    Dim Rects() As RECT
    Dim Index As Integer, YY As Long
    Dim NewDate As Date
    Dim bCancel As Boolean
   
    If mStartSize = True Then
        mStartSize = False
        Exit Sub
    End If
    
    If mStartDrag = True Then
        mStartDrag = False
        MousePointerHands = False
        Me.MousePointer = vbDefault
        Refresh
        Exit Sub
    End If
      
    RaiseEvent MouseUp(Button, Shift, X, Y)

    If eViewMode = vm_Year And YearCalIndex <> -1 Then
        xDate = DateSerial(Year(xDate), YearCalIndex + 1, 1)
        Me.ViewMode = vm_Month
        Exit Sub
    End If

    If Y > m_TopHeaderHeight + m_ColumnHeaderHeight Then
        If mSelectedEvent > -1 Then
            Index = GetEventIndexByKey(mSelectedEvent)
            Rects = tEvents(Index).Rects
            If tEvents(Index).RectsCount > 0 Then
                For i = 0 To UBound(Rects)
                    If PtInRect(Rects(i), X, Y + (VScroll1.Value * RowHeight)) Then
                        RaiseEvent EVENTCLICK(mSelectedEvent, Button)
                        Exit For
                    End If
                Next
               
            End If
        End If
    End If
    
    'CLICK EN EL DROPDOWN DE LOS MESES
    If Y > m_TopHeaderHeight + m_ColumnHeaderHeight And eViewMode = vm_Month Then
        Index = ((Y - m_TopHeaderHeight - m_ColumnHeaderHeight) \ RowHeight) * 7 + (X \ RowWidth) + 1
        If Index < 0 Or Index > 7 * m_WeekCounts Then Exit Sub
        If mEventsCount > 0 Then
            If tGrid(Index).HaveHideEvents Then
                
                YY = (Y - m_TopHeaderHeight - m_ColumnHeaderHeight) Mod RowHeight
                If YY > RowHeight - 10 * nScale Then
                    Dim FDOM As Date, d As Integer, Days As Date
                
                    FDOM = DateSerial(Year(xDate), Month(xDate), 1)
                    d = Weekday(FDOM, m_FirstDayOfWeek) - VScroll1.Value * 7
                    Days = DateAdd("d", -d + Index, FDOM)
                    
                    RaiseEvent DropDownViewMore(Days, bCancel)
                    If Not bCancel Then
                        xDate = Days
                        eViewMode = vm_Day
                        VScroll1.Visible = True
                        ProcessEvents
                        Refresh
                    End If
                    Exit Sub
                End If
            
            End If
        End If
    End If
 
    For i = 0 To UBound(tButtons)
        If PtInRect(tButtons(i).RECT, X, Y) Then
            If tButtons(i).State = Pressed Then
                Select Case i
                    Case 0
                        If m_UserCanChangeDate = False Then Exit Sub
                        Select Case eViewMode
                            Case vm_Day: NewDate = DateAdd("d", -1, xDate)
                            Case vm_Week: NewDate = DateAdd("d", -7, xDate)
                            Case vm_Month: NewDate = DateAdd("m", -1, xDate)
                            Case vm_Year: NewDate = DateAdd("YYYY", -1, xDate)
                        End Select
                        RaiseEvent PreDateChange(NewDate, bCancel)
                        If Not bCancel Then
                            xDate = NewDate
                            WeekScroll = 0
                            ProcessEvents
                        End If
                        RaiseEvent DateChange(NewDate)
                    Case 1
                        If m_UserCanChangeDate = False Then Exit Sub
                        Select Case eViewMode
                            Case vm_Day: NewDate = DateAdd("d", 1, xDate)
                            Case vm_Week: NewDate = DateAdd("d", 7, xDate)
                            Case vm_Month: NewDate = DateAdd("m", 1, xDate)
                            Case vm_Year: NewDate = DateAdd("YYYY", 1, xDate)
                        End Select
                        RaiseEvent PreDateChange(NewDate, bCancel)
                        If Not bCancel Then
                            xDate = NewDate
                            WeekScroll = 0
                            ProcessEvents
                        End If
                        RaiseEvent DateChange(NewDate)
                    Case 2
                        If m_UserCanChangeViewMode = False Then Exit Sub
                        eViewMode = vm_Year
                        VScroll1.Visible = False
                        VScroll1.Value = 0
                    Case 3
                        If m_UserCanChangeViewMode = False Then Exit Sub
                        eViewMode = vm_Month
                        VScroll1.Visible = False
                        VScroll1.Value = 0
                        ProcessEvents
                    Case 4
                        If m_UserCanChangeViewMode = False Then Exit Sub
                        eViewMode = vm_Week
                        
                        VScroll1.Visible = True
                        ProcessEvents
                        CenterCalenarInNow
                    Case 5
                        If m_UserCanChangeViewMode = False Then Exit Sub
                        eViewMode = vm_Day
                        
                        VScroll1.Visible = True
                        ProcessEvents
                        CenterCalenarInNow
                    Case 6
                        If m_UserCanChangeViewMode = False Then Exit Sub
                        xDate = Date
                        ProcessEvents
                        If eViewMode = vm_Day Or eViewMode = vm_Week Then CenterCalenarInNow
                        WeekScroll = 0
                End Select
                tButtons(i).State = Hot
                StartDay = -1
                EndDay = -1
                SelStart = -1
                SelEnd = -1
                Refresh
            End If
            
            Exit For
        End If
    Next
End Sub

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(ByVal NewValue As IPictureDisp)
    UserControl.MouseIcon = NewValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Private Property Let MousePointerHands(ByVal NewValue As Boolean)
On Error Resume Next
    GoTo salida

    If m_MousePointerHands = True And NewValue = True Then
        SetCursor hCur
        Exit Property
    End If

    m_MousePointerHands = NewValue
    If NewValue Then
        UserControl.MousePointer = vbCustom
        UserControl.MouseIcon = GetSystemHandCursor
        SetCursor hCur
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        UserControl.MousePointer = vbDefault
        UserControl.MouseIcon = Nothing
    End If
    
salida:

End Property

Private Property Get MousePointerHands() As Boolean
    MousePointerHands = m_MousePointerHands
End Property

Private Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, IPic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980: GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B: GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
 
    Set GetSystemHandCursor = IPic
End Function

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function PvFirstDayOfWeek(xDate As Date) As Date
  Dim d As Integer
  d = Weekday(xDate, m_FirstDayOfWeek)
  PvFirstDayOfWeek = DateAdd("d", -d + 1, xDate)
End Function

Private Sub VScroll1_Change()
    If VScroll1.Visible Then Refresh
End Sub

Private Sub VScroll1_Scroll()
    Refresh
End Sub

Private Function GetEventIndexByKey(key As Long) As Long
    Dim i As Long
    For i = 0 To mEventsCount - 1
        If tEvents(i).key = key Then
            GetEventIndexByKey = i
            Exit Function
        End If
    Next
    GetEventIndexByKey = -1
End Function

Private Function GetEventIcon(Index As Long)
    Dim Width As Long, Height As Long
    Dim hImage As Long, hGraphics As Long
    Width = 16
    Height = 16
    GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    With tEvents(Index)
        Select Case .EventShowAs
            Case ESA_Busy
                FillRectangleEx hGraphics, RGBtoARGB(.ForeColor, 100), RGBtoARGB(.ForeColor, 100), 0, 0, 0, Width, Height
            Case ESA_Free
                FillRectangleEx hGraphics, RGBtoARGB(UserControl.BackColor, 100), RGBtoARGB(UserControl.BackColor, 100), 0, 0, 0, Width, Height
            Case [ESA_Out of office]
                FillRectangleEx hGraphics, RGBtoARGB(&H800080, 100), RGBtoARGB(&H800080, 100), 0, 0, 0, Width, Height
            Case [ESA_Working elsewhere]
                FillRectangleEx hGraphics, RGBtoARGB(.ForeColor, 100), RGBtoARGB(UserControl.BackColor, 100), 14, 0, 0, Width, Height
            Case ESA_Tentative
                FillRectangleEx hGraphics, RGBtoARGB(.ForeColor, 100), RGBtoARGB(UserControl.BackColor, 100), 23, 0, 0, Width, Height
        End Select
    End With
    GdipCreateHICONFromBitmap hImage, GetEventIcon

    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
End Function



