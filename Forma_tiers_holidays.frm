VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiers setup"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   72
      Top             =   5880
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   71
      Top             =   4920
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   70
      Top             =   3840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   69
      Top             =   2880
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   68
      Top             =   1800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkactivar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   67
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   6120
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   4080
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorato 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   6120
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   4080
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox cbohorafrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   1080
      Width           =   975
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3255
      Left            =   11040
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      _Version        =   524288
      _ExtentX        =   5953
      _ExtentY        =   5741
      _StockProps     =   1
      BackColor       =   12632256
      Year            =   2021
      Month           =   3
      Day             =   26
      DayLength       =   0
      MonthLength     =   2
      DayFontColor    =   16711680
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   16711680
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   25
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   24
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   22
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtannualGI 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin Project1.lvButtons_H CmdCancel 
      Height          =   615
      Left            =   13680
      TabIndex        =   0
      Top             =   6360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_tiers_holidays.frx":0000
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   13200
      TabIndex        =   40
      Top             =   1800
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   8421504
      ForeColorFixed  =   14737632
      BackColorBkg    =   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btngrabar 
      Height          =   615
      Left            =   12480
      TabIndex        =   42
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "Save tiers"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_tiers_holidays.frx":11E3
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   5
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   8640
      TabIndex        =   34
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   8640
      TabIndex        =   35
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   8640
      TabIndex        =   39
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   8640
      TabIndex        =   38
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   36
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lbldateto 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   8640
      TabIndex        =   37
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   4575
      Left            =   10200
      Picture         =   "Forma_tiers_holidays.frx":1635
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   4
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   3
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   2
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   1
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000FF&
      Height          =   975
      Index           =   0
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   5
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   4
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   3
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   2
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   1
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape marco1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Height          =   975
      Index           =   0
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   6240
      TabIndex        =   33
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   32
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   31
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   30
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   29
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lbldatefrom 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6240
      TabIndex        =   28
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From                                                To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   6960
      TabIndex        =   27
      Top             =   240
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Access"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   7920
      TabIndex        =   26
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Annual GI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   4920
      TabIndex        =   19
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   11
      Left            =   3960
      TabIndex        =   18
      Top             =   5880
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   10
      Left            =   3960
      TabIndex        =   17
      Top             =   3840
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   9
      Left            =   3960
      TabIndex        =   16
      Top             =   1800
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   3960
      TabIndex        =   15
      Top             =   4920
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   3960
      TabIndex        =   14
      Top             =   2880
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   3960
      TabIndex        =   13
      Top             =   840
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   3720
      TabIndex        =   12
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Top             =   5880
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   4
      Left            =   2520
      TabIndex        =   10
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   4920
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   645
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   5280
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   645
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   3240
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   645
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   120
      Picture         =   "Forma_tiers_holidays.frx":2375A
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   120
      Picture         =   "Forma_tiers_holidays.frx":4258A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "Forma_tiers_holidays.frx":60B26
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   345
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   20
      Index           =   1
      Left            =   1320
      Top             =   4560
      Width           =   11295
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   15
      Index           =   0
      Left            =   1320
      Top             =   2520
      Width           =   11295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fila As Integer, columna As Integer

Private Sub btngrabar_Click()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset


R$ = MsgBox("Do you want to save all the tiers?", 4, "Attention")
If R$ = "7" Then Exit Sub

Dim datefrom$(6), dateto$(6)

For t = 0 To 5
  datefrom$(t) = lbldatefrom(t).Caption + " " + cbohorafrom(t).List(cbohorafrom(t).ListIndex)
  dateto$(t) = lbldateto(t).Caption + " " + cbohorafrom2(t).List(cbohorafrom2(t).ListIndex)
Next t


sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(0) + "', DateAccessTo='" + dateto$(0) & _
"', AnnualGI='" + txtannualGI(0).Text + "', active='" + Format(chkactivar(0).Value, "0") + "' where IdVacationTier=1 and priority=1"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(1) + "', DateAccessTo='" + dateto$(1) & _
"', AnnualGI='" + txtannualGI(1).Text + "', active='" + Format(chkactivar(1).Value, "0") + "' where IdVacationTier=1 and priority=2"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(2) + "', DateAccessTo='" + dateto$(2) & _
"', AnnualGI='" + txtannualGI(2).Text + "', active='" + Format(chkactivar(2).Value, "0") + "' where IdVacationTier=2 and priority=1"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(3) + "', DateAccessTo='" + dateto$(3) & _
"', AnnualGI='" + txtannualGI(3).Text + "', active='" + Format(chkactivar(3).Value, "0") + "' where IdVacationTier=2 and priority=2"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(4) + "', DateAccessTo='" + dateto$(4) & _
"', AnnualGI='" + txtannualGI(4).Text + "', active='" + Format(chkactivar(4).Value, "0") + "' where IdVacationTier=3 and priority=1"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

sSelect = "update VacationsTiersCatalog set DateAccessFrom='" + datefrom$(5) + "', DateAccessTo='" + dateto$(5) & _
"', AnnualGI='" + txtannualGI(5).Text + "', active='" + Format(chkactivar(5).Value, "0") + "' where IdVacationTier=3 and priority=2"
Rs.Open sSelect, base, adOpenUnspecified
Rs.Close

Unload Me

End Sub

Private Sub btnok_Click()

End Sub

Private Sub calendar1_Click()
On Error Resume Next

If fila = 0 Then

  lbldatefrom(columna).Caption = Calendar1.Value

Else

  lbldateto(columna).Caption = Calendar1.Value


End If


Calendar1.Visible = False
For t = 0 To 5
  marco1(t).Visible = False
   marco2(t).Visible = False
Next t
End Sub

Private Sub cbohorafrom_Click(Index As Integer)
On Error Resume Next


  cbohorato(Index).Clear
  
  num = cbohorafrom(Index).ListIndex + 1
  
  
  For t = num To 20
    cbohorato(Index).AddItem cbohorafrom(Index).List(t)
  Next t
End Sub

Private Sub cbohorafrom2_Click(Index As Integer)
On Error Resume Next


  cbohorato2(Index).Clear
  
  num = cbohorafrom2(Index).ListIndex + 1
  
  
  For t = num To 20
    cbohorato2(Index).AddItem cbohorafrom2(Index).List(t)
  Next t
End Sub

Private Sub CmdCancel_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Form_Click()
Calendar1.Visible = False
For t = 0 To 5
  marco1(t).Visible = False
   marco2(t).Visible = False
Next t
End Sub

Private Sub Form_Load()
On Error Resume Next
Left = (Screen.Width - Width) - 2500
top = 0




For t = 0 To 5
  
  cbohorafrom(t).AddItem "09:00"
  cbohorafrom(t).AddItem "09:30"
  cbohorafrom(t).AddItem "10:00"
  cbohorafrom(t).AddItem "10:30"
  cbohorafrom(t).AddItem "11:00"
  cbohorafrom(t).AddItem "11:30"
  cbohorafrom(t).AddItem "12:00"
  cbohorafrom(t).AddItem "12:30"
  cbohorafrom(t).AddItem "13:00"
  cbohorafrom(t).AddItem "13:30"
  cbohorafrom(t).AddItem "14:00"
  cbohorafrom(t).AddItem "14:30"
  cbohorafrom(t).AddItem "15:00"
  cbohorafrom(t).AddItem "15:30"
  cbohorafrom(t).AddItem "16:00"
  cbohorafrom(t).AddItem "16:30"
  cbohorafrom(t).AddItem "17:00"
  cbohorafrom(t).AddItem "17:30"
  cbohorafrom(t).AddItem "18:00"
  cbohorafrom(t).AddItem "18:30"
  cbohorafrom(t).AddItem "19:00"
  
Next t







For t = 0 To 5
  
  cbohorafrom2(t).AddItem "09:00"
  cbohorafrom2(t).AddItem "09:30"
  cbohorafrom2(t).AddItem "10:00"
  cbohorafrom2(t).AddItem "10:30"
  cbohorafrom2(t).AddItem "11:00"
  cbohorafrom2(t).AddItem "11:30"
  cbohorafrom2(t).AddItem "12:00"
  cbohorafrom2(t).AddItem "12:30"
  cbohorafrom2(t).AddItem "13:00"
  cbohorafrom2(t).AddItem "13:30"
  cbohorafrom2(t).AddItem "14:00"
  cbohorafrom2(t).AddItem "14:30"
  cbohorafrom2(t).AddItem "15:00"
  cbohorafrom2(t).AddItem "15:30"
  cbohorafrom2(t).AddItem "16:00"
  cbohorafrom2(t).AddItem "16:30"
  cbohorafrom2(t).AddItem "17:00"
  cbohorafrom2(t).AddItem "17:30"
  cbohorafrom2(t).AddItem "18:00"
  cbohorafrom2(t).AddItem "18:30"
  cbohorafrom2(t).AddItem "19:00"
  
Next t


carga_tiers
Calendar1.Visible = False



End Sub

Public Sub carga_tiers()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

sSelect = "select * from VacationsTiersCatalog"  ' where active=1"
    Dim activo As Boolean
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
   Rs.Close
   
   
   ' carga la info
   For t = 1 To Grid2.Rows - 1
     Grid2.Row = t
     Grid2.Col = 1
     id_vacation_tier$ = Grid2.Text
     
     Grid2.Col = 2
     Priority$ = Grid2.Text
     
     Grid2.Col = 3
     Id_job_title$ = Grid2.Text
     
     Grid2.Col = 4
     Annual_GI$ = Grid2.Text
     
     Grid2.Col = 5
     Fecha_from$ = Format(Grid2.Text, "mm/dd/yyyy")
     hora_from$ = Format(Grid2.Text, "hh:mm")
     
     Grid2.Col = 6
     Fecha_to$ = Format(Grid2.Text, "mm/dd/yyyy")
     hora_to$ = Format(Grid2.Text, "hh:mm")
     
     Grid2.Col = 7
     activo = Grid2.Text
     If activo = True Then
       chkactivar(t - 1).Value = "1"
     Else
       chkactivar(t - 1).Value = "0"
     End If
     
     If Val(id_vacation_tier$) = 1 Then
           If Val(Priority$) = 1 And Val(Id_job_title$) = 17 Then
              txtannualGI(0).Text = Annual_GI$
              lbldatefrom(0).Caption = Fecha_from$
              lbldateto(0).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(0).ListCount - 1
                 If cbohorafrom(0).List(Y) = hora_from$ Then
                   cbohorafrom(0).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(0).ListCount - 1
                 If cbohorafrom2(0).List(Y) = hora_to$ Then
                   cbohorafrom2(0).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
           ElseIf Val(Priority$) = 2 And Val(Id_job_title$) = 16 Then
              txtannualGI(1).Text = Annual_GI$
              lbldatefrom(1).Caption = Fecha_from$
              lbldateto(1).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(1).ListCount - 1
                 If cbohorafrom(1).List(Y) = hora_from$ Then
                   cbohorafrom(1).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(1).ListCount - 1
                 If cbohorafrom2(1).List(Y) = hora_to$ Then
                   cbohorafrom2(1).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              
           End If
     ElseIf Val(id_vacation_tier$) = 2 Then
           If Val(Priority$) = 1 And Val(Id_job_title$) = 17 Then
              txtannualGI(2).Text = Annual_GI$
              lbldatefrom(2).Caption = Fecha_from$
              lbldateto(2).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(2).ListCount - 1
                 If cbohorafrom(2).List(Y) = hora_from$ Then
                   cbohorafrom(2).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(2).ListCount - 1
                 If cbohorafrom2(2).List(Y) = hora_to$ Then
                   cbohorafrom2(2).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
           ElseIf Val(Priority$) = 2 And Val(Id_job_title$) = 16 Then
              txtannualGI(3).Text = Annual_GI$
              lbldatefrom(3).Caption = Fecha_from$
              lbldateto(3).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(3).ListCount - 1
                 If cbohorafrom(3).List(Y) = hora_from$ Then
                   cbohorafrom(3).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(3).ListCount - 1
                 If cbohorafrom2(3).List(Y) = hora_to$ Then
                   cbohorafrom2(3).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
           End If
     ElseIf Val(id_vacation_tier$) = 3 Then
           If Val(Priority$) = 1 And Val(Id_job_title$) = 17 Then
              txtannualGI(4).Text = Annual_GI$
              lbldatefrom(4).Caption = Fecha_from$
              lbldateto(4).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(4).ListCount - 1
                 If cbohorafrom(4).List(Y) = hora_from$ Then
                   cbohorafrom(4).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(4).ListCount - 1
                 If cbohorafrom2(4).List(Y) = hora_to$ Then
                   cbohorafrom2(4).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
           ElseIf Val(Priority$) = 2 And Val(Id_job_title$) = 16 Then
              txtannualGI(5).Text = Annual_GI$
              lbldatefrom(5).Caption = Fecha_from$
              lbldateto(5).Caption = Fecha_to$
              
              For Y = 0 To cbohorafrom(5).ListCount - 1
                 If cbohorafrom(5).List(Y) = hora_from$ Then
                   cbohorafrom(5).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
              For Y = 0 To cbohorafrom2(5).ListCount - 1
                 If cbohorafrom2(5).List(Y) = hora_to$ Then
                   cbohorafrom2(5).ListIndex = Y
                   Exit For
                 End If
              Next Y
              
           End If
     End If
     
     
   Next t
     
   


End Sub

Private Sub lbldatefrom_Click(Index As Integer)
On Error Resume Next
fila = 0
columna = Index

For t = 0 To 5
   marco1(t).Visible = False
   marco2(t).Visible = False
Next t
marco1(Index).Visible = True


If lbldatefrom(Index) = "" Then
   Calendar1.Value = Format(Now, "mm/dd/yyyy")
Else
   Calendar1.Value = lbldatefrom(Index).Caption
End If
Calendar1.Visible = True

End Sub

Private Sub lbldateto_Click(Index As Integer)
On Error Resume Next
fila = 1
columna = Index
For t = 0 To 5
   marco1(t).Visible = False
   marco2(t).Visible = False
Next t
marco2(Index).Visible = True

If lbldateto(Index) = "" Then
   Calendar1.Today
Else
   Calendar1.Value = lbldateto(Index).Caption
End If
Calendar1.Visible = True
End Sub


Private Sub txtannualGI_Click(Index As Integer)
Calendar1.Visible = False
For t = 0 To 5
  marco1(t).Visible = False
   marco2(t).Visible = False
Next t
End Sub

Private Sub txtannualGI_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If

Calendar1.Visible = False


End Sub


