VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "My Holidays"
   ClientHeight    =   13545
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   21540
   Icon            =   "Form2_holidays.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   13545
   ScaleWidth      =   21540
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   13440
      ScaleHeight     =   2295
      ScaleWidth      =   1695
      TabIndex        =   99
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid5 
         Height          =   1815
         Left            =   120
         TabIndex        =   100
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         ForeColorFixed  =   14737632
         BackColorBkg    =   -2147483626
         BorderStyle     =   0
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
      Begin VB.Label lblagente 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000A&
      Height          =   3735
      Left            =   19440
      ScaleHeight     =   3735
      ScaleWidth      =   2055
      TabIndex        =   67
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      Begin Project1.lvButtons_H btnacceso_empleados 
         Height          =   375
         Left            =   480
         TabIndex        =   97
         Top             =   825
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Modify list"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox lista_acceso 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   68
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   3615
         Left            =   0
         Picture         =   "Form2_holidays.frx":3336E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.ListBox lista 
      Height          =   1620
      Left            =   21000
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   6015
   End
   Begin Project1.lvButtons_H btn_configurar_accesos 
      Height          =   495
      Left            =   17640
      TabIndex        =   63
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Tiers setup "
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
      cBack           =   12632256
   End
   Begin VB.Frame marco_anos 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   17640
      TabIndex        =   64
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton op_ano 
         BackColor       =   &H80000016&
         Caption         =   "2024"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   66
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton op_ano 
         BackColor       =   &H80000016&
         Caption         =   "2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.ListBox Lista_membresia 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Index           =   2
      Left            =   17640
      TabIndex        =   59
      Top             =   11880
      Width           =   3735
   End
   Begin VB.ListBox Lista_membresia 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Index           =   1
      Left            =   17640
      TabIndex        =   58
      Top             =   10200
      Width           =   3735
   End
   Begin VB.ListBox Lista_membresia 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Index           =   0
      Left            =   17640
      TabIndex        =   57
      Top             =   8520
      Width           =   3735
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   20880
      TabIndex        =   56
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   20520
      TabIndex        =   55
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   8160
      ScaleHeight     =   675
      ScaleWidth      =   2340
      TabIndex        =   15
      Top             =   1080
      Width           =   2340
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   15480
      ScaleHeight     =   2895
      ScaleWidth      =   5895
      TabIndex        =   50
      Top             =   4320
      Visible         =   0   'False
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid grid4 
         Height          =   2175
         Left            =   0
         TabIndex        =   52
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   10
         BackColor       =   12640511
         BackColorSel    =   16761024
         BackColorBkg    =   -2147483626
         GridColor       =   4210752
         GridColorFixed  =   12632256
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This month's anniversaries"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   0
         Width           =   3345
      End
   End
   Begin Project1.lvButtons_H btnadd_event 
      Height          =   495
      Left            =   12120
      TabIndex        =   1
      Top             =   1545
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Request vacation day"
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
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   11760
      ScaleHeight     =   2595
      ScaleWidth      =   3345
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   3410
      Begin Project1.lvButtons_H btn_vacaciones_info 
         Height          =   495
         Left            =   2760
         TabIndex        =   98
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Caption         =   "Show days"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "A"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.CheckBox chk_permitir_en_2_meses_seguidos 
         Caption         =   "Vacation in 2 consecutive months"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CheckBox chk_dia_doble 
         Caption         =   "2 people in the same day"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1980
         Width           =   2415
      End
      Begin VB.ComboBox cbo_users 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   320
         Width           =   2775
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   71
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "B"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   72
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "C"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   73
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "D"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   74
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "E"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   75
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "F"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   76
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "G"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   77
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "H"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   78
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "I"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   79
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "J"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   80
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "K"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   81
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "L"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   82
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "M"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   83
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "N"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   84
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "O"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   15
         Left            =   600
         TabIndex        =   85
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "P"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   16
         Left            =   840
         TabIndex        =   86
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "Q"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   87
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "R"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   18
         Left            =   1320
         TabIndex        =   88
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "S"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   89
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "T"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   20
         Left            =   1800
         TabIndex        =   90
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "U"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   91
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "V"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   22
         Left            =   2280
         TabIndex        =   92
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "W"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   23
         Left            =   2520
         TabIndex        =   93
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "X"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   24
         Left            =   2760
         TabIndex        =   94
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "Y"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   25
         Left            =   3000
         TabIndex        =   95
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "Z"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H btnletra 
         Height          =   255
         Index           =   26
         Left            =   2760
         TabIndex        =   96
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Caption         =   "A-Z"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   2
         Value           =   -1  'True
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         Caption         =   "Employee's list"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   150
         TabIndex        =   47
         Top             =   80
         Width           =   1335
      End
   End
   Begin Project1.lvButtons_H btn_menu 
      Height          =   615
      Index           =   0
      Left            =   9600
      TabIndex        =   42
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Vacation"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   0
      CapStyle        =   1
      Mode            =   2
      Value           =   -1  'True
      ImgAlign        =   4
      ImgSize         =   40
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btn_hoy 
      Height          =   495
      Left            =   2400
      TabIndex        =   41
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "Today"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   8421504
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "Form2_holidays.frx":36DC4
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.PictureBox mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   3960
      ScaleHeight     =   1710
      ScaleWidth      =   5130
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   5160
      Begin ComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   0
         TabIndex        =   54
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Image Image6 
         Height          =   1680
         Left            =   0
         Picture         =   "Form2_holidays.frx":37A53
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "When sending an email"
      Height          =   1095
      Left            =   3120
      TabIndex        =   27
      Top             =   11280
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000016&
         Caption         =   "Send with High Importance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Request a Reading receipt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox cboimpre 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   11520
      Width           =   3135
   End
   Begin Project1.lvButtons_H btnprint 
      Height          =   615
      Left            =   6840
      TabIndex        =   23
      Top             =   11400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Caption         =   "Print Calendar"
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   13920
      Top             =   5640
   End
   Begin Project1.lvButtons_H btnbloquear_dia 
      Height          =   495
      Left            =   12000
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Block day"
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
      cBack           =   8421631
   End
   Begin Project1.lvButtons_H btnsend 
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   12000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Send a test email"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14400
      Top             =   5640
   End
   Begin VB.ListBox lista2 
      Height          =   1425
      Left            =   21000
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin Project1.lvButtons_H btnconfig_mail 
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   11400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Config email"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Image           =   "Form2_holidays.frx":3BC6D
      ImgSize         =   32
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H btnfin 
      Height          =   735
      Left            =   20400
      TabIndex        =   0
      Top             =   13080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "End"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   21000
      TabIndex        =   12
      Top             =   9360
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
      Height          =   1695
      Left            =   20280
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5640
      Left            =   25000
      TabIndex        =   26
      Top             =   1560
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   9948
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin Project1.ucCalendar ucCalendar1 
      Height          =   9735
      Left            =   600
      TabIndex        =   30
      Top             =   1080
      Width           =   10980
      _ExtentX        =   23098
      _ExtentY        =   18230
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FirstDayOfWeek  =   1
   End
   Begin Project1.lvButtons_H btn_menu 
      Height          =   615
      Index           =   1
      Left            =   8040
      TabIndex        =   43
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Evaluation"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   2
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   40
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2175
      Left            =   21000
      TabIndex        =   49
      Top             =   5520
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   10
      Cols            =   22
      BackColor       =   -2147483636
      BackColorSel    =   16761024
      BackColorBkg    =   -2147483633
      GridColor       =   4210752
      GridColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   16800
      Picture         =   "Form2_holidays.frx":3C391
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image Image5 
      Height          =   2955
      Left            =   4920
      Picture         =   "Form2_holidays.frx":4F7DF
      Stretch         =   -1  'True
      Top             =   11040
      Width           =   2955
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   360
      Top             =   13440
      Width           =   6735
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   60
      Index           =   0
      Left            =   15480
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label lblrango 
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   2
      Left            =   17760
      TabIndex        =   62
      Top             =   11640
      Width           =   3015
   End
   Begin VB.Label lblrango 
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   17760
      TabIndex        =   61
      Top             =   9960
      Width           =   2895
   End
   Begin VB.Label lblrango 
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   17760
      TabIndex        =   60
      Top             =   8280
      Width           =   2895
   End
   Begin VB.Image Image11 
      Height          =   1815
      Left            =   16680
      Picture         =   "Form2_holidays.frx":19CFC8
      Stretch         =   -1  'True
      Top             =   11160
      Width           =   1095
   End
   Begin VB.Image Image10 
      Height          =   1815
      Left            =   16680
      Picture         =   "Form2_holidays.frx":19F071
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   1815
      Left            =   16680
      Picture         =   "Form2_holidays.frx":1A1012
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000C&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   16680
      Shape           =   4  'Rounded Rectangle
      Top             =   11520
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000C&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   16680
      Shape           =   4  'Rounded Rectangle
      Top             =   9840
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000C&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   16680
      Shape           =   4  'Rounded Rectangle
      Top             =   8160
      Width           =   4815
   End
   Begin VB.Image Image8 
      Height          =   600
      Left            =   18960
      Picture         =   "Form2_holidays.frx":1A302C
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   585
   End
   Begin VB.Label lblhired_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm/dd/yyyy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   17640
      TabIndex        =   45
      Top             =   3555
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hired date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Index           =   5
      Left            =   15960
      TabIndex        =   44
      Top             =   3600
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   4575
      Left            =   11400
      Picture         =   "Form2_holidays.frx":1A4AEA
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   5055
   End
   Begin VB.Image calendario 
      Height          =   1455
      Left            =   2220
      Picture         =   "Form2_holidays.frx":1B35C7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   10140
      Left            =   540
      Top             =   795
      Width           =   11115
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taken"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   12840
      TabIndex        =   40
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approved"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   13200
      TabIndex        =   39
      Top             =   9840
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pending"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   12480
      TabIndex        =   38
      Top             =   9480
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approved"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   13200
      TabIndex        =   37
      Top             =   9120
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day "
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
      Index           =   4
      Left            =   12480
      TabIndex        =   36
      Top             =   10320
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Out"
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
      Index           =   3
      Left            =   12480
      TabIndex        =   35
      Top             =   9960
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "approval"
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
      Index           =   2
      Left            =   13560
      TabIndex        =   34
      Top             =   9600
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Half Day "
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
      Index           =   1
      Left            =   12480
      TabIndex        =   33
      Top             =   9240
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approved"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   13095
      TabIndex        =   32
      Top             =   8760
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Day "
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
      Index           =   0
      Left            =   12480
      TabIndex        =   31
      Top             =   8880
      Width           =   615
   End
   Begin VB.Shape Circulo 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   10200
      Width           =   255
   End
   Begin VB.Shape Circulo 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   9840
      Width           =   255
   End
   Begin VB.Shape Circulo 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   9480
      Width           =   255
   End
   Begin VB.Shape Circulo 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   8760
      Width           =   255
   End
   Begin VB.Shape Circulo 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   12120
      Shape           =   3  'Circle
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the printer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   25
      Top             =   11280
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form2_holidays.frx":1B40DE
      Stretch         =   -1  'True
      Top             =   11040
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   15120
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lblmanager_ID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   19560
      TabIndex        =   14
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   19560
      TabIndex        =   13
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Hector Navarro   "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   11
      Top             =   13200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2023-2024"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   12960
      Width           =   2175
   End
   Begin VB.Label lblemp_id 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   19560
      TabIndex        =   9
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   19560
      TabIndex        =   8
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label lbllocation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   15360
      TabIndex        =   7
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   15360
      TabIndex        =   6
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label lblmanager 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   15360
      TabIndex        =   5
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   15360
      TabIndex        =   4
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   15360
      TabIndex        =   3
      Top             =   360
      Width           =   900
   End
   Begin VB.Label lblemployee 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   15360
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   4455
      Left            =   11040
      Picture         =   "Form2_holidays.frx":1B6358
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   4695
   End
   Begin VB.Image Image12 
      Height          =   7215
      Left            =   0
      Picture         =   "Form2_holidays.frx":1C5D5C
      Stretch         =   -1  'True
      Top             =   620
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   3255
      Left            =   17800
      Picture         =   "Form2_holidays.frx":1CA171
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   16560
      Top             =   2880
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim empleado_en_esta_fecha(5) As Integer, bloquea_acceso As Integer, seg As Integer, total_registros As Integer
Dim fecha1$, fecha2$, tipo_ejercicio As Integer
Dim carga_inicial As Integer, ano_para_GI As Integer, permiso_doble_mes As Integer, letra$


Private mcIni                       As clsIni

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)



Const lista_gris = &HEFEFEF
Const lista_naranja = &HCEEFFD
Const lista_roja = &HC0C0FF



Private mAlpha As Long

' Declaraciones para Layered Windows (sólo Windows 2000 y superior)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2
'
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

'------------------------------------------------------------------------------
Private Const GWL_EXSTYLE = (-20)

' Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
'    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
'    (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Private Const RDW_INVALIDATE = &H1
Private Const RDW_ERASE = &H4
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_FRAME = &H400

Private Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" _
    (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, _
    ByVal fuRedraw As Long) As Long

Dim document() As New Form1



Private Function GetFileNameURL(URL As String)
On Error Resume Next
Dim ret As Long
Dim sName As String
ret = InStrRev(URL, "/")
If ret Then
sName = Mid(URL, ret + 1)
GetFileNameURL = GetStringFile(sName)
End If
End Function


Public Function GetStringFile(Value As String, Optional Force As Boolean = True) As String
    On Error Resume Next
    Dim HEX As String, ret As Long
    
    If Force Then Value = Replace(Value, "+", " ")
    Value = Replace(Value, "%25", Chr(0))
    
    ret = InStr(Value, "%")
    
    Do While ret > 0
        ret = InStr(Value, "%")
        If ret <> 0 Then
            HEX = Mid(Value, ret + 1, 2)
            Value = Replace(Value, "%" & HEX, Chr("&H" & HEX))
        End If
    Loop
    
    Value = Replace(Value, Chr(0), "%")
    
    Value = Replace(Value, "/", "\")
    
    GetStringFile = Replace(Value, " ", "_")
End Function

Public Sub carga_impresoras()
On Error Resume Next

Dim cImprGen As String
    cImprGen = cboimpre.Text
    
cboimpre.Clear
ruta$ = "c:\vacations\"
    
If Dir$(ruta$ + "printer") <> "" Then
 nf = FreeFile
 Open ruta$ + "printer" For Input Shared As #nf
 Lock #nf
 Line Input #nf, P1$
 Line Input #nf, P2$
 Unlock #nf
 Close #nf
 
 cImprGen = P1$
 cboimpre.Text = P1$

End If
    
    
    
    
For Each xprint In Printers
           If xprint.DeviceName = cImprGen Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next
        
        
        
For Each xprint In Printers
        cboimpre.AddItem xprint.DeviceName
Next
        
        
nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
 For t = 0 To cboimpre.ListCount - 1
   If cboimpre.List(t) = Printer.DeviceName Then
       cboimpre.ListIndex = t
       Exit For
   End If
 Next t
        
        
        
        
End Sub

Public Sub envia_correo2()
On Error Resume Next


 
      
      '  +++++++++++++++++++++++++++++++++++++++++++++++++
      
      fuente_original$ = App.Path & "\"
      fuente$ = "c:\vacations\"
      
      If Dir$(fuente$ + "nueva.htm") = "" Then
        FileCopy fuente_original$ + "nueva.htm", fuente$ + "nueva.htm"
      End If
      
      ' FileCopy App.Path & "\config.ini", fuente$ + "config.ini"
      
      
      Name fuente$ + "nueva.htm" As fuente$ + "nueva2.htm"

      nf2 = FreeFile
      Open fuente$ + "nueva.htm" For Output Shared As #nf2


      msg2$ = "</p><style type=" + Chr$(34) + "text/css" + Chr$(34) + "> <!--.Estilo1 {font-family: " + Chr$(34) + "Courier New" + Chr$(34) + "}--></style><span class=" + Chr$(34) + "Estilo1" + Chr$(34) + ">"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2


      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
                    
      
      
      msg2$ = "</p> This is only a test... </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p><b> If you can read this message, that means everything is well </b></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
           
            
      msg2$ = "</p> >> VACATIONS PROGRAM <<" + Space(1) + " </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
       
       Close nf2

      G = valido1
      
      
      'If correo_admin$ <> "" Then
            
       '  transfiere$ = correo_admin$
      
     ' End If
      
      
      transfiere$ = "IT@justautoins.com"
      
      
      send_email
      

End Sub

Public Sub envia_correo()
On Error Resume Next

Exit Sub
'  *************** NO SE USA  *******************************
' ************************************************************
' ************************************************************


Dim sSelect As String
Dim Rs As ADODB.Recordset
    
    
Set Rs = New ADODB.Recordset
   
 
      
      '  +++++++++++++++++++++++++++++++++++++++++++++++++
      
      fuente_original$ = App.Path & "\"
      fuente$ = "c:\vacations\"
      
      If Dir$(fuente$ + "nueva.htm") = "" Then
        FileCopy fuente_original$ + "nueva.htm", fuente$ + "nueva.htm"
      End If
      
      ' FileCopy App.Path & "\config.ini", fuente$ + "config.ini"
      
      
      Name fuente$ + "nueva.htm" As fuente$ + "nueva2.htm"

      nf2 = FreeFile
      Open fuente$ + "nueva.htm" For Output Shared As #nf2
      
      
      
      a$ = ID_vacaciones$
      
      
      
     sSelect = "select * from vacationsprogram where idvacation='" + ID_vacaciones$ + "' and active='1'"
   
      
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
    
    If grid3.Rows > 1 Then
       
       grid3.Row = 2
       grid3.Col = 2
       id_employee$ = grid3.Text
       
       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + id_employee$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre_empleado$ = Rs(0)
       apellido_empleado$ = Rs(1)
       Rs.Close
       nombre_empleado$ = nombre_empleado$ + " " + apellido_empleado$
              
       
       grid3.Col = 3
       ID_manager$ = grid3.Text
       
       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + ID_manager$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre_manager$ = Rs(0)
       apellido_manager$ = Rs(1)
       Rs.Close
       nombre_manager$ = nombre_manager$ + " " + apellido_manager$
       
       
       grid3.Col = 5
       fecha$ = Format(grid3.Text, "mm/dd/yyyy")
       
       grid3.Col = 6
       horas$ = grid3.Text
       
       grid3.Col = 7
       Status_aprobado$ = grid3.Text
       
       If Status_aprobado$ = "1" Or Status_aprobado$ = True Then
           Status_aprobado$ = "Approved"
       Else
          Status_aprobado$ = "Pending"
       End If
       
       
       
       
    Else
        Exit Sub
    End If
    
    
        
        
      


      msg2$ = "</p><style type=" + Chr$(34) + "text/css" + Chr$(34) + "> <!--.Estilo1 {font-family: " + Chr$(34) + "Courier New" + Chr$(34) + "}--></style><span class=" + Chr$(34) + "Estilo1" + Chr$(34) + ">"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2


      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
                    
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Requested date: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + fecha$ + "</b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Requested hours: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + horas$ + "</b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      'msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Customer name: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + LTrim(RTrim(UCase(txtnombre.Text))) + " " + LTrim(RTrim(UCase(txtapellido.Text))) + "</b></p></font>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
      
      'msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Customer ID: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + txtcust_id.Text + "</b></p></font>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
      
      msg2$ = "</p> ------------------------------"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p><h2></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Status: </font><font color=" + Chr$(34) + "red" + Chr$(34) + "><b>" + Status_aprobado$ + "</b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
            
      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      msg2$ = "</p> ------------------------------"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
      'msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> TXN type: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + Right(cbo_type_trans.List(cbo_type_trans.ListIndex), Len(cbo_type_trans.List(cbo_type_trans.ListIndex)) - 3) + "</b></p></font>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
      
           
      'msg2$ = "</p>&nbsp;</p>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
            
      
      X1$ = LTrim(RTrim(Left(TxtSubject.Text, Len(TxtSubject.Text))))
      X2$ = LTrim(RTrim(Left(Form1.lblmanager.Caption, Len(Form1.lblmanager.Caption))))
      
      If UCase(X1$) = UCase(X2$) Then
      
        msg2$ = "</p> Agent/Manager: " + X1$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
        
      Else
      
        msg2$ = "</p> Agent: " + nombre_empleado$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
      
        msg2$ = "</p> Manager: " + nombre_manager$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
        
      End If
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      GoTo salta_esto
            
      ' ------------------------------
      
       msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p>I M P O R T A N T : </p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      ultimo_mensaje$ = txtcomment.Text
      
      
      
      
      ' separa mensaje en lineas
      conta_caracteres = 0
      R$ = ""
      For t = 1 To Len(ultimo_mensaje$)
        R$ = R$ + Mid$(ultimo_mensaje$, t, 1)
        conta_caracteres = conta_caracteres + 1
        If conta_caracteres >= 28 And Mid$(ultimo_mensaje$, t, 1) = Space(1) Then
             msg2$ = "</p><u><b><i>" + R$ + "</i></b></u></p>"
             Lock #nf2
             Print #nf2, msg2$
             Unlock #nf2
             
             R$ = ""
             conta_caracteres = 0
         End If
      Next t
      
      If conta_caracteres > 0 Then
          msg2$ = "</p><u><b><i>" + R$ + "</i></b></u></p>"
          Lock #nf2
          Print #nf2, msg2$
          Unlock #nf2
      End If
      
       
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      ' ---------------------------------------------------
      
            
salta_esto:
            
            
      msg2$ = "</p><h4></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
            
      msg2$ = "</p> Please submitted on payroll to be paid" + Space(1) + " </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
      msg2$ = "</p> https://secure5.yourpayrollhr.com/ta/JAI04.login" + Space(1) + " </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
            
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
         
            
            
      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
            
      msg2$ = "</p> >> PLEASE, CHECK VACATION PROGRAM <<" + Space(1) + " </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
       
       Close nf2

      G = valido1
      
      
      
      
      
     
    
    
    
      
      
      If correo_admin$ <> "" Then
            
         If correo_agente$ = correo_manager$ Then
             transfiere$ = correo_admin$ + ";" + correo_manager$
         Else
             transfiere$ = correo_admin$ + ";" + correo_agente$ + ";" + correo_manager$
         End If
         
      Else
      
         If correo_agente$ = correo_manager$ Then
             transfiere$ = correo_manager$
         Else
             transfiere$ = correo_agente$ + ";" + correo_manager$
         End If
      
      
      End If
      
      
      
      
      
      'transfiere$ = "hnavarro@justautoins.com"
      
     
      send_email
     
End Sub


Private Function Random(min, max) As Long
    Random = CInt((max - min + 1) * Rnd + min)
End Function






Private Sub btn_configurar_accesos_Click()
On Error Resume Next

Load Form3
Form3.Show 1

carga_GI_anuales

End Sub

Private Sub btn_hoy_Click()
On Error Resume Next
 

   R$ = Format(Now, "mm/dd/yyyy")
   ucCalendar1.DateValue = R$
   
    
       
End Sub



Private Sub btn_menu_Click(Index As Integer)
On Error Resume Next
tipo_ejercicio = Index

If Index = 0 Then
  'Shape2.FillColor = &HE0E0E0
  'Picture2.BackColor = &HE0E0E0
Else
   'Shape2.FillColor = &HC0C0C0
   'Picture2.BackColor = &HC0C0C0
End If


End Sub


Private Sub btn_vacaciones_info_Click()
On Error Resume Next
If cbo_users.ListIndex = -1 Then
   Exit Sub
End If




Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
    
    pos = InStr(1, cbo_users.List(cbo_users.ListIndex), "  ")
    R$ = Left(cbo_users.List(cbo_users.ListIndex), pos - 1)
        
    lblagente.Caption = R$
    
    current_year$ = Format(Now, "yyyy")
    
    
    ' carga todo al inicio y despues solamente los registros nuevos o modificados
    
      
    sSelect = "select daterequested from VacationsProgram vac " & _
    "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
    "where emp.Active=1 and vac.active=1 and vac.idemployee='" + lblemp_id.Caption + "' and year(daterequested)>='" + Format(Now, "yyyy") + "' order by daterequested, approved"
    
    '   and year(daterequested)='" + current_year$ + "' "
   
   
   
    Rs.Open sSelect, base, adOpenUnspecified
   
        
     ' Permitir redimensionar las columnas
    grid5.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid5.DataSource = Rs
                         
   Rs.Close

enca1
Picture5.Visible = True



End Sub

Public Sub enca1()
On Error Resume Next


grid5.ColWidth(0) = 400
grid5.ColAlignment(0) = flexAlignLeftCenter


grid5.ColWidth(1) = 1200 'idreceiptHDR
grid5.ColAlignment(1) = flexAlignLeftCenter
 

grid5.Row = 0

grid5.Col = 1
grid5.Text = "DATE"

For t = 1 To grid5.Rows - 1
   grid5.Row = t
   grid5.Col = 0
   grid5.Text = t
   
   grid5.Col = 1
   R$ = Format(grid5.Text, "mm/dd/yyyy")
   grid5.Text = R$
Next t

grid5.FixedRows = 1
grid5.FixedCols = 1

grid5.Row = 1
grid5.Col = 1

End Sub


Private Sub btnacceso_empleados_Click()
On Error Resume Next
Load Form5
Form5.Show 1
carga_accesos

End Sub

Private Sub btnbloquear_dia_Click()
On Error Resume Next
Dim StartDate As Date, EndDate As Date
 
 
 
    If transfiere$ <> "" And evento <> "" Then
    Else
      transfiere$ = ""
    End If
   
   

 
 
    
    
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then
        
        StartDate = transfiere$
        If transfiere$ = "" Then Exit Sub
        'Exit Sub
    Else
    
    
    End If

If Weekday(StartDate) = vbSunday Then
   MsgBox "You cannot choose Sunday as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   Exit Sub

 End If




    Timer1.Enabled = False
         
        transfiere$ = StartDate
        
        
        
        With Form4
            .DTPStartDate = DateValue(DateValue(StartDate))
            .DTPStartTime = StartDate
            .DTPEndDate = DateValue(DateValue(EndDate))
            .DTPEndTime = EndDate
            
        End With
        
 

       Form4.Show vbModal, Me

       Timer1.Enabled = True
       
End Sub

Private Sub btnconfig_mail_Click()
On Error Resume Next
Load Forma_seguridad
Forma_seguridad.Show 1

If transfiere$ = "JA789!" Then
  Timer1.Enabled = False
  Load FrmConfig
  FrmConfig.Show 1
  Timer1.Enabled = True
Else
  MsgBox "Invalid password", 16, "Access denied"
End If



End Sub





Private Sub btnletra_Click(Index As Integer)
On Error Resume Next
If Index = 26 Then
  letra$ = "-"
Else
  letra$ = btnletra(Index).Caption
End If

carga_empleados

lblagente.Caption = ""
grid5.Clear

End Sub

Private Sub btnprint_Click()
On Error Resume Next


Printer.FontName = "Courier new"
' Printer.FontName = "Arial"

encabezado:

Printer.Orientation = 2  '  Landscape

'grid5.Width = Printer.Width

pagina = pagina + 1
linea = 6
Printer.FontSize = 24


altura = Printer.ScaleHeight
anchura = Printer.ScaleWidth + 800


' modifica el ancho del grid



Image1.Visible = False
Image3.Visible = False

Image4.Visible = False
calendario.Visible = False
btn_hoy.Visible = False

btn_menu(0).Visible = False
btn_menu(1).Visible = False

Picture1.Visible = True
Refresh

'Refresh
'Form1.Refresh

Me.PrintForm
'Printer.Print ucCalendar1
'Printer.PaintPicture ucCalendar1, size_screen, 1600
Printer.FontName = "Courier new"

' Printer.Orientation = vbPRORPortrait

Printer.EndDoc


Image1.Visible = True
Image3.Visible = True

Image4.Visible = True
calendario.Visible = True
btn_hoy.Visible = True

btn_menu(0).Visible = True
'btn_menu(1).Visible = True


Refresh

End Sub

Private Sub btnsend_Click()
On Error Resume Next


Load Forma_seguridad
Forma_seguridad.Show 1

If transfiere$ = "JA789!" Then
  envia_correo2
Else
  MsgBox "Invalid password", 16, "Access denied"
End If


End Sub



Private Sub cbo_users_Click()
On Error Resume Next

If cbo_users.ListIndex = -1 Then
   Exit Sub
End If


lblagente.Caption = ""
grid5.Clear

cbo_users.SelLength = 0
lblemployee.Caption = RTrim(LTrim(Left(cbo_users.List(cbo_users.ListIndex), 30)))
lblemp_id.Caption = RTrim(LTrim(Right(cbo_users.List(cbo_users.ListIndex), 6)))

lbllocation.Caption = RTrim(LTrim(Mid(cbo_users.List(cbo_users.ListIndex), 41, 25)))

End Sub


Private Sub cboimpre_Click()
On Error Resume Next


For Each xprint In Printers
           If xprint.DeviceName = cboimpre.Text Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next


nf = FreeFile
 Open "c:\vacations\printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
End Sub


Private Sub Check1_Click()
On Error Resume Next

nf = FreeFile
Open "c:\vacations\reading" For Output Shared As #nf
Lock #nf
Print #nf, Check1.Value
Unlock #nf
Close nf





End Sub

Private Sub Check2_Click()
R$ = ""
nf = FreeFile
Open "c:\vacations\importance" For Output Shared As #nf
Lock #nf
Print #nf, Check2.Value
Unlock #nf
Close nf


End Sub






Private Sub chk_permitir_en_2_meses_seguidos_Click()
On Error Resume Next
permiso_doble_mes = chk_permitir_en_2_meses_seguidos.Value
End Sub

Private Sub lista_acceso_Click()
On Error Resume Next
nf = FreeFile

Open "\\192.168.84.215\vacations\Lista_acceso" For Output Shared As #nf
Lock #nf
For t = 0 To lista_acceso.ListCount - 1
   Print #nf, lista_acceso.List(t)
   Print #nf, lista_acceso.Selected(t)
Next t
Unlock #nf
Close #nf
   


End Sub

Private Sub op_ano_Click(Index As Integer)
On Error Resume Next
ano_para_GI = Val(op_ano(Index).Caption)
carga_GI_anuales
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
If seg >= 60 Then
   actualiza_fechas
   seg = 0
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
seg = seg + 1
If seg = 3 Then
  valido1 = 777
  mensaje.Visible = True
  mensaje.Refresh
  Carga_fechas
  Timer1.Enabled = True
  Timer2.Enabled = False
  mensaje.Visible = False
  btn_menu(0).Enabled = True
  valido1 = 0
  carga_aniversarios
  seg = 0
End If


End Sub


Private Sub ucCalendar1_Click()

On Error Resume Next

 Dim StartDate As Date, EndDate As Date

ucCalendar1.SelectionColor = &HC0C0C0
 
If UCase(lblemployee.Caption) = "GABRIELA JIMENEZ" Then
  ' Exit Sub
End If
 
 
 
 
   
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then
       Exit Sub
    End If
    
 valido1 = 0
    
    
 If Weekday(StartDate) = vbSunday Then
   MsgBox "You cannot choose Sunday as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   Exit Sub
   
 End If
 
 mes = Val(Format(StartDate, "mm"))
 dia = Val(Format(StartDate, "dd"))
 
 
 
 ' AÑO NUEVO
 If mes = 1 And dia = 1 Then
   MsgBox "Sorry, but the first day of January is not allowed to be taken as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   titulo_del_dia$ = "New Year's Day"
   Exit Sub
 End If
 
 
 ' NAVIDAD
 If mes = 12 And dia = 25 Then
   MsgBox "Sorry, but December 25 cannot be taken as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   titulo_del_dia$ = "Christmas Day"
   Exit Sub
 End If
 
 
 
 ' DICIEMBRE
 'If mes = 12 And dia >= 23 Then
 '  MsgBox "Sorry, but these days of December cannot be taken as a vacation day", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 ' ENERO
 'If mes = 1 And dia < 5 Then
 '  MsgBox "Sorry, but these days of January cannot be taken as a vacation day", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 
 
 ' FEBRERO
 'If mes = 2 And dia >= 19 Then
 '  MsgBox "Sorry, but these days of the month of February cannot be taken on vacation", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 ' MARZO
 'If mes = 3 Then
 '   MsgBox "Sorry, but you can't take vacations this month", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
  
 
 ' ABRIL
 'If mes = 4 And dia <= 15 Then
 '  MsgBox "Sorry, but these days of the month of April cannot be taken on vacation", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 ' JULIO
 'If mes = 7 And (dia = 7 Or dia = 8 Or dia = 9 Or dia = 10 Or dia = 11 Or dia = 12 Or dia = 13 Or dia = 21 Or dia = 22 Or dia = 23 Or dia = 24 Or dia = 25 Or dia = 26 Or dia = 27) Then
 '  MsgBox "Sorry, but these days of July cannot be taken as a vacation day", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 '  AGOSTO
 'If mes = 8 And (dia = 4 Or dia = 5 Or dia = 6 Or dia = 7 Or dia = 8 Or dia = 9 Or dia = 10 Or dia = 18 Or dia = 19 Or dia = 20 Or dia = 21 Or dia = 22 Or dia = 23 Or dia = 24) Then
 '  MsgBox "Sorry, but these days of August cannot be taken as a vacation day", 16, "Attention"
 '  dia_bloqueado = 1
 '  Exit Sub
 'End If
 
 
 
 ' THANKSGIVING DAY
 ' calcular dia de accion de gracias
 cuenta = 0
 ano$ = Val(Format(Now, "yyyy"))
 For t = 1 To 30
    fecha$ = Format("11/" + Format(t, "00") + "/" + ano$, "mm/dd/yyyy")
    dia_semana = Weekday(fecha$)
    If dia_semana = vbThursday Then
       cuenta = cuenta + 1
       If cuenta = 4 Then
          dia_thanks = t
          Exit For
       End If
    End If
 Next t
 
 
    ' calcular dia de accion de gracias del siguiente año
 cuenta = 0
 ano2$ = Format(Val(Format(Now, "yyyy")) + 1, "0000")
 
 For t = 1 To 30
    fecha$ = Format("11/" + Format(t, "00") + "/" + ano2$, "mm/dd/yyyy")
    dia_semana = Weekday(fecha$)
    If dia_semana = vbThursday Then
       cuenta = cuenta + 1
       If cuenta = 4 Then
          dia_thanks2 = t
          Exit For
       End If
    End If
 Next t
    
    
 
 If mes = 11 And dia = dia_thanks And Val(ano$) = Val(Format(StartDate, "yyyy")) Then
   MsgBox "Sorry, but Thanksgiving day cannot be taken as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   titulo_del_dia$ = "Thanksgiving Day"
   Exit Sub
 End If
 
 
 If mes = 11 And dia = dia_thanks2 And Val(ano2$) = Val(Format(StartDate, "yyyy")) Then
   MsgBox "Sorry, but Thanksgiving day cannot be taken as a vacation day", 16, "Attention"
   dia_bloqueado = 2
   titulo_del_dia$ = "Thanksgiving Day"
   Exit Sub
 End If
 


 dia_bloqueado = 0
 titulo_del_dia$ = ""

End Sub

Private Sub ucCalendar1_DateBackColor(CellDate As Date, Color As stdole.OLE_COLOR, eHatchStyle As HatchStyle)
On Error Resume Next

    If Weekday(CellDate) = vbSunday Then  'Or Weekday(CellDate) = vbSaturday Then
        Color = &HF7F7F7
    End If
    
    
    
    If ucCalendar1.ViewMode = vm_Week Then
        
        If VBA.TimeValue(CellDate) > CDate("00:00") And VBA.TimeValue(CellDate) <= CDate("08:00") Then
            'Color = &HF7F7F7
        End If
        
    Else
    
        If CellDate = CDate("04/01/2023") Then
           
            'Color = &HCCCCCC
          '  eHatchStyle = 2
        End If
        
    End If
    
    
End Sub

Private Sub ucCalendar1_DropDownViewMore(CellDate As Date, CancelViewModeDay As Boolean)
'    Dim Subject As String, StartDate As Date, EndDate As Date, Color As Long, AllDayEvent As Boolean, body As String
'    Dim IsSerie As Boolean, IsPrivate As Boolean, bNotify As Boolean, EventShowAs As eEventShowAs
'    List1.Clear
'    List1.Move Me.ScaleWidth / 2 - List1.Width / 2, Me.ScaleHeight / 2 - List1.Height / 2
'    List1.Visible = True
'    Dim cCol As Collection
'    Dim i As Long
'    CancelViewModeDay = True
'    Set cCol = ucCalendar1.GetEventsFromDay(CellDate)
'    For i = 1 To cCol.Count
'        If ucCalendar1.GetEventData(cCol.Item(i), Subject, StartDate, EndDate, Color, AllDayEvent, body, , IsSerie, bNotify, IsPrivate, EventShowAs) Then
'            List1.AddItem Subject
'        End If
'    Next


End Sub

Private Sub ucCalendar1_EventChangeDate(ByVal EventKey As Long, ByVal StartDate As Date, ByVal EndDate As Date, ByVal AllDay As Boolean)
    'On Error Resume Next
  '  Debug.Print EventKey, StartDate, EndDate
    
    
End Sub

Private Sub ucCalendar1_PreEventChangeDate(ByVal EventKey As Long, ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
   ' Debug.Print StartDate, EndDate
    If DateValue(StartDate) <= CDate("28/4/2022") And DateValue(EndDate) >= CDate("28/4/2022") Then
        Cancel = True
        'Debug.Print StartDate, EndDate
    End If
    
    If Weekday(StartDate) = vbSaturday And VBA.TimeValue(StartDate) < CDate("08:00") Then
        Cancel = True
    End If
    
    
    
    
    
End Sub



Private Sub ucCalendar1_EventClick(ByVal EventKey As Long, Button As Integer)
    On Error Resume Next

    Dim Subject As String, StartDate As Date, EndDate As Date, Color As Long, AllDayEvent As Boolean, body As String
    Dim IsSerie As Boolean, IsPrivate As Boolean, bNotify As Boolean, EventShowAs As eEventShowAs
    Dim office As String, idvacaciones As Integer
    'ucCalendar1.RemoveEvent EventKey
    
    
    If Button <> vbLeftButton Then
      Exit Sub
    End If
    
    If UCase(lblemployee.Caption) = "GABRIELA JIMENEZ" Then
       'lblmsg.Caption = "Loading the information..."
       
      
    Else
    
      'lblmsg.Caption = "This day is blocked"
      
    End If
    
    
    Timer1.Enabled = False
    
    mensaje.Visible = True
    mensaje.Refresh
     
     verifica_disponibilidad
       
     verifica_numero_de_managers
     
     
     
    
    If ucCalendar1.GetEventData(EventKey, Subject, StartDate, EndDate, Color, AllDayEvent, body, , IsSerie, bNotify, IsPrivate, EventShowAs, office, idvacaciones) Then
        valido1 = 777
        
        If UCase(usuario$) = "GABRIELA JIMENEZ" Then ' Or lblemp_id.Caption = "76" Then
            
            transfiere$ = StartDate
            evento = EventKey
            mensaje.Visible = False
            btnbloquear_dia_Click
            
            Timer1.Enabled = True
            transfiere$ = ""
            Exit Sub
        End If
        
        
        If usuario$ = "Christmas Day" Then
            transfiere$ = StartDate
            evento = EventKey
            mensaje.Visible = False
            
            
            Timer1.Enabled = True
            transfiere$ = ""
            MsgBox "Sorry, but December 25 cannot be taken as a vacation day", 16, "Attention"
            valido1 = 5
            Exit Sub
        ElseIf usuario$ = "New Year's Day" Then
            transfiere$ = StartDate
            evento = EventKey
            mensaje.Visible = False
            
            
            Timer1.Enabled = True
            transfiere$ = ""
            MsgBox "Sorry, but the first day of January is not allowed to be taken as a vacation day", 16, "Attention"
            Exit Sub
            
         ElseIf usuario$ = "Thanksgiving Day" Then
            transfiere$ = StartDate
            evento = EventKey
            mensaje.Visible = False
            
            
            Timer1.Enabled = True
            transfiere$ = ""
            MsgBox "Sorry, but Thanksgiving day cannot be taken as a vacation day", 16, "Attention"
            Exit Sub
            
        End If
        
        
        
        
        
        With Form2
            .TxtSubject = Subject
            .TxtLocation = office
            .lblid_vac = ID_vacaciones$
            .DTPStartDate = DateValue(StartDate)
            .DTPStartTime = StartDate
            .DTPEndDate = DateValue(EndDate)
            .DTPEndTime = EndDate
            .ChkAllDay.Value = IIf(AllDayEvent, 1, 0)
            .EventKey = EventKey
            .TxtBody = nota$
            .PicColor.BackColor = colorx
            .ChkIsSerie.Value = IIf(IsSerie, 1, 0)
            .ChkPrivate.Value = IIf(IsPrivate, 1, 0)
            .ChkNotify.Value = IIf(bNotify, 1, 0)
            .ImageCombo1.ComboItems.Item(EventShowAs + 1).Selected = True
        End With
        valido1 = 1
        
        
        If IsPrivate = False Then
           fecha_accesible = 1
        'Else
           'fecha_accesible = 0
        End If
        
        mensaje.Visible = False
        
        Form1.pantalla_transparente
        Form2.Show vbModal, Me
        Form1.pantalla_solida
        
        
       If transfiere$ = "ENVIA CORREO2" Then
        Form1.envia_correo_aprobado
        
        transfiere$ = ""
       End If
        
    End If
    
    Timer1.Enabled = True
   ' lblmsg.Caption = "Loading the information..."
    mensaje.Visible = False
    'ucCalendar1.Refresh
End Sub

Private Sub ucCalendar1_EventMouseEnter(ByVal EventKey As Long)
    'Debug.Print "1 EventMouseEnter", EventKey
  
End Sub

Private Sub ucCalendar1_EventMouseLeave(ByVal EventKey As Long)
    'Debug.Print "2 EventMouseLeave", EventKey
End Sub

Private Sub ucCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
  
   Exit Sub
   
    Dim StartDate As Date, EndDate As Date
   
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then Exit Sub
    
    If Button = vbRightButton Then
        With Form2
            .TxtSubject = lblemployee.Caption
            .DTPStartDate = DateValue(StartDate)
            .DTPStartTime = StartDate
            .DTPEndDate = DateValue(EndDate)
            .DTPEndTime = EndDate
            If ucCalendar1.ViewMode = vm_Month Or (StartDate = EndDate) Then
                .op_dias(1).Value = True
                '.ChkAllDay.Value = 1
            End If
            .EventKey = 0
        End With
        
        If Format(Form2.DTPStartDate, "mm/dd/yyyy") = Format(Form2.DTPEndDate, "mm/dd/yyyy") Then
            Form2.op_dias(0).Value = True
        End If
        
        Form1.pantalla_transparente
        Form2.Show vbModal, Me
        Form1.pantalla_solida
        
    End If
End Sub

Private Sub ucCalendar1_PreDateChange(NewDate As Date, Cancel As Boolean)
On Error Resume Next
ano_actual = Val(Format(Now, "yyyy"))
ano_nuevo = Val(Format(NewDate, "yyyy"))

If ano_nuevo >= (ano_actual + 2) Then
   MsgBox "You cannot select vacations in that year yet", 16, "Upss!"
   NewDate = Format("12/31/" + Format(ano_nuevo - 1, "0000"), "mm/dd/yyyy")
   
End If


If ano_nuevo <= (ano_actual - 2) Then
   MsgBox "You cannot select vacations in that year yet", 16, "Upss!"
   NewDate = Format("01/01/" + Format(ano_actual - 1, "0000"), "mm/dd/yyyy")
   
End If




'    With ucCalendar1
'        .Redraw = False
'        .Clear
'        .AddEvents "Dinamico", NewDate, NewDate + 1, &HBFA7E3, False
'        .Redraw = True
'    End With
End Sub



Private Sub btnadd_event_Click()
On Error Resume Next
 
 Dim StartDate As Date, EndDate As Date
 
 
 valido1 = 0
 
 If cbo_users.ListIndex >= 0 Then
   If lblemp_id.Caption = "76" Then
     lblemployee.Caption = RTrim(LTrim(Left(cbo_users.List(cbo_users.ListIndex), 30)))
     lblemp_id.Caption = RTrim(LTrim(Right(cbo_users.List(cbo_users.ListIndex), 10)))
     lbllocation.Caption = RTrim(LTrim(Mid(cbo_users.List(cbo_users.ListIndex), 41, 25)))
  
     GoTo acceso_permitido
   End If
 End If
 
 
 If lblemp_id.Caption = "76" Then
  Exit Sub
 End If
 
 
acceso_permitido:
 
 
   
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then
       If user$ = "GJIMENEZ" Then
          lblemployee.Caption = lblmanager.Caption
          lblemp_id.Caption = lblmanager_ID.Caption
          lbllocation.Caption = "JA - HAVEN"
          cbo_users.ListIndex = -1
       End If
       
       Exit Sub
    End If
    
     Timer1.Enabled = False
    
    
    verifica_dias_seleccionados
       
       
    If bloquea_acceso = 1 Then
       MsgBox "I'm sorry, but you cannot select more than 5 consecutive days", 16, "Attention"
       bloquea_acceso = 0
       Timer1.Enabled = True
       
       If user$ = "GJIMENEZ" Then
          lblemployee.Caption = lblmanager.Caption
          lblemp_id.Caption = lblmanager_ID.Caption
               lbllocation.Caption = "JA - HAVEN"
               cbo_users.ListIndex = -1
       End If
    
       Exit Sub
    
    End If
    
    
    
    
    verifica_disponibilidad
    
     verifica_numero_de_managers
    
    
    If dia_bloqueado >= 1 Or valido1 = 5 Then
      
      If chk_dia_doble.Value = 0 Or dia_bloqueado = 2 Then
      
        Timer1.Enabled = True
       
        If user$ = "GJIMENEZ" Then
          lblemployee.Caption = lblmanager.Caption
          lblemp_id.Caption = lblmanager_ID.Caption
               lbllocation.Caption = "JA - HAVEN"
               cbo_users.ListIndex = -1
               
        End If
       
        Exit Sub
      
      End If
        
    End If
 
 
    
    
    If fecha_accesible >= 2 Then
       MsgBox "Sorry, this date is no longer available", 16, "Upss!"
       Timer1.Enabled = True
       
       If user$ = "GJIMENEZ" Then
          lblemployee.Caption = lblmanager.Caption
          lblemp_id.Caption = lblmanager_ID.Caption
               lbllocation.Caption = "JA - HAVEN"
               cbo_users.ListIndex = -1
       End If
    
       Exit Sub
    ElseIf fecha_accesible = 1 Then
    
    End If
    
    
    
    
    
    
    
    
        ID_vacaciones$ = ""
        nota$ = ""
       
        
        If valido1 = 5 Then
           valido1 = 0
           If user$ = "GJIMENEZ" Then
               lblemployee.Caption = lblmanager.Caption
               lblemp_id.Caption = lblmanager_ID.Caption
               lbllocation.Caption = "JA - HAVEN"
               cbo_users.ListIndex = -1
           End If
           Exit Sub
           
        End If
        
        
    
        With Form2
            .lblid_vac = ID_vacaciones$
            .TxtSubject = lblemployee.Caption
            .TxtLocation = lbllocation.Caption
            .DTPStartDate = DateValue(StartDate)
            .DTPStartTime = StartDate
            .DTPEndDate = DateValue(EndDate)
            .DTPEndTime = EndDate
            If ucCalendar1.ViewMode = vm_Month Or (StartDate = EndDate) Then
                .op_dias(0).Value = True
                '.ChkAllDay.Value = 1
            End If
            .EventKey = 0
        End With
        
        If Format(Form2.DTPStartDate, "mm/dd/yyyy") = Format(Form2.DTPEndDate, "mm/dd/yyyy") Then
           ' Form2.op_dias(0).Value = True
        End If
        
        Form1.pantalla_transparente
        Form2.Show vbModal, Me
        Form1.pantalla_solida
       
    
       If transfiere$ = "ENVIA CORREO1" Then
        Form1.envia_correo_aprobado
        transfiere$ = ""
       End If
    
       Timer1.Enabled = True
    
        a$ = user$
        
        If user$ = "GJIMENEZ" Then
          lblemployee.Caption = lblmanager.Caption
          lblemp_id.Caption = lblmanager_ID.Caption
               lbllocation.Caption = "JA - HAVEN"
               cbo_users.ListIndex = -1
               'reservado = 0
       End If
    
    
       
    
    
End Sub

Private Sub btnfin_Click()
On Error Resume Next

base.Close
End

End Sub

Private Sub Form_Load()
On Error Resume Next

top = 0
Left = (Screen.Width - Width) / 2
tipo_ejercicio = 0

carga_inicial = 0
Erase llavero
letra$ = "-"

   Refresh
  Set mcIni = New clsIni

  nf = FreeFile
  Open "\\192.168.84.215\vacations\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\vacations\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     R$ = Shell("\\192.168.84.215\vacations\actualizador.exe", vbNormalFocus)
     
     Hide
     Refresh
     End
     Exit Sub
  End If
  
  

If (App.PrevInstance = True) Then
  'base.Close
  End
End If

FileCopy "\\192.168.84.215\vacations\imagen2.jpg", "c:\vacations\imagen2.jpg"
FileCopy "\\192.168.84.215\vacations\config.ini", "c:\vacations\config.ini"
FileCopy "\\192.168.84.215\vacations\imagen.jpg", "c:\vacations\imagen.jpg"
FileCopy "\\192.168.84.215\vacations\important.jpg", "c:\vacations\important.jpg"





base.Close
Conecta_SQL


carga_employee




ano_para_GI = Val(op_ano(0).Caption)


Status = 0


Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 940 '940
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
        'ScaleFactorX = (Xpixels / DesignX)
        'ScaleFactorY = (Ypixels / DesignY)
      
        If Xpixels <= 1366 Then  ' Si es laptop
      
           ScaleFactorX = 980 / DesignX  ' 1360
           ScaleFactorY = 680 / DesignY   ' 1024
        
        Else  ' Si es Desktop con monitor de alta resolucion
          
           ScaleFactorX = 1360 / DesignX
           ScaleFactorY = 1024 / DesignY
        
        
        End If
        
        
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        Forma_main.Height = 9000 'Me.Height ' Remember the current size
        Forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0


' Carga_fechas


R$ = ""
nf = FreeFile
Open "c:\vacations\reading" For Input Shared As #nf
Lock #nf
Line Input #nf, R$
Unlock #nf
Close nf
Check1.Value = Val(R$)


R$ = ""
nf = FreeFile
Open "c:\vacations\importance" For Input Shared As #nf
Lock #nf
Line Input #nf, R$
Unlock #nf
Close nf
Check2.Value = Val(R$)


' btn_menu(1).Enabled = False
carga_empleados

If lblemp_id.Caption = "76" Then
  Picture2.Visible = True
  btnbloquear_dia.Visible = True
  btn_configurar_accesos.Visible = True
  btnprint.Enabled = True
  Frame1.Visible = True
  marco_anos.Visible = True
  op_ano(0).Caption = Format(Val(Format(Now, "yyyy")) - 1, "0000")
  op_ano(1).Caption = Format(Val(Format(Now, "yyyy")), "0000")
  Picture4.Visible = True
  
  ' btn_menu(1).Enabled = True
  
  Picture3.Visible = True
  
  carga_aniversarios
  
  
End If


carga_impresoras
carga_GI_anuales

   
  
 ''lista_acceso.Clear
 'lista_acceso.AddItem "JMIRELES"
 'lista_acceso.AddItem "LAURA"
 'lista_acceso.AddItem "MFUENTES"
 'lista_acceso.AddItem "JSALCIDO"
 'lista_acceso.AddItem "GPLASCENCIA"
 'lista_acceso.AddItem "YCLARA"
 'lista_acceso.AddItem "HMENDEZ"
 'lista_acceso.AddItem "NBAEZ"
 'lista_acceso.AddItem "MARIAA"
 'lista_acceso.AddItem "MARIA.ESPARZA"
 'lista_acceso.AddItem "JOSEV"
 'lista_acceso.AddItem "MROSARIO"
 'lista_acceso.AddItem "MORALESA"
 'lista_acceso.AddItem "LRODRIGUEZ"
 'lista_acceso.AddItem "NGARCIA"
 'lista_acceso.AddItem "AALVARADO"
 'lista_acceso.AddItem "BJIMENEZ"
 'lista_acceso.AddItem "MPOLANCO"
 'lista_acceso.AddItem "EVASQUEZ"
 'lista_acceso.AddItem "NDELATORRE"
 'lista_acceso.AddItem "ECARRILLO"
 'lista_acceso.AddItem "YROJAS"
 'lista_acceso.AddItem "AHUERTA"
 
 
 carga_accesos
 
 

Exit Sub


 With ucCalendar1
    
        
        

       ' .SetStrLanguage "Año", "Mes", "Semana", "Día", "Hoy", "Comienza", "Finaliza"

       Dim i As Long
       Dim StartDate As Date
       Dim EndDate As Date

       .Redraw = False
       
       '.AddEvents "All day", Date, Date + 1, vbamarillo, True
       '.AddEvents lblemployee.Caption, Date & " 08:00", Date & " 19:00", vbBlack
       .AddEvents "Crystal Coronel", Date & " 08:00", Date & " 19:00", vbBlack
       
     ' .AddEvents "Actual Event", Now, DateAdd("h", 2, Now), vbBlue
      ' .AddEvents "Future Event", DateAdd("h", 3, Now), DateAdd("h", 5, Now), vborange



    '   For i = 1 To 3000
    '        StartDate = DateAdd("d", i + Random(0, 1), Date - 7)
    '        StartDate = DateAdd("n", Random(0, 60 * 12), StartDate)
    '        EndDate = DateAdd("d", i + Random(1, 2), Date - 7)
    '        EndDate = DateAdd("n", Random(60 * 12, 60 * 24), EndDate)
    '        If StartDate < EndDate Then
    '            .AddEvents "Event " & i, StartDate, EndDate, QBColor(Random(1, 14))
    '        End If
    '   Next

       .Redraw = True
    End With
 
 
 
 
 

End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1


    'ucCalendar1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight


End Sub


Private Sub Form_Terminate()
On Error Resume Next
 base.Close
 
 
End Sub



Public Sub carga_employee()
On Error Resume Next


 Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
    
    sSelect = "select firstname, lastname1, idemployee, emailwork from employeeinfo where username='" + user$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
    
      nombre$ = RTrim(LTrim(Rs(0)))
      apellido$ = RTrim(LTrim(Rs(1)))
      id_emp$ = Rs(2)
      correo_agente$ = Rs(3)
      Rs.Close
   
      lblemployee.Caption = nombre$ + " " + apellido$
      lblemp_id.Caption = id_emp$
    
    
    ' carga la oficina donde esta ubicado
    
    
     sSelect = "select Office, IdJobTitle, hireDate  from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and Username='" + user$ + "'" ' and empjob.Active='1' and IdJobTitle in (16,17,18, 28,2,24,37)


    
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    If Grid2.Rows > 1 Then
       Grid2.Row = 2
       Grid2.Col = 1
       lbllocation.Caption = Grid2.Text
       oficina$ = Grid2.Text
       
       Grid2.Col = 2
       Id_jobtitle$ = Grid2.Text
       
       
       Grid2.Col = 3
       lblhired_date.Caption = Format(Grid2.Text, "mm/dd/yyyy")
       
    
 ' carga supervisor
 
 
   If Id_jobtitle$ = "17" Then
         ID_manager$ = "17"   'id_emp$
         lblmanager_ID.Caption = ID_manager$
         lblmanager.Caption = "Gabriela Jimenez"   'lblemployee.Caption
         lblmanager_ID.Caption = "76"  'lblemp_id.Caption
         correo_manager$ = "Gaby@justautoins.com"  'correo_agente$
         GoTo termina
   End If
   
   
   If Id_jobtitle$ = "14" Then
         ID_manager$ = "17"   'id_emp$
         lblmanager_ID.Caption = ID_manager$
         lblmanager.Caption = "Karina Delgadillo"   'lblemployee.Caption
         lblmanager_ID.Caption = "258"  'lblemp_id.Caption
         correo_manager$ = "HR@justautoins.com"  'correo_agente$
         GoTo termina
   End If
       
       
       Regional_manager$ = ""
       
       Select Case Val(Id_jobtitle$)
       Case 16  ' agente venta
          job_manager$ = "17"
         ' Regional_manager$ = " and username <> 'Gjimenez'"     ' and username <> 'mesparza'"
       Case 17
          job_manager$ = "17"
         ' Regional_manager$ = " and username <> 'Gjimenez'"
          
       Case 2, 37, 34 ' agente phone sales
          job_manager$ = "4, 5, 17"
          'Regional_manager$ = " and username = 'Vgalvano'"
       Case 1  ' IT
          job_manager$ = "1"
       Case 6  ' DMV
          job_manager$ = "7"
       Case 8, 10, 18, 35  ' UW agent
          job_manager$ = "9,11,19"
       Case 3, 32  ' agente comercial
          job_manager$ = "5 "
       Case 24 ' QC
          job_manager$ = "30, 17"
       Case 37 ' agente independiente
          job_manager$ = "5, 17"
       Case 38  ' registrations
          job_manager$ = "17"
       Case 36  ' accounting agent
          job_manager$ = "12"
       Case 14   ' HR
          job_manager$ = "14"
       Case 1, 5, 7, 9, 11, 12, 13, 14, 19, 27, 31
          job_manager$ = "1, 23"
          'Regional_manager$ = " and username = 'lfregoso'"
          
       Case 30  ' managers
          job_manager$ = "30"
          'Regional_manager$ = " and username = 'Gjimenez'"
          
       End Select
       
       
       
       
       
    Else
       lbllocation.Caption = "---"
       oficina$ = ""
    End If
       
    
 
 
 If Id_jobtitle$ = "16" Then
      
   sSelect = "select emp.IdEmployee, Username from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (" + job_manager$ + ") and ofc.office='" + oficina$ + "' and username<>'MESPARZA'  " '+ Regional_manager$
  
 Else
 
   sSelect = "select emp.IdEmployee, Username from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (" + job_manager$ + ") " '+ Regional_manager$
  
 
 End If


 ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    If Grid2.Rows > 1 Then
      
    
    
       Grid2.Row = 1
       Grid2.Col = 1
       ID_manager$ = Grid2.Text
       lblmanager_ID.Caption = Grid2.Text
       
       Grid2.Col = 2
       lblmanager.Caption = Grid2.Text
       
        sSelect = "select firstname, lastname1, idemployee, emailwork from employeeinfo where username='" + lblmanager.Caption + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
    
      nombre$ = RTrim(LTrim(Rs(0)))
      apellido$ = RTrim(LTrim(Rs(1)))
      id_emp$ = Rs(2)
      correo_manager$ = Rs(3)
      Rs.Close
   
     lblmanager.Caption = nombre$ + " " + apellido$
       
       
    Else
       
       
    End If
      

    
    If lblemployee.Caption = lblmanager.Caption Then
       
       'lblmanager_ID.Caption = lblemp_id.Caption
       'correo_manager$ = correo_agente$
       
       ID_manager$ = "76"   'id_emp$
         lblmanager_ID.Caption = ID_manager$
         lblmanager.Caption = "Karina Delgadillo"   'lblemployee.Caption
         lblmanager_ID.Caption = "258"  'lblemp_id.Caption
         correo_manager$ = "HR@justautoins.com"  'correo_agente$
       
    End If
    
    
    
    
    If lblemployee.Caption = lblmanager.Caption Then
       
       'lblmanager_ID.Caption = lblemp_id.Caption
       'correo_manager$ = correo_agente$
       
       ID_manager$ = "76"   'id_emp$
         lblmanager_ID.Caption = ID_manager$
         lblmanager.Caption = "Gabriela Jimenez"   'lblemployee.Caption
         lblmanager_ID.Caption = "76"  'lblemp_id.Caption
         correo_manager$ = "Gaby@justautoins.com"  'correo_agente$
       
    End If


termina:


End Sub


Public Sub Carga_fechas()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
    
    btn_menu(0).Enabled = False
    btn_menu(1).Enabled = False
    
    current_year$ = Format(Now, "yyyy")
    
    
    ' carga todo al inicio y despues solamente los registros nuevos o modificados
    
    If carga_inicial = 0 Then
      tipo_de_carga$ = ""
      
    Else
      tipo_de_carga$ = " and cast (vac.LastUpdated as date)='" + Format(Now, "mm/dd/yyyy") + "' "
    End If
    
   
    sSelect = "select idvacation, vac.idemployee, idmanager, approvedby, daterequested, hours, approved, notes, emp.payrolllinkid, emp.firstname, emp.lastname1, emp.username, emp.emailwork, cashout from VacationsProgram vac " & _
    "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
    "where emp.Active=1 and vac.active=1 " + tipo_de_carga$ + " order by daterequested, approved"
    
    '   and year(daterequested)='" + current_year$ + "' "
   
   
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
   Rs.Close
   
    
   total_registros = Grid2.Rows - 1
  ' total_registros = List1.ListCount
    
   barra.max = total_registros + 1
   barra.min = 1
      
      
      
   Dim autorizado As Boolean
   
   
   ucCalendar1.Visible = False
   ' ucCalendar1.Clear
   
   
   
    ' calcular dia de accion de gracias
 cuenta = 0
 ano$ = Format(Val(Format(Now, "yyyy")), "0000")
 
 For t = 1 To 30
    fecha$ = Format("11/" + Format(t, "00") + "/" + ano$, "mm/dd/yyyy")
    dia_semana = Weekday(fecha$)
    If dia_semana = vbThursday Then
       cuenta = cuenta + 1
       If cuenta = 4 Then
          dia_thanks = t
          Exit For
       End If
    End If
 Next t
 
 
   
   
   cont = 0
   For t = 1 To Grid2.Rows - 1
      
      cont = (cont + 1)
      barra.Value = cont
      mensaje.Refresh
      openforms = DoEvents
      
      
      black_day = 0
      Grid2.Row = t
      
      Grid2.Col = 10
      firstname$ = Grid2.Text
      
      Grid2.Col = 11
      lasttname$ = Grid2.Text
      
      nombre$ = firstname$ + " " + lasttname$
      
      If UCase(nombre$) = UCase(lblemployee.Caption) And UCase(nombre$) <> "GABRIELA JIMENEZ" Then
        
      
      
      ElseIf UCase(nombre$) = UCase(lblemployee.Caption) And UCase(nombre$) = "GABRIELA JIMENEZ" Then
     
        black_day = 1
        
      ElseIf UCase(lblemployee.Caption) = "GABRIELA JIMENEZ" Then
      
         
      
      ElseIf UCase(lblemployee.Caption) = "BRENDA MARQUEZ" Then
        
      Else
         nombre$ = "RESERVED"
      End If
      
      
      Grid2.Col = 3
      ID_manager$ = Grid2.Text
      
      
      sSelect = "select emailwork from employeeinfo where idemployee='" + ID_manager$ + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
       ' Conecta_SQL
      End If
      correo_manager$ = Rs(0)
      Rs.Close
   
        
    
      
      Grid2.Col = 5
      fecha_solicitada$ = Format(Grid2.Text, "mm/dd/yyyy")
          
          
      If t < (Grid2.Rows - 1) Then
        Grid2.Row = t + 1
        fecha_siguiente$ = Format(Grid2.Text, "mm/dd/yyyy")
      Else
        fecha_siguiente$ = ""
      End If
            
      
      Grid2.Row = t
      
      Grid2.Col = 8
      nota$ = Grid2.Text
      
      Grid2.Col = 12
      UserName$ = Grid2.Text
      
      'If UCase(UserName$) = "AALVARADO" Then
      '   Stop
      'End If
      
      
      Grid2.Col = 13
      correo_agente$ = Grid2.Text
      
      Grid2.Col = 14
      cashout = Grid2.Text
      
            
      
      Grid2.Col = 1
      ID_vacaciones$ = Grid2.Text
      
      
      
      Grid2.Col = 6
      horas = Val(Grid2.Text)
      
      Grid2.Col = 7
      autorizado = Grid2.Text
      
      
      If fecha_solicitada$ = fecha_siguiente$ Then
           If autorizado = False Then
                PicColor.BackColor = vbamarillo
                colorx = vbamarillo
                GoTo saltado
           End If
      End If
      
      
      ' pone en otro color las fechas que ya pasaron al dia actual
      '*******************
      
      fecha_actual$ = Format(Now, "mm/dd/yyyy")
      dia_actual$ = Mid$(fecha_actual$, 4, 2)
      mes_actual$ = Left(fecha_actual$, 2)
      ano_actual$ = Right(fecha_actual$, 4)
      
      dia_solicitado$ = Mid$(fecha_solicitada$, 4, 2)
      mes_solicitado$ = Left(fecha_solicitada$, 2)
      ano_solicitado$ = Right(fecha_solicitada$, 4)
      
      If Val(ano_solicitado$) < Val(ano_actual$) Then
          PicColor.BackColor = &HFFC0C0
          colorx = &HFFC0C0
          GoTo saltado
      End If
      
      
      If Val(ano_solicitado$) = Val(ano_actual$) Then
          If Val(mes_solicitado$) < Val(mes_actual$) Then
             
            If UCase(lblemployee.Caption) = "GABRIELA JIMENEZ" Then
               If cashout = True Or cashout = 1 Then
                 PicColor.BackColor = rosa
                 colorx = rosa
               Else
                 PicColor.BackColor = &HFFC0C0
                 colorx = &HFFC0C0
               End If
            Else
               PicColor.BackColor = &HFFC0C0
               colorx = &HFFC0C0
            End If
             
             GoTo saltado
          End If
      End If
      
      
      
      If Val(ano_solicitado$) = Val(ano_actual$) Then
          If Val(mes_solicitado$) = Val(mes_actual$) Then
               If Val(dia_solicitado$) < Val(dia_actual$) Then
                   If UCase(lblemployee.Caption) = "GABRIELA JIMENEZ" Then
                       If cashout = True Or cashout = 1 Then
                           PicColor.BackColor = rosa
                           colorx = rosa
                       Else
                           PicColor.BackColor = &HFFC0C0
                           colorx = &HFFC0C0
                       End If
                   Else
                       PicColor.BackColor = &HFFC0C0
                       colorx = &HFFC0C0
                   End If
                    
                    GoTo saltado
               End If
          End If
      End If
      
      
      
            
      ' *******************
      
      
      If autorizado = True Then
              
         If black_day = 1 Then
           PicColor.BackColor = vbRed
           colorx = vbRed
           
         Else
              
           If horas = 8 Then
             If cashout = "False" Then
                PicColor.BackColor = vbverde
                colorx = vbverde
             Else
                PicColor.BackColor = rosa
                colorx = rosa
             End If
             
           Else
             If cashout = "0" Then
                PicColor.BackColor = vbverde_claro
                colorx = vbverde_claro
             Else
                PicColor.BackColor = rosa_claro
                colorx = rosa
             End If
           End If
           
         End If
            
            
      Else
           ' no esta autorizado aun
            PicColor.BackColor = vborange
            colorx = vborange
            
      End If
                 
                 
saltado:
                 
                             
                
                
                
       ' carga la oficina donde esta ubicado
    
    
      sSelect = "select Office from EmployeeInfo emp " & _
      "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
      "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
      "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
      "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
      "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
      "where emp.Active=1 and empofc.active=1 and Username='" + UserName$ + "'"
    
      Rs.Open sSelect, base, adOpenUnspecified
    
      grid3.AllowUserResizing = flexResizeColumns

      Set grid3.DataSource = Rs
                         
      Rs.Close
        
        
      ubicacion_de_trabajo$ = ""
      If grid3.Rows > 1 Then
          grid3.Row = 1
          grid3.Col = 1
          ubicacion_de_trabajo$ = grid3.Text
            
      End If
                   
      
      
      
                   
      With ucCalendar1
    
        
        

       ' .SetStrLanguage "Año", "Mes", "Semana", "Día", "Hoy", "Comienza", "Finaliza"

       Dim i As Long
       Dim StartDate As Date
       Dim EndDate As Date



      
         StartDate = Format(fecha_solicitada$, "mm/dd/yy") & " 00:00"
         EndDate = Format(fecha_solicitada$, "mm/dd/yy") & " 23:59"
    
         f1$ = StartDate & " 00:00:00 AM"
         f2$ = EndDate
        
         .Redraw = False
         
         
         'Optional Subject As String, _
         '       Optional StartDate As Date, _
         '       Optional EndDate As Date, _
         '       Optional Color As Long, _
         '       Optional AllDayEvent As Boolean, _
         '       Optional body As String, _
         '       Optional Tag As Variant, _
         '       Optional IsSerie As Boolean, _
         '       Optional NotifyIcon As Boolean, _
         '       Optional IsPrivate As Boolean, _
          '      Optional EventShowAs As eEventShowAs, _
          '      Optional office As String, _
          '      Optional idvacaciones As Integer, _
          '      Optional email1 As String, _
         '       Optional email2 As String)
       
         ' .AddEvents nombre$, f1$, f2$, PicColor.BackColor, ChkAllDay.Value, nota$, , ChkIsSerie.Value, ChkNotify.Value, ChkPrivate.Value, ImageCombo1.SelectedItem.Index - 1
         
         If carga_inicial = 0 Then
         
                     
                .AddEvents nombre$, StartDate, EndDate, colorx, , , , , , autorizado, , , , correo_agente$, correo_manager$
            
                

         Else
            Indice_del_evento = 0
           .GetEventsFromDay (StartDate)
           
           If Indice_del_evento <> 0 Then
            '  .UpdateEventData Indice_del_evento, nombre$, , , StartDate, EndDate, colorx, , , , , , , , correo_agente$, correo_manager$
           
           Else
              
              .AddEvents nombre$, StartDate, EndDate, colorx, , , , , , autorizado, , , , correo_agente$, correo_manager$
           End If
           
           
         End If
      
         .Redraw = True
         
       
      End With
    
   
    
                 
  
  Next t

  'agrega navidad, ano nuevo y accion de gracias
    
 If carga_inicial = 0 Then
    PicColor.BackColor = vbWhite
    colorx = vbWhite
    
    fecha_festiva$ = "12/25/" + Format(Now, "yyyy")
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "Christmas Day", StartDate, EndDate, colorx, , , , , , autorizado
    
    fecha_festiva$ = "01/01/" + Format(Now, "yyyy")
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "New Year's Day", StartDate, EndDate, colorx, , , , , , autorizado
    
    fecha_festiva$ = "11/" + Format(dia_thanks, "00") + "/" + Format(Now, "yyyy")
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "Thanksgiving Day", StartDate, EndDate, colorx, , , , , , autorizado
    
    
      ' calcular dia de accion de gracias del año que viene
    cuenta = 0
    ano$ = Format(Val(Format(Now, "yyyy")) + 1, "0000")
 
    For t = 1 To 30
      fecha$ = Format("11/" + Format(t, "00") + "/" + ano$, "mm/dd/yyyy")
      dia_semana = Weekday(fecha$)
      If dia_semana = vbThursday Then
       cuenta = cuenta + 1
       If cuenta = 4 Then
          dia_thanks2 = t
          Exit For
       End If
      End If
    Next t
 
 
    fecha_festiva$ = "12/25/" + ano$
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "Christmas Day", StartDate, EndDate, colorx, , , , , , autorizado
    
    fecha_festiva$ = "01/01/" + ano$
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "New Year's Day", StartDate, EndDate, colorx, , , , , , autorizado
    
    fecha_festiva$ = "11/" + Format(dia_thanks2, "00") + "/" + ano$
    StartDate = Format(fecha_festiva$, "mm/dd/yy") & " 00:00"
    EndDate = Format(fecha_festiva$, "mm/dd/yy") & " 23:59"
    ucCalendar1.AddEvents "Thanksgiving Day", StartDate, EndDate, colorx, , , , , , autorizado
    
 
 
    
 End If
   

  ' elimina fechas borradas del calendario
  ' =====================================================================================================================================
  ' ucCalendar1.Update
    
    a = ucCalendar1.GetAllEvents

  If carga_inicial = 1 Then
    
    tipo_de_carga$ = " and cast (lastupdated as date)='" + Format(Now, "mm/dd/yyyy") + "' "
    sSelect = "select * from VacationsProgram where Active=0 " + tipo_de_carga$ + " order by daterequested"
   
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      ' Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    List2.Clear
    For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       Grid2.Col = 5
       
       fecha_solicitada$ = Format(Grid2.Text, "mm/dd/yyyy")
       
        tipo_de_carga$ = " and cast (daterequested as date)='" + fecha_solicitada$ + "' "
       sSelect = "select * from VacationsProgram where Active=1 " + tipo_de_carga$ + " order by daterequested"
       Rs.Open sSelect, base, adOpenUnspecified
         ' Permitir redimensionar las columnas
       grid3.AllowUserResizing = flexResizeColumns
         ' Asignar el recordset al FlexGrid
       Set grid3.DataSource = Rs
       Rs.Close
        
       If grid3.Rows <= 1 Then
              
          existe = 0
          For z = 0 To List2.ListCount - 1
              If List2.List(z) = fecha_solicitada$ Then
                 existe = 1
                 Exit For
              End If
          Next z
          
          If existe = 1 Then GoTo saltalo
                 
          List2.AddItem fecha_solicitada$
                       
          existe = 0
          For Y = 0 To List1.ListCount - 1
             indicex$ = RTrim(LTrim(Right(List1.List(Y), Len(List1.List(Y)) - 10)))
             fechax$ = Left(List1.List(Y), 10)
          
             If fechax$ = fecha_solicitada$ Then
                existe = 1
                Exit For
             End If
          Next Y
       
       
          If existe = 1 Then
            With ucCalendar1
            
           .RemoveEvent indicex$
                
           End With
         End If
saltalo:
       End If
       
       
    Next t
    
   

  End If

  ' =====================================================================================================================================


   ucCalendar1.Update
   ucCalendar1.Visible = True
   
   btn_menu(0).Enabled = True
   carga_inicial = 1
   ' btn_menu(1).Enabled = True

  
    

End Sub

Public Sub verifica_disponibilidad()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
    
Set Rs = New ADODB.Recordset
   
    
   Dim StartDate As Date, EndDate As Date
   
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then
      Exit Sub
    End If
    
        fecha_solicitada$ = DateValue(StartDate)
        
            
     ano_solicitado$ = Format(fecha_solicitada$, "yyyy")
     mes_solicitado$ = Format(fecha_solicitada$, "mm")
     dia_solicitado$ = Format(fecha_solicitada$, "dd")
            
            
            
            
            
     fecha_hoy$ = Format(Now, "mm/dd/yyyy")
     
     ano_actual$ = Format(fecha_hoy$, "yyyy")
     mes_actual$ = Format(fecha_hoy$, "mm")
     dia_actual$ = Format(fecha_hoy$, "dd")
            
            
    If Val(ano_solicitado$) >= (Val(ano_actual$) + 2) Then
      MsgBox "You cannot select vacations in that year yet", 16, "Upss!"
      If dia_bloqueado <> 2 Then
         dia_bloqueado = 1
      End If
      Exit Sub
    End If
                 
            
            
     If Val(ano_solicitado$) < Val(ano_actual$) Then
        If dia_bloqueado <> 2 Then
         dia_bloqueado = 1
        End If
        Exit Sub
     ElseIf Val(ano_solicitado$) = Val(ano_actual$) Then
        If Val(mes_solicitado$) < Val(mes_actual$) Then
            If dia_bloqueado <> 2 Then
              dia_bloqueado = 1
            End If
            Exit Sub
        ElseIf Val(mes_solicitado$) = Val(mes_actual$) Then
             If Val(dia_solicitado$) < Val(dia_actual$) Then
                 If dia_bloqueado <> 2 Then
                    dia_bloqueado = 1
                 End If
                 Exit Sub
             End If
        End If
        
        
     End If
            
            
            
' -------------------------------------------------------------------------------------------------------
            
     ' revisa si no tiene vacaciones un mes antes y/o un mes despues
     
   dia_solicitado$ = Format(StartDate, "dd")
     
   mes_solicitado$ = Format(StartDate, "mm")    ' mes_solicitado$ = Format(ucCalendar1.DateValue, "mm")
   ano_solicitado$ = Format(StartDate, "yyyy")
    
    
   
   If Val(mes_solicitado$) > 1 Then
      mes_antes = Val(mes_solicitado$) - 1
      ano_antes = Val(ano_solicitado$)
   Else
      mes_antes = 12
      ano_antes = Val(ano_solicitado$) - 1
   End If
   
   
   If Val(mes_solicitado$) < 12 Then
      mes_despues = Val(mes_solicitado$) + 1
      ano_despues = Val(ano_solicitado$)
   Else
      mes_despues = 1
      ano_despues = Val(ano_solicitado$) + 1
   End If
   
      
   estado = 0
   
   
   If Val(dia_solicitado$) >= 28 Or Val(dia_solicitado$) <= 4 Then
      GoTo saltadox
   End If
   
   
   If Val(mes_solicitado$) = 1 Then
      mes_antes = 12
      ano_antes = Val(ano_solicitado$) - 1
      estado = 1
   ElseIf Val(mes_solicitado$) = 12 Then
      mes_despues = 1
      ano_despues = Val(ano_solicitado$) + 1
      estado = 1
   Else
      
   End If
   
   
   
   
   
   
   
  If estado = 0 Then
   

  sSelect = "select idvacation, emp.idemployee, office, ciarel.IdJobTitle, daterequested from VacationsProgram vac " & _
  "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel  depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and (month(daterequested)='" + Format(mes_antes, "00") + "' or month(daterequested)='" + Format(mes_despues, "00") + "') " & _
  "and year(daterequested)='" + ano_solicitado$ + "' " & _
  "and emp.IDEmployee = '" + lblemp_id.Caption + "' and vac.active='1' "
  
  ElseIf estado = 1 Then
  
  sSelect = "select idvacation, emp.idemployee, office, ciarel.IdJobTitle, daterequested from VacationsProgram vac " & _
  "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel  depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and (month(daterequested)='" + Format(mes_antes, "00") + "' or month(daterequested)='" + Format(mes_despues, "00") + "') " & _
  "and (year(daterequested)='" + Format(ano_antes, "0000") + "' or year(daterequested)='" + Format(ano_despues, "0000") + "') " & _
  "and emp.IDEmployee = '" + lblemp_id.Caption + "' and vac.active='1' "
  
  
  End If
  


  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close

     
     
    If grid3.Rows > 1 And permiso_doble_mes = 0 Then
    
     ' ***************  se agrego esto para permitir solamente una persona por dia de vacaciones ***************
      MsgBox "Sorry, but you already have a vacation assigned near this month", 16, "Upss!"
      valido1 = 6
      Exit Sub
     
    End If
     
     
     
     
     ' ---------------------------------------------------------------------------------------------------------------
     
     
saltadox:
            
    
   
    sSelect = "select idvacation, emp.idemployee from VacationsProgram vac " & _
    "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
    "where emp.Active=1 and daterequested='" + fecha_solicitada$ + "'  and vac.active='1'"
   
      
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
    
    
    
    
    Erase empleado_en_esta_fecha
    
    
    
    If grid3.Rows > 1 Then
    
     ' ***************  se agrego esto para permitir solamente una persona por dia de vacaciones ***************
      If chk_dia_doble.Value = 0 Then
        MsgBox "Sorry, but this day is already booked as a vacation", 16, "Upss!"
      
      valido1 = 5
      Exit Sub
      
      End If
     ' ******************************************************************************************************************
      
      
    
      existe = 0
    ' verifica agentes de ese dia
      For t = 1 To grid3.Rows - 1
         grid3.Row = t
         grid3.Col = 2
     
         empleado_en_esta_fecha(t) = Val(grid3.Text)
      
      
         If empleado_en_esta_fecha(t) = Val(lblemp_id.Caption) Then
            existe = 1
         End If
    
      Next t
    
    
      If existe = 1 Then
        valido1 = 5
      End If
    End If
    
    
    
    
    
    If grid3.Rows > 1 Then
      If grid3.Rows = 2 Then
         fecha_accesible = 1
      Else
         fecha_accesible = 2
      End If
         
    Else
       fecha_accesible = 0
       'Exit Sub
    End If
    
    
    
    ' verifica si son de la misma oficina
    
    
    

  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (17) and empjob.Active='1' and Username='" + user$ + "'"


  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close


    If grid3.Rows > 1 Then
       ES_MANAGER = 1
    Else
       ES_MANAGER = 0
    End If
       
   
   
   


    
    
    
    
    
     sSelect = "select idvacation, emp.idemployee, office, ciarel.IdJobTitle from VacationsProgram vac " & _
     "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
     "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
     "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
     "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
     "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
     "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
     "where emp.Active=1 and empofc.active=1 and daterequested='" + fecha_solicitada$ + "' and vac.active='1'"

    
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
    
     If grid3.Rows > 1 Then
      
       existe = 0
       For t = 1 To grid3.Rows - 1
         grid3.Row = t
         grid3.Col = 3
         oficina_emp$ = grid3.Text
         
         grid3.Col = 2
         id_del_employee$ = grid3.Text
         
         
         grid3.Col = 4
         titulo_empleado$ = grid3.Text
         
         
         
         
         If titulo_empleado$ = "17" And ES_MANAGER = 1 And id_del_employee$ <> lblemp_id.Caption Then
            MsgBox "We are sorry, but it is not possible to take this vacation day because another Manager has requested it", 16, "Upps!"
            valido1 = 5
            Exit For
         End If
         
         
         
         
         If oficina_emp$ = lbllocation.Caption And id_del_employee$ <> lblemp_id.Caption Then
            existe = 1
            Exit For
         End If
       Next t
       
       If existe = 1 Then
         MsgBox "We are sorry, but it is not possible to take this vacation day because another agent from the same office has requested it", 16, "Upps!"
         valido1 = 5
       End If
       
    
     End If
     
     
     
     
     
    
    
    
    
     
     
     
    
End Sub



Public Sub verifica_dias_seleccionados()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
    bloquea_acceso = 0
    current_year$ = Format(Now, "yyyy")
    
     Dim StartDate As Date, EndDate As Date
   
    If ucCalendar1.GetSelectionRangeDate(StartDate, EndDate) = False Then Exit Sub

    
    
    
    fecha_seleccionada$ = Format(StartDate, "mm/dd/yyyy")
    
     
    
   
    sSelect = "select idvacation, vac.idemployee, idmanager, approvedby, daterequested, hours, approved, notes, emp.payrolllinkid, emp.firstname, emp.lastname1, emp.username, emp.emailwork from VacationsProgram vac " & _
    "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
    "where emp.Active=1 and year(daterequested)='" + current_year$ + "' and vac.idemployee='" + lblemp_id.Caption + "'  and vac.active='1' order by daterequested, approved"
   
   
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
   Rs.Close
   
   
   
   dia_seleccionado = Format(fecha_seleccionada$, "y")
   
   
   lista.Clear
   lista2.Clear
   
   lista2.AddItem dia_seleccionado
   
   For t = 1 To grid3.Rows - 1
     grid3.Row = t
     grid3.Col = 5
     lista.AddItem Format(grid3.Text, "y")
   Next t
   
   
   conta = 0
   For t = 0 To lista.ListCount - 1
   
    For Y = 1 To 9
      If lista.List(t) = (dia_seleccionado - Y) Then
            lista2.AddItem lista.List(t)
            conta = conta + 1
      End If
    Next Y
          
          
    For Y = 1 To 9
      If lista.List(t) = (dia_seleccionado + Y) Then
            lista2.AddItem lista.List(t)
            conta = conta + 1
      End If
    Next Y
                  
            
      If conta > 5 Then
         Exit For
      End If
      
   Next t
   
   
   
   If lista2.ListCount > 5 Then
      bloquea_acceso = 1
   End If
   
   
   
   
   
End Sub


Public Sub envia_correo_aprobado()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
    
 'Exit Sub
    
    
Set Rs = New ADODB.Recordset
   
 
      
      '  +++++++++++++++++++++++++++++++++++++++++++++++++
      
      fuente_original$ = App.Path & "\"
      fuente$ = "c:\vacations\"
      
      If Dir$(fuente$ + "nueva.htm") = "" Then
        FileCopy fuente_original$ + "nueva.htm", fuente$ + "nueva.htm"
      End If
      
      ' FileCopy App.Path & "\config.ini", fuente$ + "config.ini"
      
      
      Name fuente$ + "nueva.htm" As fuente$ + "nueva2.htm"

      nf2 = FreeFile
      Open fuente$ + "nueva.htm" For Output Shared As #nf2
      
      
      
      a$ = ID_vacaciones$
      
      
      
     sSelect = "select * from vacationsprogram where idvacation='" + ID_vacaciones$ + "' and active='1'"
   
      
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
    
    If grid3.Rows > 1 Then
       
       grid3.Row = 2
       grid3.Col = 2
       id_employee$ = grid3.Text
       
         grid3.Col = 8
         cashout$ = grid3.Text
         
       
       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + id_employee$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre_empleado$ = Rs(0)
       apellido_empleado$ = Rs(1)
       Rs.Close
       nombre_empleado$ = nombre_empleado$ + " " + apellido_empleado$
              
       
       grid3.Col = 3
       ID_manager$ = grid3.Text
       
       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + ID_manager$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre_manager$ = Rs(0)
       apellido_manager$ = Rs(1)
       Rs.Close
       nombre_manager$ = nombre_manager$ + " " + apellido_manager$
       
       
       grid3.Col = 5
       fecha$ = Format(grid3.Text, "mm/dd/yyyy")
       
       grid3.Col = 6
       horas$ = grid3.Text
       
       grid3.Col = 7
       Status_aprobado$ = grid3.Text
       
       If Status_aprobado$ = "1" Or Status_aprobado$ = True Then
           Status_aprobado$ = "Approved"
       Else
          Status_aprobado$ = "Pending"
       End If
       
       
       
       
    Else
        Exit Sub
    End If
    
    
        
        
      


      msg2$ = "</p><style type=" + Chr$(34) + "text/css" + Chr$(34) + "> <!--.Estilo1 {font-family: " + Chr$(34) + "Courier new" + Chr$(34) + "}--></style><span class=" + Chr$(34) + "Estilo1" + Chr$(34) + ">"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
    

      If cashout$ = "False" Then
          msg2$ = "</p><img src=" + Chr$(34) + "c:\vacations\imagen.jpg" + Chr$(34) + " width=" + Chr$(34) + "450" + Chr$(34) + " height=" + Chr$(34) + "280" + Chr$(34) + "></a></p>"
      Else
          msg2$ = "</p><img src=" + Chr$(34) + "c:\vacations\imagen2.jpg" + Chr$(34) + " width=" + Chr$(34) + "450" + Chr$(34) + " height=" + Chr$(34) + "280" + Chr$(34) + "></a></p>"
      End If
          
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2



      msg2$ = "</p><h1></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      

      msg2$ = "<font color=" + Chr$(34) + "blue" + Chr$(34) + "><p> Congratulations! </font><b> </b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
    
    
     
    
     
    
    
      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> your vacations has been approved </font></b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      

      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
    




      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
                    
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Requested date: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + fecha$ + "</b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Requested hours: </font><font color=" + Chr$(34) + "Blue" + Chr$(34) + "><b>" + horas$ + "</b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
     
      
      msg2$ = "</p> -------------------------"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p><h2></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p> Status: </font><font color=" + Chr$(34) + "red" + Chr$(34) + "><b> Approved </b></p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
            
      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      msg2$ = "</p> -------------------------"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
       
      
      If UCase(nombre_empleado$) = UCase(nombre_manager$) Then
      
        msg2$ = "</p> Agent/Manager: " + nombre_empleado$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
        
      Else
      
        msg2$ = "</p> Agent: " + nombre_empleado$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
      
        msg2$ = "</p> Manager: " + nombre_manager$ + " </p>"
        Lock #nf2
        Print #nf2, msg2$
        Unlock #nf2
        
      End If
      
      
      
            
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      GoTo salta_esto
            
      ' ------------------------------
      
       msg2$ = "<font color=" + Chr$(34) + "black" + Chr$(34) + "><p>I M P O R T A N T : </p></font>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      ultimo_mensaje$ = txtcomment.Text
      
      
      
      
      ' separa mensaje en lineas
      conta_caracteres = 0
      R$ = ""
      For t = 1 To Len(ultimo_mensaje$)
        R$ = R$ + Mid$(ultimo_mensaje$, t, 1)
        conta_caracteres = conta_caracteres + 1
        If conta_caracteres >= 28 And Mid$(ultimo_mensaje$, t, 1) = Space(1) Then
             msg2$ = "</p><u><b><i>" + R$ + "</i></b></u></p>"
             Lock #nf2
             Print #nf2, msg2$
             Unlock #nf2
             
             R$ = ""
             conta_caracteres = 0
         End If
      Next t
      
      If conta_caracteres > 0 Then
          msg2$ = "</p><u><b><i>" + R$ + "</i></b></u></p>"
          Lock #nf2
          Print #nf2, msg2$
          Unlock #nf2
      End If
      
       
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      ' ---------------------------------------------------
      
            
salta_esto:
            
            
      msg2$ = "</p><h5></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      'msg2$ = "</p>&nbsp;</p>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
            
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
      msg2$ = "<p><a href='https://secure5.yourpayrollhr.com/ta/JAI04.login' ><img src=" + Chr$(34) + "c:\vacations\important.jpg" + Chr$(34) + " width=" + Chr$(34) + "400" + Chr$(34) + " height=" + Chr$(34) + "150" + Chr$(34) + "></a></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
            
     
            
      'msg2$ = "<a href='https://secure5.yourpayrollhr.com/ta/JAI04.login' ><p> " + Space(1) + " </p></a>"
      'Lock #nf2
      'Print #nf2, msg2$
      'Unlock #nf2
            
            
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
            
         
            
            
      msg2$ = "</p><h3></p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
            
      msg2$ = "</p> >> PLEASE, CHECK VACATION PROGRAM <<" + Space(1) + " </p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      msg2$ = "</p>&nbsp;</p>"
      Lock #nf2
      Print #nf2, msg2$
      Unlock #nf2
      
      
       
       Close nf2

      G = valido1
      
      
      
      
      If UCase(nombre_empleado$) = UCase(nombre_manager$) Then
      
          transfiere$ = correo_agente$ + "; it@justautoins.com"
          
      Else
          transfiere$ = correo_agente$ + "; " + correo_manager$ + "; it@justautoins.com"
      
      End If
      
      
      
      'transfiere$ = "hnavarro@justautoins.com"
     
      send_email
      
End Sub

Public Sub actualiza_fechas()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
   
    sSelect = "select count(*) from vacationsprogram where active=1"
       Rs.Open sSelect, base, adOpenUnspecified
       registros = Val(Rs(0))
       Rs.Close
   
   
       If List1.ListCount <> (registros + 6) Then
          Carga_fechas
       End If
   
End Sub

Public Sub verifica_numero_de_managers()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
    
Set Rs = New ADODB.Recordset


  If valido1 = 6 Then
    valido1 = 5
    Exit Sub
  End If
  
  
  If user$ = "GJIMENEZ" Then
     id_emp$ = RTrim(LTrim(Right(cbo_users.List(cbo_users.ListIndex), 5)))
      sSelect = "select username from employeeinfo where idemployee='" + id_emp$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
         
      R$ = UCase(RTrim(LTrim(Rs(0))))
      Rs.Close
    
  Else
     id_emp$ = "76"
     R$ = user$
  End If

  ' verifica si es manager

  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (17) and empjob.Active='1' and Username='" + R$ + "'"


  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close


    If grid3.Rows > 1 Then
       ES_MANAGER = 1
    Else
       ES_MANAGER = 0
    End If
       
   
   If ES_MANAGER = 0 Then
      Exit Sub
   End If
   

   lista.Clear
    
   mes_solicitado$ = Format(ucCalendar1.DateValue, "mm")
   ano_solicitado$ = Format(ucCalendar1.DateValue, "yyyy")
    
    ' verifica total de managers en el mes

  sSelect = "select idvacation, emp.idemployee, office, ciarel.IdJobTitle, daterequested from VacationsProgram vac " & _
  "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel  depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (17) and month(daterequested)='" + mes_solicitado$ + "' and year(daterequested)='" + ano_solicitado$ + "' " & _
  "and emp.IDEmployee <> '" + lblemp_id.Caption + "'  and vac.active='1'"



  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
          
    
    
    If grid3.Rows > 1 Then
       lista.Clear
       lista.AddItem Val(id_emp$)
       
       For t = 1 To grid3.Rows - 1
         grid3.Row = t
         grid3.Col = 2
         ID_manager = Val(grid3.Text)
         If ID_manager = 76 Then GoTo saltalo
         If Val(id_emp$) = ID_manager Then
           GoTo saltalo
         End If
         
         existe = 0
         For Y = 0 To lista.ListCount - 1
            If Val(lista.List(Y)) = ID_manager Then
               existe = 1
               Exit For
            End If
         Next Y
         
         If existe = 0 Then
            lista.AddItem ID_manager
            
         End If
         
saltalo:
       
       Next t
       
       
       
    End If
    
    
    
    If lista.ListCount > 1 Then
        MsgBox "Sorry, but it is not possible to take this day on vacation because during this month, a manager has already booked", 16, "Upps!"
        valido1 = 5
        Exit Sub
        
    End If
    
    
    
    
  
     
     
    
    
    

End Sub

Public Sub send_email()
On Error Resume Next

Const sch = "http://schemas.microsoft.com/cdo/configuration/"

Dim INI_PATH As String
Dim loCfg As Object
Dim loMsg As Object
Dim loBP As Object
Dim i As Long
Dim DestImg As String
Dim TempHTML As String
Dim TempHTMLMail As String
Dim strImg As String


Dim objMessage, objConfig, Fields

Set mcIni = New clsIni

fuente_original$ = App.Path & "\"
fuente$ = "c:\vacations\"


WebBrowser1.RegisterAsDropTarget = False


WebBrowser1.navigate "c:\vacations\Nueva.htm"

Do: DoEvents: Loop Until WebBrowser1.readyState = READYSTATE_COMPLETE
WebBrowser1.document.designMode = "On"
Set HTML = WebBrowser1.document


INI_PATH = "c:\vacations\config.ini"

DoEvents





Set objMessage = CreateObject("CDO.Message")

Set objConfig = CreateObject("CDO.Configuration")
Set Fields = objConfig.Fields

With Fields
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mcIni.getValue(INI_PATH, "datos", "servidor")    '"smtp.office365.com"
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mcIni.getValue(INI_PATH, "datos", "puerto")  '25
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = mcIni.getValue(INI_PATH, "datos", "Aut", 0)  '1
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mcIni.getValue(INI_PATH, "datos", "usuario")   '"Vacations@justautoins.com"
  .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mcIni.Encriptar(App.EXEName, mcIni.getValue(INI_PATH, "datos", "password"), 2)
  .Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = True
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = mcIni.getValue(INI_PATH, "datos", "ssl", 0)   'True
  .Update
End With



 

Set objMessage.Configuration = objConfig

With objMessage
  .Subject = "Vacations status"
  .From = "Vacations@justautoins.com"
  .To = transfiere$
  
  
  ' carga todo el cuerpo del WEB HTML
  
   TempHTML = HTML.documentElement.outerHTML
  
  If HTML.All.tags("BODY")(0).Background <> "" Then
  
        strImg = HTML.All.tags("BODY")(0).Background
        
        If PathIsURL(strImg) <> 1 Then
            HTML.All.tags("BODY")(0).Background = "cid:" & "BackGround"
            
            Set loBP = .AddRelatedBodyPart(strImg, "BackGround", 1)
            
            With loBP.Fields
                .Item("urn:schemas:mailheader:Content-ID") = "BackGround"
                .Update
            End With
        End If
        
  End If
  
  
  For i = 0 To HTML.images.Length - 1
  
        strImg = HTML.images.Item(i).src
        
        If Left(strImg, 8) = "file:///" Then
            
            DestImg = GetFileNameURL(strImg)
            If DestImg <> "" Then
    
                HTML.images.Item(i).src = "cid:" & DestImg & i
    
                Set loBP = .AddRelatedBodyPart(strImg, DestImg & i, 1)
                
                With loBP.Fields
                    .Item("urn:schemas:mailheader:Content-ID") = DestImg & i
                    .Update
                End With
                
            End If
        End If
    Next

    ' carga el cuerpo de la forma html
    .HTMLBody = HTML.documentElement.outerHTML

    HTML.body.innerHTML = TempHTML
    
    
    
    
  
   
 
 '  .AutoGenerateTextBody = False
  ' .AddRelatedBodyPart "c:\vacations\imagen.jpg", "imagen.jpg", cdoRefTypeId  'cdoRefTypeLocation ' Si aca le pones el otro valor posible (cdoRefTypeId),entonces se adjunta la imagen
  
                            
    '.Fields(CdoMailHeader.cdoDispositionNotificationTo).Value = "it@justautoins.com"   ' recibir notificacion de leido
  
  
   'Prioridad
     If Check2.Value = False Or Check2.Value = 0 Then
         mPrioridad = 0
     Else
         mPrioridad = 1
     End If
     
    ' -1=Low, 0=Normal, 1=High
     
     
    .Fields("urn:schemas:httpmail:priority") = mPrioridad
    .Fields("urn:schemas:mailheader:X-Priority") = mPrioridad
    'Importancia
    
    
    '0=Low, 1=Normal, 2=High
    .Fields("urn:schemas:httpmail:importance") = mPrioridad + 1
    
    
    
    
      'Solicitar confirmación de lectura
     If Check1.Value = False Or Check1.Value = 0 Then
         
     Else
         .Fields("urn:schemas:mailheader:disposition-notification-to") = "it@justautoins.com"  '.From
         .Fields("urn:schemas:mailheader:return-receipt-to") = "it@justautoins.com"   '.From
    
     End If
     
     
  
  
  .Fields.Update
        
    ' anexa un archivo al correo
 ' .AddAttachment "C:\vacations\version.txt"
  
End With
objMessage.Send




'webBrowser1.Close


Kill fuente$ + "nueva.htm"
Name fuente$ + "nueva2.htm" As fuente$ + "nueva.htm"
      

MsgBox "The message was sent correctly", 64, "Attention"
transfiere$ = ""

'WebBrowser1.navigate ("about:blank")

' borra el navegador y permite iniciar uno nuevo sin el mensaje de grabado
WebBrowser1.Refresh

End Sub

Public Sub carga_empleados()
On Error Resume Next

' agrega los usuarios al combo
If valido1 = 999 Then
  Exit Sub
End If
 
 
 Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   
 Grid1.Clear
 
 
    cbo_users.Clear
         
     'oficina$ = ""
   'If cbo_oficina.ListIndex >= 0 Then
    ' oficina$ = UCase(LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 25))))
   'End If
         
         
         
         
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle, emp.emailwork, emp.firstname, emp.lastname1 from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,6) "
   
  ' 1 = IT manager
  ' 16 = Agente regular
  ' 17 = Manager
  ' 2 = Phone Sales
  ' 18 = Underwriting
  ' 24 = QC Monterrey
  ' 28 = Customer service Monterrey
  
      
      
   
 ' ---------------------------------------------------------------------------
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Grid1.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Grid1.Cols = Rs.Fields.Count + 1
    
    
    
   Grid1.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Grid1.TextMatrix(0, i + 1) = Rs.Fields(i).Name
   Next

   Grid1.Row = 1
   Grid1.Col = 1

   ' Set range of cells in the grid
   Grid1.RowSel = Grid1.Rows - 1
   Grid1.ColSel = Grid1.Cols - 1
   Grid1.clip = rsVar

   ' Reset the grid's selected range of cells
   Grid1.RowSel = Grid1.Row
   Grid1.ColSel = Grid1.Col

   Rs.Close

   '
   

' ----------------------------------------------------------------------------

    cbo_users.Clear
         
         
    For t = 1 To Grid1.Rows - 1
       Grid1.Row = t
       
       Grid1.Col = 1
       id_agente$ = Grid1.Text
           
       Grid1.Col = 2
       usuario$ = Grid1.Text
       
       Grid1.Col = 3
       oficina_agente$ = Grid1.Text
              
       Grid1.Col = 6
       nombre$ = UCase(Grid1.Text)
       
       Grid1.Col = 7
       apellido$ = UCase(Grid1.Text)
       
       
       name1$ = nombre$ + Space(1) + apellido$
       
       existe = 0
       For Y = 0 To cbo_users.ListCount - 1
          If UCase(RTrim(Left(cbo_users.List(Y), 30))) = UCase(name1$) Then
             existe = 1
             Exit For
          End If
       Next Y
       
       
       
       If existe = 0 And oficina_agente$ <> "JA - PHONE SALES" And name1$ <> "GABRIELA JIMENEZ" Then
         If letra$ = "-" Then
            cbo_users.AddItem Format(name1$, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(10) + Format(oficina_agente$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(5) + id_agente$
         Else
            If UCase(Left(name1$, 1)) = letra$ Then
                cbo_users.AddItem Format(name1$, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(10) + Format(oficina_agente$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(5) + id_agente$
            End If
         End If
       End If
       
    Next t
              
       
       
    
    
    
    

End Sub


Public Sub pantalla_transparente()
        ' Aplicar el efecto
        Dim tAlpha As Long
        
        tAlpha = Val(txtAlpha)
        If tAlpha < 1 Or tAlpha > 100 Then
            tAlpha = 30  ' 70
        End If
        
        '// Set WS_EX_LAYERED on this window
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        
        '// Make this window tAlpha% alpha
        Call SetLayeredWindowAttributes(hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
        'Image3.Visible = False
    
 
End Sub

Public Sub pantalla_solida()
                              ' Quitar el efecto
        '// Remove WS_EX_LAYERED from this window styles
        Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not WS_EX_LAYERED)
        
        '// Ask the window and its children to repaint
        Call RedrawWindow2(hwnd, 0&, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN)
        'Image3.Visible = True

End Sub

Public Sub Actualiza_horas()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset


fecha_hoy$ = Format(Now, "mm/dd/yyyy")

dia_hoy$ = Mid$(fecha_hoy$, 4, 2)
mes_hoy$ = Left(fecha_hoy$, 2)
ano_hoy$ = Right(fecha_hoy$, 4)


dia_aniversario$ = Mid$(lblhired_date.Caption, 4, 2)
mes_aniversario$ = Left(lblhired_date.Caption, 2)
ano_aniversario$ = Right(lblhired_date.Caption, 4)



If Val(ano_hoy$) = Val(ano_aniversario$) Then
   Exit Sub
ElseIf Val(ano_hoy$) > Val(ano_aniversario$) Then
  ' significa que ya cumplio +de 1 aniversario
    If Val(mes_hoy$) >= Val(mes_aniversario$) Then
          If Val(dia_hoy$) >= Val(dia_aniversario$) Then
            
          End If
    
    End If

End If





End Sub

Public Sub carga_GI_anuales()
On Error Resume Next

 Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   ' ano_actual = Val(Format(Now, "yyyy"))
   ano_actual = ano_para_GI
   f1$ = "01/01/" + Format(ano_actual, "0000")
   f2$ = "12/31/" + Format(ano_actual, "0000")
   
   
   Grid1.Clear
   Grid2.Clear
   
   
   Lista_membresia(0).Clear
   Lista_membresia(1).Clear
   Lista_membresia(2).Clear
   
   
   sSelect = "SELECT [IdEmployee] ,sum ([GI]) as GI FROM [LAESystemJA].[dbo].[EmplGoalsCalc] goals " & _
   "inner join ReceiptsHDR rechdr on rechdr.IDReceiptHDR= goals.IdReceiptHDR " & _
   "inner join InvoiceItemCatalog ii on ii.IdInvoiceItem=goals.IdInvoiceItem " & _
   "where  cast(rechdr.date as Date) >= '" + f1$ + "' AND cast( rechdr.date as Date) <= '" + f2$ + "' " & _
   "and rechdr.Active=1  group by goals.IdEmployee  order by  goals.IdEmployee"
   
   
   
      ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Grid1.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Grid1.Cols = Rs.Fields.Count + 1
    
    
    
   Grid1.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Grid1.TextMatrix(0, i + 1) = Rs.Fields(i).Name
   Next

   Grid1.Row = 1
   Grid1.Col = 1

   ' Set range of cells in the grid
   Grid1.RowSel = Grid1.Rows - 1
   Grid1.ColSel = Grid1.Cols - 1
   Grid1.clip = rsVar

   ' Reset the grid's selected range of cells
   Grid1.RowSel = Grid1.Row
   Grid1.ColSel = Grid1.Col

   Rs.Close



   ' carga los idJOBtitles
   
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
   "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
   "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
   "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
   "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
   "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
   "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,6) and ofc.Office <> 'JA - PHONE SALES' and ofc.Office <> 'JA - MONTERREY'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close


 ' carga los titutos del TIER
    sSelect = "select annualGI from vacationstierscatalog where idvacationtier=1"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_1$ = Rs(0)
    Rs.Close
    lblrango(0).Caption = Format(annual_GI_1$, "$###,##0.00") + " -"
    
    sSelect = "select annualGI from vacationstierscatalog where idvacationtier=2"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_2$ = Rs(0)
    Rs.Close
    lblrango(1).Caption = Format(annual_GI_2$, "$###,##0.00") + " - " + Format(Val(annual_GI_1$) - 1, "$###,##0.00")
    
    sSelect = "select annualGI from vacationstierscatalog where idvacationtier=3"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_3$ = Rs(0)
    Rs.Close
    lblrango(2).Caption = Format(annual_GI_3$, "$###,##0.00") + " - " + Format(Val(annual_GI_2$) - 1, "$###,##0.00")
    
    
    



   ' separa nombres de agentes en base a su anual GI
   
   For t = 1 To Grid1.Rows - 1
      Grid1.Row = t
      Grid1.Col = 1
      id_empleado$ = Grid1.Text
      
      Grid1.Col = 2
      GI$ = Grid1.Text
      
      
      existe = 0
      For Y = 0 To Grid2.Rows - 1
         Grid2.Row = Y
         Grid2.Col = 1
         id_empl$ = Grid2.Text
          
         If id_empleado$ = id_empl$ Then
                          
            sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + id_empleado$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            nombre$ = Rs(0)
            apellido$ = Rs(1)
            Rs.Close
            
            If UCase(user$) <> "GJIMENEZ" Then
            
              If Val(GI$) >= Val(annual_GI_1$) Then
                Lista_membresia(0).AddItem nombre$ + " " + apellido$
              ElseIf Val(GI$) >= Val(annual_GI_2$) And Val(GI$) < (Val(annual_GI_1$)) Then
                Lista_membresia(1).AddItem nombre$ + " " + apellido$
              ElseIf Val(GI$) < Val(annual_GI_2$) Then
                Lista_membresia(2).AddItem nombre$ + " " + apellido$
              End If
              
            Else
            
              If Val(GI$) >= Val(annual_GI_1$) Then
                Lista_membresia(0).AddItem Format(nombre$ + " " + apellido$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + " " + Format(Format(GI$, "$###,##0.00"), "@@@@@@@@@@@")
              ElseIf Val(GI$) >= Val(annual_GI_2$) And Val(GI$) < (Val(annual_GI_1$)) Then
                Lista_membresia(1).AddItem Format(nombre$ + " " + apellido$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + " " + Format(Format(GI$, "$###,##0.00"), "@@@@@@@@@@@")
              ElseIf Val(GI$) < Val(annual_GI_2$) Then
                Lista_membresia(2).AddItem Format(nombre$ + " " + apellido$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + " " + Format(Format(GI$, "$###,##0.00"), "@@@@@@@@@@@")
              End If
              
            
            End If
            
            existe = 1
            Exit For
         End If
         
      Next Y
      
      
    Next t
      
      
   ' ordena las listas x cantidad
   
   If UCase(user$) = "GJIMENEZ" Then
   
    For t = 0 To 2
      lista.Clear
      For Y = 0 To Lista_membresia(t).ListCount - 1
           nombre$ = Left(Lista_membresia(t).List(Y), Len(Lista_membresia(t).List(Y)) - 12)
           cantidad$ = Right(Lista_membresia(t).List(Y), 11)
           lista.AddItem cantidad$ + " " + nombre$
      Next Y
      
      Lista_membresia(t).Clear
      For Y = lista.ListCount - 1 To 0 Step -1
           nombre$ = Right(lista.List(Y), Len(lista.List(Y)) - 12)
           cantidad$ = Left(lista.List(Y), 11)
          Lista_membresia(t).AddItem nombre$ + " " + cantidad$
      Next Y
      
    Next t
    
   Else
     For t = 0 To 2
       lista.Clear
       For Y = 0 To Lista_membresia(t).ListCount - 1
         lista.AddItem Lista_membresia(t).List(Y)
       Next Y
       Lista_membresia(t).Clear
       For Y = 0 To lista.ListCount - 1
          Lista_membresia(t).AddItem lista.List(Y)
       Next Y
     Next t
     
   End If

End Sub

Public Sub carga_titulos_de_tiers()

   

    
End Sub

Public Sub carga_accesos()
On Error Resume Next
nf = FreeFile
lista_acceso.Clear
Open "\\192.168.84.215\vacations\Lista_acceso" For Input Shared As #nf
Lock #nf
i = 0
Do Until EOF(nf)
   Line Input #nf, N$
   Line Input #nf, R$
   If N$ <> "" Then
      lista_acceso.AddItem N$
   End If
   If R$ = "True" Then
      lista_acceso.Selected(i) = True
   Else
     lista_acceso.Selected(i) = False
   End If
   i = i + 1
Loop
Unlock #nf
Close #nf
End Sub
