VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Allowed access"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H cmddown 
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   2400
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   661
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
      Image           =   "Form5_lista_vacations.frx":0000
      ImgSize         =   40
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H cmdup 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   1920
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   661
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
      Image           =   "Form5_lista_vacations.frx":0CB2
      ImgSize         =   40
      cBack           =   8421504
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   6240
      ScaleHeight     =   2865
      ScaleWidth      =   465
      TabIndex        =   10
      Top             =   1080
      Width           =   495
   End
   Begin Project1.lvButtons_H btnagregar 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "Add"
      CapAlign        =   2
      BackStyle       =   2
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   2
      Image           =   "Form5_lista_vacations.frx":196B
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000A&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   3735
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   360
      Width           =   2055
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
         ItemData        =   "Form5_lista_vacations.frx":1DBD
         Left            =   240
         List            =   "Form5_lista_vacations.frx":1DBF
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin Project1.lvButtons_H btneliminar_empleado 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   825
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Delete"
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
         Image           =   "Form5_lista_vacations.frx":1DC1
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   3615
         Left            =   0
         Picture         =   "Form5_lista_vacations.frx":2283
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.ListBox List1 
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
      Height          =   3000
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin Project1.lvButtons_H CmdCancel 
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   3960
      Width           =   495
      _ExtentX        =   873
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
      Image           =   "Form5_lista_vacations.frx":5CD9
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2175
      Left            =   7800
      TabIndex        =   6
      Top             =   840
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
   Begin Project1.lvButtons_H btnletra 
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   11
      Top             =   840
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
   Begin Project1.lvButtons_H btnletra 
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   12
      Top             =   1080
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1320
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
      Left            =   2880
      TabIndex        =   14
      Top             =   1560
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
      Left            =   2880
      TabIndex        =   15
      Top             =   1800
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
      Left            =   2880
      TabIndex        =   16
      Top             =   2040
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
      Left            =   2880
      TabIndex        =   17
      Top             =   2280
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
      Left            =   2880
      TabIndex        =   18
      Top             =   2520
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
      Left            =   2880
      TabIndex        =   19
      Top             =   2760
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
      Left            =   2880
      TabIndex        =   20
      Top             =   3000
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
      Left            =   2880
      TabIndex        =   21
      Top             =   3240
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
      Left            =   2880
      TabIndex        =   22
      Top             =   3480
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
      Left            =   2880
      TabIndex        =   23
      Top             =   3720
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
      Left            =   3120
      TabIndex        =   24
      Top             =   840
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
      Left            =   3120
      TabIndex        =   25
      Top             =   1080
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
      Left            =   3120
      TabIndex        =   26
      Top             =   1320
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
      Left            =   3120
      TabIndex        =   27
      Top             =   1560
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
      Left            =   3120
      TabIndex        =   28
      Top             =   1800
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
      Left            =   3120
      TabIndex        =   29
      Top             =   2040
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
      Left            =   3120
      TabIndex        =   30
      Top             =   2280
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
      Left            =   3120
      TabIndex        =   31
      Top             =   2520
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
      Left            =   3120
      TabIndex        =   32
      Top             =   2760
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
      Left            =   3120
      TabIndex        =   33
      Top             =   3000
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
      Left            =   3120
      TabIndex        =   34
      Top             =   3240
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
      Left            =   3120
      TabIndex        =   35
      Top             =   3480
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
      Left            =   3120
      TabIndex        =   36
      Top             =   3720
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
      Height          =   375
      Index           =   26
      Left            =   2880
      TabIndex        =   37
      Top             =   3960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   5160
      TabIndex        =   39
      Top             =   4080
      Width           =   90
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   4680
      TabIndex        =   38
      Top             =   4080
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim id_emp$, letra$
Private Sub btnagregar_Click()
On Error Resume Next
Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   If List1.ListIndex = -1 Then
      MsgBox "Select the employee name from the list", 64, "Attention"
      Exit Sub
   End If
   
   
      sSelect = "select username from employeeinfo where idemployee='" + id_emp$ + "'"
 
      Rs.Open sSelect, base, adOpenUnspecified
     
      UserName$ = RTrim(LTrim(Rs(0)))
      Rs.Close

  ' verifica si existe en la lista
  existe = 0
  For t = 0 To lista_acceso.ListCount - 1
     If UCase(lista_acceso.List(t)) = UCase(UserName$) Then
         existe = 1
         Exit For
     End If
  Next t
  
  If existe = 1 Then
      MsgBox "The username already exists in the list", 64, "Attention"
      Exit Sub
  End If

  lista_acceso.AddItem UCase(UserName$)
  lista_acceso_Click

  List1.Selected(List1.ListIndex) = False
  lbltotal.Caption = Format(lista_acceso.ListCount, "###,##0")
      
End Sub

Private Sub btneliminar_empleado_Click()
On Error Resume Next

seleccionados = 0
For t = 0 To lista_acceso.ListCount - 1
  If lista_acceso.Selected(t) = True Then
      seleccionados = 1
      Exit For
  End If
Next t

If seleccionados = 0 Then
  MsgBox "You have not selected any name from the list", 16, "Attention"
  Exit Sub
End If


R$ = MsgBox("Do you want to delete the selected names?", 4, "Attention")
If R$ = "7" Then Exit Sub

valido1 = 888


nf = FreeFile

Open "\\192.168.84.215\vacations\Lista_acceso" For Output Shared As #nf
Lock #nf

For t = 0 To lista_acceso.ListCount - 1
  If lista_acceso.Selected(t) = False Then
   Print #nf, lista_acceso.List(t)
   Print #nf, lista_acceso.Selected(t)
  End If
Next t

Unlock #nf
Close #nf

valido1 = 0
carga_accesos


End Sub


Private Sub btnletra_Click(Index As Integer)
On Error Resume Next
If Index = 26 Then
  letra$ = "-"
Else
  letra$ = btnletra(Index).Caption
End If

carga_empleados

End Sub

Private Sub CmdCancel_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub cmddown_Click()
Dim currIndex As Integer
    Dim currItem  As String
    Dim bSel As Boolean
    
    If lista_acceso.ListIndex <> -1 And lista_acceso.ListIndex < lista_acceso.ListCount - 1 Then
        currIndex = lista_acceso.ListIndex
        currItem = lista_acceso.List(currIndex)
        bSel = lista_acceso.Selected(currIndex)
        lista_acceso.RemoveItem currIndex
        lista_acceso.AddItem currItem, currIndex + 1
        lista_acceso.ListIndex = currIndex + 1
        lista_acceso.Selected(currIndex + 1) = bSel
    End If
End Sub

Private Sub cmdup_Click()
On Error Resume Next
Dim currIndex As Integer
    Dim currItem  As String
    Dim bSel As Boolean
 
    If lista_acceso.ListIndex >= 1 Then
        currIndex = lista_acceso.ListIndex
        currItem = lista_acceso.List(currIndex)
        bSel = lista_acceso.Selected(currIndex)
        lista_acceso.RemoveItem currIndex
        lista_acceso.AddItem currItem, currIndex - 1
        lista_acceso.ListIndex = currIndex - 1
        lista_acceso.Selected(currIndex - 1) = bSel
    End If
    
    
End Sub

Private Sub Form_Load()
On Error Resume Next
top = 0
Left = (Screen.Width - Width) / 2

Conecta_SQL

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


letra$ = "-"
carga_empleados
carga_accesos

End Sub

Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
 
 
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com"

 
 

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
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
 
 
    List1.Clear
         
         
         
         
         
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

    List1.Clear
         
         
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
       For Y = 0 To List1.ListCount - 1
          If UCase(RTrim(Left(List1.List(Y), 30))) = UCase(name1$) Then
             existe = 1
             Exit For
          End If
       Next Y
       
       
       
       If existe = 0 And oficina_agente$ <> "JA - PHONE SALES" And name1$ <> "GABRIELA JIMENEZ" Then
          If letra$ = "-" Then
            List1.AddItem Format(name1$, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(10) + Format(oficina_agente$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(5) + id_agente$
          Else
             If UCase(Left(name1$, 1)) = letra$ Then
                List1.AddItem Format(name1$, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(10) + Format(oficina_agente$, "!@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(5) + id_agente$
             End If
          End If
       End If
       
    Next t
              
       
       
  
       
        lbltotal.Caption = Format(lista_acceso.ListCount, "###,##0")
       
         
    
    
    
    

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

  lbltotal.Caption = Format(lista_acceso.ListCount, "###,##0")
      
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
End Sub

Private Sub List1_Click()
On Error Resume Next
id_emp$ = LTrim(RTrim(Right(List1.List(List1.ListIndex), 6)))



End Sub

Private Sub lista_acceso_Click()
On Error Resume Next
If valido1 = 888 Then
  Exit Sub
End If

nf = FreeFile

Open "\\192.168.84.215\vacations\Lista_acceso" For Output Shared As #nf
Lock #nf
For t = 0 To lista_acceso.ListCount - 1
   Print #nf, lista_acceso.List(t)
   Print #nf, lista_acceso.Selected(t)
Next t
Unlock #nf
Close #nf
   
valido1 = 0
End Sub


