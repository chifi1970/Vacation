VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEditEvent_holidays.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   7440
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   1920
   End
   Begin Project1.lvButtons_H CmdDelete 
      Height          =   1095
      Left            =   8400
      TabIndex        =   46
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "FrmEditEvent_holidays.frx":E3CE
      ImgSize         =   40
      cBack           =   0
   End
   Begin VB.ComboBox cbo_horas 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton op_dias 
      BackColor       =   &H8000000C&
      Caption         =   "Hours"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   38
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton op_dias 
      BackColor       =   &H8000000C&
      Caption         =   "Half Day"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton op_dias 
      BackColor       =   &H8000000C&
      Caption         =   "All day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin Project1.lvButtons_H CmdAccept 
      Height          =   1095
      Left            =   8400
      TabIndex        =   32
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   1
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "FrmEditEvent_holidays.frx":F4F9
      ImgSize         =   40
      cBack           =   16777215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   8760
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   4875
      Begin VB.Frame marco_dias 
         Height          =   1215
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPEndDate 
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   213778433
            CurrentDate     =   44676
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "End date:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame marco_dias 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPStartTime 
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   213778435
         UpDown          =   -1  'True
         CurrentDate     =   44676
      End
      Begin MSComCtl2.DTPicker DTPEndTime 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   213778435
         UpDown          =   -1  'True
         CurrentDate     =   44676
      End
   End
   Begin VB.PictureBox PicPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   5880
      Picture         =   "FrmEditEvent_holidays.frx":104C9
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   11
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkIsSerie 
      Caption         =   "Is serie (Only Icon)"
      Height          =   195
      Left            =   5400
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox ChkNotify 
      Caption         =   "Notify (Only Icon)"
      Height          =   195
      Left            =   5280
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox ChkPrivate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bloqueado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditEvent_holidays.frx":4E93B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditEvent_holidays.frx":4EC8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditEvent_holidays.frx":4EFDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditEvent_holidays.frx":4F331
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEditEvent_holidays.frx":4F683
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicColor 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   220
      Left            =   4320
      ScaleHeight     =   225
      ScaleWidth      =   1530
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtBody 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3680
      Width           =   2415
   End
   Begin VB.TextBox TxtLocation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CheckBox ChkAllDay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All day"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtSubject 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   560
      Width           =   3375
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   390
      Left            =   6240
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1455
      Left            =   9120
      TabIndex        =   26
      Top             =   7440
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
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
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12632319
      CalendarForeColor=   0
      CalendarTitleBackColor=   12632319
      Format          =   213778433
      CurrentDate     =   44676
   End
   Begin Project1.lvButtons_H op_aprobado 
      Height          =   345
      Index           =   1
      Left            =   1080
      TabIndex        =   43
      Top             =   8720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "Sales Director"
      CapAlign        =   2
      BackStyle       =   1
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
      Enabled         =   0   'False
      cBack           =   12632319
   End
   Begin Project1.lvButtons_H btnanular_permiso 
      Height          =   315
      Left            =   480
      TabIndex        =   44
      Top             =   7320
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "Cancel permission"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   255
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H op_aprobado 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   45
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "HR"
      CapAlign        =   2
      BackStyle       =   1
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
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H CmdCancel 
      Height          =   855
      Left            =   8400
      TabIndex        =   47
      Top             =   8280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      CapAlign        =   2
      BackStyle       =   1
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
      Image           =   "FrmEditEvent_holidays.frx":4F9D5
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btncash_out 
      Height          =   1095
      Left            =   8400
      TabIndex        =   50
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   1
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "FrmEditEvent_holidays.frx":50BB8
      ImgSize         =   40
      cBack           =   0
   End
   Begin Project1.lvButtons_H btnsick 
      Height          =   1095
      Left            =   8400
      TabIndex        =   58
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "FrmEditEvent_holidays.frx":58527
      ImgSize         =   40
      cBack           =   0
   End
   Begin VB.Label lblaprobada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date conditional on approval"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      Left            =   4500
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   1965
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   20
      Left            =   2280
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblfecha_horas_ganadas 
      BackStyle       =   0  'Transparent
      Caption         =   "dd/mm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3180
      TabIndex        =   57
      Top             =   3300
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "from"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2760
      TabIndex        =   56
      Top             =   3300
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "New hours earned"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2760
      TabIndex        =   55
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblhoras_nuevas 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   2280
      TabIndex        =   54
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label lblano 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2023"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1320
      TabIndex        =   52
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblday 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   1260
      TabIndex        =   51
      Top             =   1020
      Width           =   570
   End
   Begin VB.Image img_cash_out 
      Height          =   2295
      Left            =   360
      Picture         =   "FrmEditEvent_holidays.frx":59DA4
      Stretch         =   -1  'True
      Top             =   3800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblhired_date 
      BackStyle       =   0  'Transparent
      Caption         =   "-----"
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
      Height          =   255
      Left            =   4920
      TabIndex        =   49
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hired date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3960
      TabIndex        =   48
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   2040
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image sello_aprobado 
      Height          =   1335
      Left            =   360
      Picture         =   "FrmEditEvent_holidays.frx":5E559
      Stretch         =   -1  'True
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape barra 
      FillColor       =   &H00808080&
      Height          =   135
      Index           =   2
      Left            =   2880
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape barra 
      FillColor       =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   2640
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Half day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape barra 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   2400
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Full day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblhrs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   35
      Top             =   4200
      Width           =   195
   End
   Begin VB.Label lblhours 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblhoras_totales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2280
      TabIndex        =   30
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Used hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblhoras_usadas 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   2280
      TabIndex        =   28
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label lblid_vac 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7580
      TabIndex        =   25
      Top             =   220
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   24
      Top             =   195
      Width           =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Hours "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Top             =   3280
      Width           =   1740
   End
   Begin VB.Label lblhoras_disponibles 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   855
      TabIndex        =   15
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show time as:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   1000
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public EventKey As Long
Dim aprobadoby As Integer, aprobado_status As Boolean, horas_seleccionadas As Integer, total_horas_en_tabla As Integer, total_horas_usadas As Integer
Dim datos(12), horas_disponibles As Integer, horas_tomadas As Integer, horas_nuevas As Integer, ID_vacation_profile As Integer, id_vacationhours As Integer, manager As Integer
Dim seg As Integer, cashout As Boolean, horas_reales_disponibles As Integer, horas_usadas As Integer
Private Sub btnanular_permiso_Click()
On Error Resume Next

   ' verifica si la fecha ya paso
    dia_solicitado$ = Mid$(Format(DTPStartDate, "mm/dd/yyyy"), 4, 2)
    mes_solicitado$ = Left(Format(DTPStartDate, "mm/dd/yyyy"), 2)
    ano_solicitado$ = Right(Format(DTPStartDate, "mm/dd/yyyy"), 4)
    
    dia_actual$ = Mid$(Format(Now, "mm/dd/yyyy"), 4, 2)
    mes_actual$ = Left(Format(Now, "mm/dd/yyyy"), 2)
    ano_actual$ = Right(Format(Now, "mm/dd/yyyy"), 4)
    
    
    If Val(ano_solicitado$) < Val(ano_actual$) Then
       ' no se puede
       MsgBox "You can no longer cancel this permission", 16, "Attention"
       Exit Sub
    ElseIf Val(ano_solicitado$) = Val(ano_actual$) Then
         If Val(mes_solicitado$) < Val(mes_actual$) Then
            ' no se puede
            MsgBox "You can no longer cancel this permission", 16, "Attention"
            Exit Sub
         ElseIf Val(mes_solicitado$) = Val(mes_actual$) Then
               If Val(dia_solicitado$) < Val(dia_actual$) Then
                    ' no se puede
                    MsgBox "You can no longer cancel this permission", 16, "Attention"
                    Exit Sub
               Else
                    ' si se puede
                    
               End If
         End If
    End If
    
    
    
    


op_aprobado(0).Value = False
op_aprobado(1).Value = False

aprobado_status = False
aprobadoby = ""

ChkPrivate.Value = False
datos(5) = False

 sello_aprobado.Visible = False

End Sub

Private Sub btncash_out_Click()
On Error Resume Next




 If Form1.lblemp_id.Caption = "76" Or user$ = "GJIMENEZ" Then
 
 Else
  ' verifica si la fecha ya paso
    dia_solicitado$ = Mid$(Format(DTPStartDate, "mm/dd/yyyy"), 4, 2)
    mes_solicitado$ = Left(Format(DTPStartDate, "mm/dd/yyyy"), 2)
    ano_solicitado$ = Right(Format(DTPStartDate, "mm/dd/yyyy"), 4)
    
    dia_actual$ = Mid$(Format(Now, "mm/dd/yyyy"), 4, 2)
    mes_actual$ = Left(Format(Now, "mm/dd/yyyy"), 2)
    ano_actual$ = Right(Format(Now, "mm/dd/yyyy"), 4)
    
    
    If Val(ano_solicitado$) < Val(ano_actual$) Then
       ' no se puede
       MsgBox "You can no longer modify this date", 16, "Attention"
       Exit Sub
    ElseIf Val(ano_solicitado$) = Val(ano_actual$) Then
         If Val(mes_solicitado$) < Val(mes_actual$) Then
            ' no se puede
            MsgBox "You can no longer modify this date", 16, "Attention"
            Exit Sub
         ElseIf Val(mes_solicitado$) = Val(mes_actual$) Then
               If Val(dia_solicitado$) < Val(dia_actual$) Then
                    ' no se puede
                    MsgBox "You can no longer modify this date", 16, "Attention"
                    Exit Sub
               Else
                    ' si se puede
                    
               End If
         End If
    End If
    
End If





If img_cash_out.Visible = True Then
  img_cash_out.Visible = False
  cashout = "0"
Else
  img_cash_out.Visible = True
  cashout = "1"
End If


End Sub

Private Sub cbo_horas_Click()
On Error Resume Next
  datos(4) = cbo_horas.List(cbo_horas.ListIndex)
    
  lblhoras_disponibles.Caption = Format(total_horas_en_tabla - datos(4), "##0")
  lblhoras_usadas.Caption = Format(total_horas_usadas + datos(4), "##0")
  
End Sub


Private Sub CmdAccept_Click()
On Error Resume Next

   Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   
   ubicacion_de_trabajo$ = TxtLocation.Text
   
   If datos(4) = "" And horas_disponibles > 0 Then
       datos(4) = 8
   End If
   
   
  If user$ <> "GJIMENEZ" Then
   If (datos(4) > horas_disponibles And manager < 1) Or Val(lblhoras_disponibles.Caption) <= 0 Then
      MsgBox "Sorry, but you don't have enough hours available to use as vacation", 16, "Upss!"
      Exit Sub
   
   End If
 End If
    
   ' verifica que no este esa fecha ya grabada
   
   
   existe = 0
   ID_vacation$ = ""
   sSelect = "select idvacation from vacationsprogram where daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "' and idemployee='" + datos(0) + "'  and active='1'"
   Rs.Open sSelect, base, adOpenUnspecified
   If Err Then
        Conecta_SQL
   End If
   ID_vacation$ = Rs(0)
   Rs.Close
      
      
   If Val(ID_vacation$) <> ID_vacaciones$ And ID_vacaciones$ <> "" Then
       ID_vacaciones$ = ID_vacation$
   End If
   
   
    
   
   
   If aprobadoby = 2 Or UCase(user$) = "GJIMENEZ" Then
      ' obtiene el ID de GABY
      sSelect = "select idemployee from employeeinfo where username='gjimenez'"
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
      ID_aprobado_por$ = Rs(0)
      Rs.Close
      'datos(5) = True
      
   ElseIf aprobadoby = 1 Or UCase(user$) = "KDELGADILLO" Then
      ' obtiene el ID de KARINA
      sSelect = "select idemployee from employeeinfo where username='kdelgadillo'"
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
      ID_aprobado_por$ = Rs(0)
      Rs.Close
      
   ElseIf aprobadoby = 1 Then
   
      ' obtiene el ID de HR
      sSelect = "select idemployee from employeeinfo where username='kdelgadillo'"
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
      ID_aprobado_por$ = Rs(0)
      Rs.Close
      'datos(5) = True
      
   Else
         
      ID_aprobado_por$ = "NULL"
      aprobado_status = False
      'datos(5) = False
      
   End If
   
   
   
   
   
   NUEVO = 0
   
   
   
   If ID_vacaciones$ = "" Then
      ' graba registro
      
       
       If datos(4) = "" Then
         MsgBox "Sorry, but the agent doesn't have enough hours available to use as vacation", 16, "Upss!"
         Unload Me
         Exit Sub
       End If
      
      
       ' Cash Out=0  si es vacacion normal ,  =1 si requiere el dinero y trabajará ese dia
      
      
       sSelect = "insert into vacationsprogram (idemployee, idmanager, approvedby, daterequested, hours, approved, notes, datecreated, lastupdated, active, cashout) " & _
       "VALUES ('" + datos(0) + "', '" + datos(1) + "', null, '" + Format(DTPStartDate, "mm/dd/yyyy") + "', '" + Format(datos(4), "0") + "', '0', '" & _
       TxtBody.Text + "', '" + Format(Now, "mm/dd/yyyy hh:mm") + "', '" + Format(Now, "mm/dd/yyyy hh:mm") + "', '1', '" + Format(cashout, "0") + "')"
       
       PicColor.BackColor = vborange
       colorx = vborange
      
       transfiere$ = "NONE"
       
       NUEVO = 1
           
  Else
  
  
        If datos(4) = "" Then
         Unload Me
         Exit Sub
        End If
      
  
       
        sSelect = "update vacationsprogram set idemployee='" + datos(0) + "', idmanager='" + datos(1) + "', approvedby=" + ID_aprobado_por$ + ", " & _
        "daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "', hours='" + Format(datos(4), "0") + "', approved='" & _
        Format(aprobado_status, "0") + "', notes='" + TxtBody.Text + "', lastupdated='" + Format(Now, "mm/dd/yyyy hh:mm") + "', active='1', cashout='" + Format(cashout, "0") + "' " & _
        "where idvacation='" + ID_vacaciones$ + "'"
       
        If datos(5) = True Or datos(5) = 1 Then
        
           If datos(4) = "8" Then
           
             If cashout = "False" Then
                PicColor.BackColor = vbverde
                colorx = vbverde
             Else
                PicColor.BackColor = rosa
                colorx = rosa
             End If
             
           Else
             
             If cashout = "False" Then
                PicColor.BackColor = vbverde_claro
                colorx = vbverde_claro
             Else
                   PicColor.BackColor = rosa_claro
                   colorx = rosa_claro
             End If
             
           End If
            
        Else
         
         If fecha_accesible <= 1 Then
           ' no esta autorizado aun
            PicColor.BackColor = vborange
            colorx = vborange
         Else
            ' fecha condicionada
            PicColor.BackColor = vbamarillo
            colorx = vbamarillo
         
         End If
            
        End If
        
        
        If ID_aprobado_por$ = "76" Then
          transfiere$ = "ENVIA CORREO2"
        Else
          transfiere$ = "NONE"
        End If
       
       
  End If
  
  
  Rs.Open sSelect, base, adOpenUnspecified
   
    If Err Then
      Conecta_SQL
    End If
    
   
   Rs.Close
   
  
  
  
  ' actualiza el ID de vacaciones
   
   If ID_vacaciones$ = "" Then
    
     sSelect = "select idvacation from vacationsprogram where daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "' and idemployee='" + datos(0) + "' and idmanager='" + datos(1) + "'  and active='1'"
   
     Rs.Open sSelect, base, adOpenUnspecified
   
     If Err Then
       Conecta_SQL
     End If
    
    ID_vacaciones$ = RTrim(LTrim(Rs(0)))
    Rs.Close
   
   End If
  
   
   
   If NUEVO = 1 Then
    
       ' actualiza las horas que le quedan
        If horas_reales_disponibles > 0 Then
           hrs_disp = horas_disponibles - datos(4)
           Hrs_usadas = horas_tomadas + datos(4)
           hold_hrs_disp = horas_nuevas
           If horas_usadas < 0 Then horas_usadas = 0
           hold_hrs_usadas = horas_usadas
           
        Else
           hrs_disp = 0
           Hrs_usadas = horas_tomadas
           hold_hrs_disp = horas_nuevas - datos(4)
           hold_hrs_usadas = horas_usadas + datos(4)
        End If
    
        sSelect = "update vacationshoursemprel set idemployee='" + datos(0) + "', idvacationprofile='" + Format(ID_vacation_profile, "#0") + "', " & _
        "earnedhours='" + Format(hrs_disp, "##0") + "', takenhours='" + Format(Hrs_usadas, "##0") + "', Holdearnhours='" + Format(hold_hrs_disp, "##0") + "', " & _
        "Holdtakenhours='" + Format(hold_hrs_usadas, "##0") + "', active=1  where idemployee='" + Format(datos(0), "##0") + "'"
        
        Rs.Open sSelect, base, adOpenUnspecified
      
        Rs.Close
        
        
   End If


    Dim StartDate As Date
    Dim EndDate As Date
    
    StartDate = Format(DateValue(DTPStartDate.Value), "mm/dd/yy") & " 00:00"
    'StartDate = DateAdd("h", DTPStartTime.Hour, StartDate)
    'StartDate = DateAdd("n", DTPStartTime.Minute, StartDate)
    EndDate = Format(DateValue(DTPEndDate.Value), "mm/dd/yy") & " 23:59"
    'EndDate = DateAdd("H", DTPEndTime.Hour, EndDate)
    'EndDate = DateAdd("n", DTPEndTime.Minute, EndDate)
    
    f1$ = StartDate & " 00:00:00 AM"
    f2$ = EndDate
    
         
    ' PicColor.BackColor
    If EventKey > 0 Then
    
                
        Form1.ucCalendar1.UpdateEventData EventKey, TxtSubject.Text, TxtLocation.Text, ID_vacaciones$, f1$, f2$, colorx, ChkAllDay.Value, TxtBody.Text, , ChkIsSerie.Value, ChkNotify.Value, ChkPrivate.Value, ImageCombo1.SelectedItem.Index - 1, correo_agente$, correo_manager$
    Else
        Form1.ucCalendar1.AddEvents TxtSubject.Text, f1$, f2$, colorx, ChkAllDay.Value, TxtBody.Text, , ChkIsSerie.Value, ChkNotify.Value, ChkPrivate.Value, ImageCombo1.SelectedItem.Index - 1, , , correo_agente$, correo_manager$
        
      
    End If
    
   ' .AddEvents lblemployee.Caption, Date & " 00:00", Date & " 23:59", vbBlack
   
   
   
   ' -----------------------------------------------------------------------------
   
   
   ' --------------------------------------------------------------------------------
   
   
    fecha_accesible = 0
   ' transfiere$ = "ENVIA CORREO"
    
    Unload Me
End Sub

Private Sub CmdCancel_Click()
On Error Resume Next
    
    fecha_accesible = 0
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    On Error Resume Next
    
   'GoTo graba
 
    
    ' verifica si la fecha ya paso
    dia_solicitado$ = Mid$(Format(DTPStartDate, "mm/dd/yyyy"), 4, 2)
    mes_solicitado$ = Left(Format(DTPStartDate, "mm/dd/yyyy"), 2)
    ano_solicitado$ = Right(Format(DTPStartDate, "mm/dd/yyyy"), 4)
    
    dia_actual$ = Mid$(Format(Now, "mm/dd/yyyy"), 4, 2)
    mes_actual$ = Left(Format(Now, "mm/dd/yyyy"), 2)
    ano_actual$ = Right(Format(Now, "mm/dd/yyyy"), 4)
    
    
   If UCase(Form1.lblemployee.Caption) <> "GABRIELA JIMENEZ" Then
    
     If Val(ano_solicitado$) < Val(ano_actual$) Then
       ' no se puede
       MsgBox "You can no longer cancel this date", 16, "Attention"
       Exit Sub
     ElseIf Val(ano_solicitado$) = Val(ano_actual$) Then
         If Val(mes_solicitado$) < Val(mes_actual$) Then
            ' no se puede
            MsgBox "You can no longer cancel this date", 16, "Attention"
            Exit Sub
         ElseIf Val(mes_solicitado$) = Val(mes_actual$) Then
               If Val(dia_solicitado$) < Val(dia_actual$) Then
                    ' no se puede
                    MsgBox "You can no longer cancel this date", 16, "Attention"
                    Exit Sub
               Else
                    ' si se puede
                    
               End If
         End If
     End If
    
   End If
    
graba:
    
    
    
    Dim sSelect As String
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
    
    
    If lblid_vac.Caption = "" Then
       Exit Sub
    End If
    
    
    
    R$ = MsgBox("Are you sure you want to delete this event?", 4, "Attention")
    If R$ = "7" Then Exit Sub
    
    
    
    Form1.ucCalendar1.RemoveEvent EventKey
    fecha_accesible = 0
    
    
    
    
   
   
    ' borra la reservacion del dia de vacaciones
     ' sSelect = "delete from VacationsProgram where idvacation='" + lblid_vac.Caption + "'"
     
    sSelect = "update VacationsProgram set active='0' where idvacation='" + lblid_vac.Caption + "'"
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
    
     Rs.Close
     
    
    
    ' regresa las horas borradas
    sSelect = "select earnedhours, takenhours, idvacationhours, holdearnhours, holdtakenhours from vacationshoursemprel where idemployee='" + Format(datos(0), "##0") + "'"
    Rs.Open sSelect, base, adOpenUnspecified
     horas_ganadas = Rs(0)
     horas_tomadas = Rs(1)
     id_vacaciones_horas$ = Rs(2)
     hold_horas_ganadas = Rs(3)
     hold_horas_tomadas = Rs(4)
    Rs.Close



    If lblhired_date.Caption <> "" Then
                 
          
          If hold_horas_ganadas > 0 And hold_horas_tomadas > 0 And horas_ganadas = 0 Then
              hold_horas_ganadas = hold_horas_ganadas + Val(datos(4))
              hold_horas_tomadas = hold_horas_tomadas - Val(datos(4))
          ElseIf hold_horas_ganadas > 0 And hold_horas_tomadas <= 0 And horas_ganadas = 0 Then
              hold_horas_ganadas = hold_horas_ganadas + Val(datos(4))
              hold_horas_tomadas = 0
          
          ElseIf horas_ganadas >= 0 And horas_tomadas > 0 Then
              horas_ganadas = horas_ganadas + Val(datos(4))
              horas_tomadas = horas_tomadas - Val(datos(4))
              
          ElseIf horas_ganadas >= 0 And horas_tomadas = 0 Then
              horas_ganadas = horas_ganadas + Val(datos(4))
              horas_tomadas = 0
                    
          End If
          
          
        
    End If
        
       
    
    sSelect = "update vacationshoursemprel set earnedhours='" + Format(horas_ganadas, "##0") + "' , takenhours='" + Format(horas_tomadas, "##0") + "', " & _
    "Holdearnhours='" + Format(hold_horas_ganadas, "##0") + "', Holdtakenhours='" + Format(hold_horas_tomadas, "##0") + "'  where idvacationhours='" + id_vacaciones_horas$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
    
    
    Unload Me
End Sub

Private Sub Command1_Click()
    PicPalette.Move Command1.Left, Command1.top + Command1.Height
    PicPalette.Visible = Not PicPalette.Visible
End Sub

Private Sub DTPStartDate_Change()
On Error Resume Next
DTPEndDate.Value = DTPStartDate.Value

End Sub

Private Sub Form_Activate()
On Error Resume Next

    'If EventKey = 0 Then
    '    CmdAccept.Caption = "Add New"
    'Else
    '    CmdAccept.Caption = "Update"
    'End If
    'CmdDelete.Visible = EventKey <> 0
End Sub




Private Sub Form_Load()
    On Error Resume Next
    
      
       
    'If valido1 = 777 Then
        
    '   Unload Me
       cashout = "0"
       Exit Sub
       
       
       
    'End If
    
    
   
 
    
    
End Sub

Private Sub op_aprobado_Click(Index As Integer)
On Error Resume Next
aprobadoby = Index + 1
aprobado_status = True
ChkPrivate.Value = 1
datos(5) = True
sello_aprobado.Visible = True
transfiere$ = "ENVIA CORREO"

End Sub




Private Sub op_dias_Click(Index As Integer)
 On Error Resume Next
  If valido1 = 777 Then
    valido1 = 0
    op_dias(Index).Value = True
    Exit Sub
  End If
  
  marco_dias(1).Visible = False
  DTPEndDate.Value = DTPStartDate.Value
  
  If Index = 0 Then
    cbo_horas.Enabled = False
    barra(0).FillStyle = 0
    barra(1).FillStyle = 1
    barra(2).FillStyle = 1
    cbo_horas.ListIndex = -1
    PicColor.BackColor = vbverde
    colorx = vbverde
    datos(4) = 8
    
  ElseIf Index = 1 Then
    cbo_horas.Enabled = False
    barra(0).FillStyle = 1
    barra(1).FillStyle = 0
    barra(2).FillStyle = 1
    cbo_horas.ListIndex = -1
    PicColor.BackColor = vbverde_claro
    colorx = vbverde_claro
    datos(4) = 4
    
  Else
    cbo_horas.Enabled = True
    barra(0).FillStyle = 1
    barra(1).FillStyle = 1
    barra(2).FillStyle = 0
    PicColor.BackColor = vbverde_claro
    colorx = vbverde_claro
    If cbo_horas.ListIndex = -1 Then
       cbo_horas.ListIndex = Val(datos(4)) - 1
    End If
    datos(4) = cbo_horas.List(cbo_horas.ListIndex)
  
    
  End If
  
   
   
  
  lblhoras_disponibles.Caption = Format(horas_disponibles, "##0")
  lblhoras_usadas.Caption = Format(horas_tomadas, "##0")
  
  
  op_dias(Index).Value = True
  valido1 = 777
  op_dias_Click (Index)
  valido1 = 0
End Sub

Private Sub PicPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicColor.BackColor = PicPalette.Point(X, Y)
    PicPalette.Visible = False
End Sub

Public Sub carga_datos()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
    
    
   '  Erase datos
    
    sSelect = "select idvacation, vac.idemployee, idmanager, approvedby, daterequested, hours, approved, notes, emp.payrolllinkid, emp.firstname, " & _
    "emp.lastname1, emp.username from VacationsProgram vac " & _
    "join employeeinfo emp on emp.idemployee=vac.idemployee " & _
    "where idvacation='" + ID_vacaciones$ + "'  and vac.active='1'"

   
   ' sSelect = "select * from vacationsprogram where idvacation='" + ID_vacaciones$ + "'"
       
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
   
    Erase datos
   
    If Grid2.Rows > 1 Then
      Grid2.Row = 1
      Grid2.Col = 2
      datos(0) = Grid2.Text  ' id_employee
      
      Grid2.Col = 3
      datos(1) = Grid2.Text  ' id_manager
      
      Grid2.Col = 4
      datos(2) = Grid2.Text  ' aprobado_por
      
      Grid2.Col = 5
      datos(3) = Format(Grid2.Text, "mm/dd/yyyy") ' Fecha
      
      DTPStartDate.Value = datos(3)
      
      
      Grid2.Col = 6
      datos(4) = Grid2.Text  ' horas
      
      horas_seleccionadas = Val(Grid2.Text)
      lblhrs.Caption = Format(horas_seleccionadas, "#0")
      
            
      If datos(4) = "8" Then
        op_dias_Click (0)
        op_dias(0).Value = True
       
      ElseIf datos(4) = "4" Then
        op_dias_Click (1)
        op_dias(1).Value = True
      Else
        op_dias_Click (2)
        op_dias(2).Value = True
        
      End If
      
      
      
      Grid2.Col = 7
      datos(5) = Grid2.Text  ' aprobado o no?
      
      If datos(5) = True Or datos(5) = 1 Then
          If datos(2) = 76 Then
             op_aprobado(1).Value = True

             
             sello_aprobado.Visible = True

             ChkPrivate.Value = 1

          ElseIf datos(2) = 155 Then
             op_aprobado(0).Value = True
             
              sello_aprobado.Visible = True

             ChkPrivate.Value = 1

          Else
             op_aprobado(0).Value = False
             op_aprobado(1).Value = False
             
              sello_aprobado.Visible = False

             ChkPrivate.Value = False

          End If
      Else
         op_aprobado(0).Value = False
         op_aprobado(1).Value = False
         
          sello_aprobado.Visible = False

         ChkPrivate.Value = False

      End If
      
      
      
      Grid2.Col = 8
      datos(6) = Grid2.Text  ' notas
      TxtBody.Text = datos(6)
      nota$ = datos(6)
      
      
      Grid2.Col = 9
      datos(7) = Grid2.Text  ' payroll id
      
      Grid2.Col = 10
      datos(8) = Grid2.Text  ' nombre
      
      Grid2.Col = 11
      datos(9) = Grid2.Text  ' apellido
      
      Grid2.Col = 12
      datos(10) = Grid2.Text  ' username
      
      TxtSubject.Text = datos(8) + " " + datos(9)
      
      
    Else
    
      
       datos(0) = Form1.lblemp_id.Caption
       datos(1) = Form1.lblmanager_ID.Caption
       
       datos(3) = Format(DTPStartDate.Value, "mm/dd/yyyy")
       ubicacion_de_trabajo$ = TxtLocation.Text
       
    
    End If
   
   
   
   ' carga la oficina donde esta ubicado
    
    
     sSelect = "select Office from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and Username='" + datos(10) + "'"


    
     
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
       TxtLocation.Text = Grid2.Text
    End If
    
   

End Sub


Public Sub Verifica_horas()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

  
   
      
    sSelect = "select * from vacationshoursemprel where idemployee='" + datos(0) + "' and active=1"
       
   
    Rs.Open sSelect, base, adOpenUnspecified
            
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
    If Grid2.Rows > 1 Then
       Grid2.Row = 2
       
       Grid2.Col = 1
       id_vacationhours = Val(Grid2.Text)
       
       Grid2.Col = 3
       ID_vacation_profile = Val(Grid2.Text)
       
       
       Grid2.Col = 4
       horas_disponibles = Val(Grid2.Text)
       horas_reales_disponibles = horas_disponibles
       
       lblhoras_disponibles.Caption = Format(horas_disponibles, "##0")
       
       
       Grid2.Col = 5
       horas_tomadas = Val(Grid2.Text)
       lblhoras_usadas.Caption = Format(horas_tomadas, "##0")
       
       
       Grid2.Col = 6
       horas_nuevas = Val(Grid2.Text)
       lblhoras_nuevas.Caption = Format(horas_nuevas, "##0")
       lblfecha_horas_ganadas.Caption = Format(Left(Form1.lblhired_date.Caption, 5), "mmm,dd")
       
       
       Grid2.Col = 7
       horas_usadas = Val(Grid2.Text)
       
       
        ' carga fecha de contratacion
       sSelect = "select CONVERT(varchar,HireDate,32) as Date from EmployeeInfo where active=1 and HireDate is not null and idemployee='" + datos(0) + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       lblhired_date.Caption = Rs(0)
       Rs.Close
       
       
       
       If lblhired_date.Caption <> "" Then
       
          mes_aniversario = Val(Left(lblhired_date.Caption, 2))
          dia_aniversario = Val(Mid$(lblhired_date.Caption, 4, 2))
          
       
          dia_selecto = Val(Format(DTPStartDate, "dd"))
          mes_selecto = Val(Format(DTPStartDate, "mm"))
          ano_selecto = Format(DTPStartDate, "yyyy")
          
          dia_actual = Val(Format(Now, "dd"))
          mes_actual = Val(Format(Now, "mm"))
          ano_actual = Val(Format(Now, "yyyy"))
          
          
          If mes_selecto = mes_aniversario And Val(ano_selecto) >= ano_actual Then
              If dia_selecto >= dia_aniversario Then
                    horas_disponibles = horas_disponibles + horas_nuevas
                    lblhoras_disponibles.Caption = Format(horas_disponibles, "##0")
                    ' horas_nuevas = 0
                    lblhoras_nuevas.Caption = Format(0, "##0")
                    lblhoras_usadas.Caption = Format(horas_tomadas + horas_usadas, "##0")
              
              End If
              
          ElseIf mes_selecto > mes_aniversario And Val(ano_selecto) >= ano_actual Then
                    
                    horas_disponibles = horas_disponibles + horas_nuevas
                    lblhoras_disponibles.Caption = Format(horas_disponibles, "##0")
                    ' horas_nuevas = 0
                    lblhoras_nuevas.Caption = Format(0, "##0")
                    lblhoras_usadas.Caption = Format(horas_tomadas + horas_usadas, "##0")
                    
          ElseIf mes_selecto < mes_aniversario And Val(ano_selecto) > ano_actual Then
              
                    horas_disponibles = horas_disponibles + horas_nuevas
                    lblhoras_disponibles.Caption = Format(horas_disponibles, "##0")
                    ' horas_nuevas = 0
                    lblhoras_nuevas.Caption = Format(0, "##0")
                    lblhoras_usadas.Caption = Format(horas_tomadas + horas_usadas, "##0")
                    
              
          End If
          
          
          
        
       End If
        
       
       
       
       
       
       total_horas = horas_disponibles + horas_tomadas + horas_usadas
       
      lblhoras_totales.Caption = Format(total_horas, "##0")
      
      
      
      
       
    End If
    
    
    
    
    
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1

If seg >= 1 Then
  carga_inicial
  Show
  seg = 0
  Timer1.Enabled = False
End If


End Sub



Public Sub carga_inicial()
On Error Resume Next

 If valido1 = 5 Then
   Exit Sub
 End If


Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
   
 
 
 

 manager = 0
 
 
 
 
    contador = 0
    
    
    If reservado = True Then
      If Form1.lblemp_id.Caption = "76" Or Form1.lblemp_id.Caption = "155" Or Form1.lblemp_id.Caption = "136" Or user$ = "GJIMENEZ" Then
         manager = 1
      Else
        'usuario$ = ""
       
        'Unload Me
        'Exit Sub
      End If
    Else
    
    
    End If
    
    
    
    
    
    
    
    
    If UCase$(usuario$) <> UCase(Form1.lblemployee.Caption) And usuario$ <> "" Then
       If Form1.lblemp_id.Caption = "76" Or Form1.lblemp_id.Caption = "155" Or Form1.lblemp_id.Caption = "136" Then
         manager = 1
       End If
       
       If manager = 0 Then
          usuario$ = ""
          Unload Me
          Exit Sub
       End If
    Else
         
    End If
    
    
    
    
    'Erase datos
    
    
    If Form1.lblemp_id.Caption = "76" Or Form1.lblemp_id.Caption = "155" Or Form1.lblemp_id.Caption = "136" Then
      btnanular_permiso.Visible = True
      op_aprobado(1).Enabled = True
      CmdDelete.Visible = True
      Frame1.Enabled = True
     ' CmdAccept.Visible = True
      
    End If
    
    
   
    
    
    Form2.top = 0
    Form2.Left = (Screen - Form2.Width) / 2
    
    PicPalette.PaintPicture PicPalette.Picture, 0, 0, PicPalette.ScaleWidth, PicPalette.ScaleHeight
    
    ImageCombo1.ComboItems.Add , , "Busy", 1, 1
    ImageCombo1.ComboItems.Add , , "Free", 2, 2
    ImageCombo1.ComboItems.Add , , "Out of office", 3, 3
    ImageCombo1.ComboItems.Add , , "Tentative", 4.4
    ImageCombo1.ComboItems.Add , , "Working elsewhere", 5, 5
    ImageCombo1.ComboItems.Item(1).Selected = True
    
    
    'DTPStartDate.Value = Format(Now, "mm/dd/yyyy")
    'DTPEndDate.Value = Format(Now, "mm/dd/yyyy")
    
   
    
       
       
    lblid_vac.Caption = ID_vacaciones$
     
    
    cbo_horas.Clear
    For t = 1 To 8
      cbo_horas.AddItem t
    Next t
    
    carga_datos
    
     If fecha_accesible = 1 Then
        lblaprobada.Visible = True
    
    End If
    
    'TxtLocation.Text = Form1.lbllocation.Caption
    'TxtSubject.Text =  Form1.lblemployee.Caption
    
    
    Verifica_horas
    
    'total_horas_en_tabla = horas_disponibles + horas_seleccionadas
    'total_horas_usadas = horas_tomadas - horas_seleccionadas
    
   If Form1.lblemp_id.Caption = "76" Or Form1.lblemp_id.Caption = "155" Then
   
   Else
    If Val(lblid_vac.Caption) > 0 Then
       Frame1.Enabled = False
       ' CmdAccept.Visible = False
    End If
   End If
   
   lblhired_date.Caption = Form1.lblhired_date.Caption
   
   
   a$ = TxtSubject.Text
   
   For t = 0 To Form1.cbo_users.ListCount - 1
      B$ = LTrim(RTrim(Left(Form1.cbo_users.List(t), 20)))
      If UCase(a$) = UCase(B$) Then
        id_empl$ = RTrim(LTrim(Right(Form1.cbo_users.List(t), 6)))
        Exit For
      End If
   Next t
   
   ' carga ID del empleado
   
   
   a$ = TxtSubject.Text
   
   existe = 0
    id_emp$ = ""
   For t = 0 To Form1.cbo_users.ListCount - 1
       nombre_empleado$ = UCase(RTrim(LTrim(Left(Form1.cbo_users.List(t), 30))))
       
   
       If nombre_empleado$ = UCase(a$) Then
          id_emp$ = RTrim(LTrim(Right(Form1.cbo_users.List(t), 6)))
          existe = 1
          Exit For
       End If
       
   Next t
   
   
   
   
   
   
   
   
   ' carga fecha de contratacion
   sSelect = "select CONVERT(varchar,HireDate,32) as Date from EmployeeInfo where active=1 and HireDate is not null and idemployee='" + id_emp$ + "'"
   
   Rs.Open sSelect, base, adOpenUnspecified
      
   lblhired_date.Caption = Rs(0)
     
   Rs.Close
   
 
 
 
   If lblid_vac.Caption <> "" Then
 
   ' carga imagen de dinero si es CASH OUT
     sSelect = "SELECT cashout FROM VacationsProgram where active=1 and idemployee='" + id_emp$ + "' and idvacation='" + lblid_vac.Caption + "'"
   
     Rs.Open sSelect, base, adOpenUnspecified
      
     cashout = Rs(0)
     
     Rs.Close
     
     
     If cashout = "1" Then
       img_cash_out.Visible = True
     Else
       img_cash_out.Visible = False
     End If
     
   
   
   End If
   
 
 
    ' carga la fecha
    lblday.Caption = Format(DTPStartDate, "dd")
    lblmonth.Caption = Format(DTPStartDate, "mmmm")
    lblano.Caption = Format(DTPStartDate, "yyyy")
 
   
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

If DTPStartDate = "4/25/2022" Then
    Unload Me
    Exit Sub
End If


End Sub


