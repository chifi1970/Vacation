VERSION 5.00
Begin VB.Form Forma_seguridad 
   BackColor       =   &H80000010&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btn_ok 
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
   Begin Project1.lvButtons_H btnborra 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   873
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
      Image           =   "Forma_seguridad_holidays.frx":0000
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the password:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Forma_seguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_ok_Click()
On Error Resume Next
transfiere$ = txtPassword.Text
Unload Me

End Sub

Private Sub btnborra_Click()
On Error Resume Next
txtPassword.Text = ""
txtPassword.SetFocus
End Sub



Private Sub Form_Load()
On Error Resume Next
top = 4700  '(Screen.Height - Height) / 2
Left = ((Screen.Width - Width) / 2) + 800


End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
  btn_ok_Click
  Exit Sub
End If
End Sub


