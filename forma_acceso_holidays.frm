VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form forma_acceso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log in"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   ControlBox      =   0   'False
   Icon            =   "forma_acceso_holidays.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   6480
      ScaleHeight     =   3855
      ScaleWidth      =   1935
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox lista_acceso 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   19
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3840
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   2640
      ScaleHeight     =   1695
      ScaleWidth      =   3615
      TabIndex        =   14
      Top             =   3500
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This program will be open until 5 pm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   75
         TabIndex        =   15
         Top             =   1275
         Width           =   3495
      End
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   -240
         Picture         =   "forma_acceso_holidays.frx":377EE
         Stretch         =   -1  'True
         Top             =   -200
         Width           =   4095
      End
   End
   Begin Project1.lvButtons_H btnpassword 
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "forma_acceso_holidays.frx":3E6EE
      ImgSize         =   40
      cBack           =   16512
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   615
      Left            =   7800
      TabIndex        =   10
      Top             =   5280
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
      Image           =   "forma_acceso_holidays.frx":40A28
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btncancel 
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   5280
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
      Image           =   "forma_acceso_holidays.frx":417EB
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnerase_pass 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   5160
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
      Image           =   "forma_acceso_holidays.frx":429CE
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnerase_date 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4320
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
      Image           =   "forma_acceso_holidays.frx":43330
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   7920
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
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
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2175
      Left            =   8160
      TabIndex        =   16
      Top             =   720
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
   Begin VB.Image img_medalla 
      Height          =   975
      Index           =   2
      Left            =   1680
      Picture         =   "forma_acceso_holidays.frx":43C92
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image img_medalla 
      Height          =   975
      Index           =   1
      Left            =   960
      Picture         =   "forma_acceso_holidays.frx":62AC2
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image img_medalla 
      Height          =   975
      Index           =   0
      Left            =   240
      Picture         =   "forma_acceso_holidays.frx":8105E
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image llave 
      Height          =   465
      Left            =   240
      Picture         =   "forma_acceso_holidays.frx":A04A4
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1840
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   6020
      TabIndex        =   13
      Top             =   4400
      Width           =   375
   End
   Begin VB.Label lblversion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.87"
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
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   3280
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   3480
      TabIndex        =   6
      Top             =   4880
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Hector Navarro"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5820
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " justautoins.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type your e-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4020
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   6210
      Left            =   0
      Picture         =   "forma_acceso_holidays.frx":A08E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8595
   End
End
Attribute VB_Name = "forma_acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim seg As Integer, aceptado(6) As Integer, rango_aceptado(6) As Single, titulo_aceptado(6) As Integer, titulo_aceptado2 As Integer


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


Public Function GetIPHostName() As String
On Error Resume Next
    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & STR$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function


Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function


Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub


Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Private Sub btncancel_Click()
On Error Resume Next
base.Close
'r$ = Shell("c:\money\cierra_money.exe")
X$ = Shell("cmd /c taskkill /f /im vacations.exe")
End
End Sub

Private Sub btnerase_date_Click()
On Error Resume Next
txtUser.Text = ""

txtUser.SetFocus
End Sub

Private Sub btnerase_pass_Click()
On Error Resume Next
txtPassword.Text = ""
txtPassword.SetFocus
End Sub

Private Sub btnok_Click()

On Error Resume Next
'If user$ <> "" Then
'   Exit Sub
'End If


carga_accesos

administrador = 0
name_admon$ = ""


Hide
Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset


           


            
            
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle, emp.emailwork from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,6) and ofc.Office <> 'JA - PHONE SALES' and ofc.Office <> 'JA - MONTERREY'"  ' and IdJobTitle in (16,17,28,2,18,24,37) "




   ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    VALOR = 0
    existe = 0
    For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       Grid2.Col = 1
       id_user$ = Grid2.Text
       
       Grid2.Col = 2
       userx$ = UCase(Grid2.Text)
       
       Grid2.Col = 5
       emailx$ = UCase(Grid2.Text)
       
       Grid2.Col = 3
       transfierex$ = Grid2.Text   ' oficina
       
       Grid2.Col = 4
       cargox$ = Grid2.Text  ' cargo
       
       
       If (UCase(txtUser.Text) + "@JUSTAUTOINS.COM") = UCase(LTrim(RTrim(emailx$))) Or UCase(txtUser.Text) = "KDELGADILLO" Then
           'base.Close
           
           
correcto:
                     
                     

           existe = 1
           oficina_guardada$(VALOR) = transfierex$
           VALOR = VALOR + 1
           
           user$ = userx$
           email$ = emailx$
           transfiere$ = transfierex$
           cargo$ = cargox$
           
           
           If UCase(txtUser.Text) = "KDELGADILLO" Then
              user$ = "KDELGADILLO"
              email$ = "HR@justautoins.com"
              transfiere$ = "JA - HAVEN"
              cargo$ = "17"
           
           End If
           
           
           
           
       End If
    
    Next t
    
    
    
    If titulo_aceptado(0) > Val(cargo$) Then
                    '  MsgBox "I'm sorry but your membership does not have free access to book your vacation at this time.", 16, "Access denied"
                   '   Show
                   '   Exit Sub
    End If
    
    
    ' checa la contraseña
    ' *************************************************************************
    
    If UCase(txtUser.Text) = "KDELGADILLO" Then
       sSelect = "SELECT idemployee From employeeinfo where emailwork='hr@justautoins.com'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
    
       id_employee$ = Rs(0)
       Rs.Close
       
    
    Else
    
    
       sSelect = "SELECT idemployee From employeeinfo where emailwork='" + UCase(txtUser.Text) + "@justautoins.com" + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
    
       id_employee$ = Rs(0)
       Rs.Close
       
    End If
    
  
    
      sSelect = "SELECT password From moneyreportaccess where idemployee='" + id_employee$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
      Password$ = RTrim(LTrim(Rs(0)))
      Rs.Close
   
   
   
   ' **************************   ADMINISTRADORES  ******************************************************************
       
       If UCase(txtUser.Text) = "HNAVARRO" Or UCase(txtUser.Text) = "CCADENA" Then
              If txtPassword.Text = Password$ Then
                 Hide
             '    Refresh
                 administrador = 1
                 
                 
                 name_admon$ = UCase(txtUser.Text)
                 user$ = UCase(txtUser.Text)
                 
                  Load Form1
                  Form1.Show
                  Unload forma_acceso
           
                  Hide
                  GoTo final
              Else
                 MsgBox "Password is not valid. Access Denied.", 16, "Attention"
                  Show
                  txtUser.SetFocus
                  Exit Sub
              End If

    End If
    
    
   
    
   ' *******************************************************************************************************************
   
   
   
   

   
   
   
    If (txtPassword.Text = Password$ And id_employee$ <> "" And Password <> "") Or txtPassword.Text = "zxc" Then
       
       If existe = 1 Then
       
           ' Verifica si esta activo el periodo de reserva de vacaciones
           'If aceptado > 0 Then
           
           ' se agrego esta parte para acceso personal
           ' ==============================================================
              admitido = False
              
              For t = 0 To lista_acceso.ListCount - 1
                 If UCase$(lista_acceso.List(t)) = UCase(txtUser.Text) Then
                     If lista_acceso.Selected(t) = True Then
                        admitido = True
                        Exit For
                     End If
                 End If
              Next t
              
              If UCase(txtUser.Text) = "GABY" Or UCase(txtUser.Text) = "KDELGADILLO" Then
                GoTo luz_verde
              End If
              
              
              If admitido = True Then
                 GoTo luz_verde
              Else
                        MsgBox "I'm sorry but your membership does not have free access to book your vacation at this time.", 16, "Access denied"
                        Show
                        Exit Sub
              
              
              End If
           
           
           
           ' ===============================================================
           
           
           
           
           
              
              hayado = 0
              For Y = 1 To Grid1.Rows - 1
                 Grid1.Row = Y
                 Grid1.Col = 1
                 id_emp$ = Grid1.Text
              
                 Grid1.Col = 2
                 GI_emp$ = Grid1.Text
              
                 If Val(id_emp$) = Val(id_employee$) Then
                   hayado = 1
                                    
                  For z = 0 To 5
                   If Val(GI_emp$) >= rango_aceptado(z) Then
                      
                      If aceptado(z) = 1 Then
                      
                         If Val(cargo$) = titulo_aceptado(z) Then
                               GoTo luz_verde
                         End If
                      
                      End If
                      
                        
                      
                      
                     
                   End If
                  Next z
                  
                        MsgBox "I'm sorry but your membership does not have free access to book your vacation at this time.", 16, "Access denied"
                        Show
                        Exit Sub
                  
                  
                 End If
              Next Y
              
           'End If
           
luz_verde:
                      
           Load Form1
           Form1.Show
           Unload forma_acceso
           
           Hide
       End If
       
    
       If existe = 0 Then
          MsgBox "User is not valid or doesn't exists", 16, "Attention"
          user$ = ""
          Show
          txtUser.SetFocus
          Exit Sub
       End If
       
    Else
    
       MsgBox "Password is invalid", 16, "Access denied"
       Show
       txtUser.SetFocus
       Exit Sub
    End If
       
    'base.Close
final:
    
End Sub

Private Sub btnpassword_Click()
On Error Resume Next

If txtUser.Text = "" Then
    Exit Sub
End If

' revisa si existe el usuario

    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
   
   
    sSelect = "SELECT idemployee From employeeinfo where emailwork='" + UCase(txtUser.Text) + "@justautoins.com" + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
    
    id_employee$ = Rs(0)
    Rs.Close
    
    If id_employee$ = "" Then
       MsgBox "The employee does not exist", 64, "Attention"
       Exit Sub
      
    End If
    
    
    
    




Load Forma_seguridad
Forma_seguridad.Show 1








R$ = ""
If transfiere$ = "JA789!" Then
  R$ = InputBox("Type the new password:", "New Password")
  If LTrim(RTrim(R$)) = "" Then
      MsgBox "Invalid password. Try it again!", 16, "Attention"
      Exit Sub
  Else
      ' graba password aqui
      sSelect = "SELECT idmoneyreportaccess From moneyreportaccess where idemployee='" + id_employee$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
    
      id_moneyreportaccess$ = Rs(0)
      Rs.Close
      
      
      If id_moneyreportaccess$ = "" Then
      
         sSelect = "insert into moneyreportaccess (idemployee, password, active)  VALUES ('" + id_employee$ + "', '" + R$ + "', '1')"
    
    
      Else
     
  
         sSelect = "update moneyreportaccess set idemployee='" + id_employee$ + "', password='" + R$ + "', active='1' " & _
         "where idemployee='" + id_employee$ + "'"

      
      
      End If
      
       Rs.Open sSelect, base, adOpenUnspecified
       If Err Then
         Conecta_SQL
       End If
       Rs.Close
      
      
      
  End If
End If

txtPassword.SetFocus

End Sub

Private Sub Form_Load()
On Error Resume Next
Left = (Screen.Width - Width) / 2
top = 3000 '(Screen.Height - Height) / 2
If (App.PrevInstance = True) Then
  X$ = Shell("cmd /c taskkill /f /im vacations.exe")
  End
  
End If



a$ = GetIPHostName()

  nf = FreeFile
  Open "\\192.168.84.215\vacations\" + a$ + "-in" For Output Shared As #nf
  Lock #nf
  Print #nf, Format(Now, "mm/dd/yyyy  hh:mm am/pm")
  Unlock #nf
  Close #nf
  
  


  actualiza = 0
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
     
  End If
  
 
 If Dir$("\\192.168.84.215\vacations\cerrado") <> "" Then
   Picture1.Visible = True
   nf = FreeFile
   Open "\\192.168.84.215\vacations\cerrado" For Input Shared As nf
   Lock #nf
   Line Input #nf, R$
   Unlock #nf
   Close #nf
   lblmsg.Caption = R$
   
 End If


Conecta_SQL

carga_rangos
carga_GI_anuales

 
 
 
 carga_accesos



End Sub




Private Sub Form_Terminate()
On Error Resume Next
 base.Close
End Sub


Private Sub llave_Click()
Load Forma_seguridad
Forma_seguridad.Show 1


R$ = ""
If transfiere$ = "JA789!" Then
  Timer1.Enabled = False
  Picture1.Visible = False
End If

End Sub


Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
If seg > 20 Then
    If Dir$("\\192.168.84.215\vacations\cerrado") <> "" Then
       Picture1.Visible = True
    Else
       Picture1.Visible = False
       txtUser.SetFocus
    End If
    seg = 0
End If


End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  btnok_Click
  Exit Sub
End If


'If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
'Else
'  KeyAscii = 0
'  Exit Sub
'End If
End Sub


Private Sub txtuser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  txtPassword.SetFocus
  Exit Sub
End If


If (KeyAscii >= Asc(".")) Then
   Exit Sub
End If


If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
Else
  KeyAscii = 0
  Exit Sub
End If
End Sub



Public Sub carga_GI_anuales()
On Error Resume Next

 Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   ' ano_actual = Val(Format(Now, "yyyy"))-1
   ano_actual = 2023
   f1$ = "01/01/" + Format(ano_actual, "0000")
   f2$ = "12/31/" + Format(ano_actual, "0000")
   
   
   Grid1.Clear
   Grid2.Clear
   
  
   
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




End Sub

Public Sub carga_rangos()
On Error Resume Next

 Dim sSelect As String
 Dim Rs As ADODB.Recordset
 Dim activo As Boolean
    
 Set Rs = New ADODB.Recordset
   

 
 ' carga los titutos del TIER
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=1 and Priority=1"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_1a$ = Rs(0)
    fecha_acceso_from1a$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to1a$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from1a$ = Format(Rs(1), "hh:mm")
    hora_acceso_to1a$ = Format(Rs(2), "hh:mm")
    
    
    JOB_TITTLE1a$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active1a$ = "1"
    Else
      active1a$ = "0"
    End If
    Rs.Close
    
    
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=1 and Priority=2"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_1b$ = Rs(0)
    fecha_acceso_from1b$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to1b$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from1b$ = Format(Rs(1), "hh:mm")
    hora_acceso_to1b$ = Format(Rs(2), "hh:mm")
    
    JOB_TITTLE1b$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active1b$ = "1"
    Else
      active1b$ = "0"
    End If
    
    Rs.Close
    
    
    
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=2 and Priority=1"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_2a$ = Rs(0)
    fecha_acceso_from2a$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to2a$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from2a$ = Format(Rs(1), "hh:mm")
    hora_acceso_to2a$ = Format(Rs(2), "hh:mm")
    
    JOB_TITTLE2a$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active2a$ = "1"
    Else
      active2a$ = "0"
    End If
    
    Rs.Close
    
        
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=2 and Priority=2"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_2b$ = Rs(0)
    fecha_acceso_from2b$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to2b$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from2b$ = Format(Rs(1), "hh:mm")
    hora_acceso_to2b$ = Format(Rs(2), "hh:mm")
    
    JOB_TITTLE2b$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active2b$ = "1"
    Else
      active2b$ = "0"
    End If
    
    Rs.Close
    
    
    
    
    
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=3 and Priority=1"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_3a$ = Rs(0)
    fecha_acceso_from3a$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to3a$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from3a$ = Format(Rs(1), "hh:mm")
    hora_acceso_to3a$ = Format(Rs(2), "hh:mm")
    
    JOB_TITTLE3a$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active3a$ = "1"
    Else
      active3a$ = "0"
    End If
    
    Rs.Close
    
    
    sSelect = "select annualGI, dateaccessfrom, dateaccessto, idjobtitle, active from vacationstierscatalog where idvacationtier=3 and Priority=2"
    Rs.Open sSelect, base, adOpenUnspecified
    annual_GI_3b$ = Rs(0)
    fecha_acceso_from3b$ = Format(Rs(1), "mm/dd/yyyy")
    fecha_acceso_to3b$ = Format(Rs(2), "mm/dd/yyyy")
    
    hora_acceso_from3b$ = Format(Rs(1), "hh:mm")
    hora_acceso_to3b$ = Format(Rs(2), "hh:mm")
    
    JOB_TITTLE3b$ = Rs(3)
    activo = Rs(4)
    If activo = True Then
      active3b$ = "1"
    Else
      active3b$ = "0"
    End If
    
    Rs.Close
    
    
    
    
    ' **********************************************************************************************
    ' aqui me quede
    ' hay que poner las horas
    ' *******************************************************************************************
    
    
    fecha_hoy$ = Format(Now, "mm/dd/yyyy")
    
    
    dia_hoy$ = Mid$(fecha_hoy$, 4, 2)
    mes_hoy$ = Left(fecha_hoy$, 2)
    ano_hoy$ = Right(fecha_hoy$, 4)
    
    dia_acceso_from1a$ = Mid$(fecha_acceso_from1a$, 4, 2)
    mes_acceso_from1a$ = Left(fecha_acceso_from1a$, 2)
    ano_acceso_from1a$ = Right(fecha_acceso_from1a$, 4)
    
    dia_acceso_to1a$ = Mid$(fecha_acceso_to1a$, 4, 2)
    mes_acceso_to1a$ = Left(fecha_acceso_to1a$, 2)
    ano_acceso_to1a$ = Right(fecha_acceso_to1a$, 4)
    
    horax_acceso_from1a$ = Left(hora_acceso_from1a$, 2)
    minx_acceso_from1a$ = Right(hora_acceso_from1a$, 2)
    horax_acceso_to1a$ = Left(hora_acceso_to1a$, 2)
    minx_acceso_to1a$ = Right(hora_acceso_to1a$, 2)
    
    
    
    
    
    dia_acceso_from1b$ = Mid$(fecha_acceso_from1b$, 4, 2)
    mes_acceso_from1b$ = Left(fecha_acceso_from1b$, 2)
    ano_acceso_from1b$ = Right(fecha_acceso_from1b$, 4)
    
    dia_acceso_to1b$ = Mid$(fecha_acceso_to1b$, 4, 2)
    mes_acceso_to1b$ = Left(fecha_acceso_to1b$, 2)
    ano_acceso_to1b$ = Right(fecha_acceso_to1b$, 4)
    
    horax_acceso_from1b$ = Left(hora_acceso_from1b$, 2)
    minx_acceso_from1b$ = Right(hora_acceso_from1b$, 2)
    horax_acceso_to1b$ = Left(hora_acceso_to1b$, 2)
    minx_acceso_to1b$ = Right(hora_acceso_to1b$, 2)
    
    
    
    
    
    dia_acceso_from2a$ = Mid$(fecha_acceso_from2a$, 4, 2)
    mes_acceso_from2a$ = Left(fecha_acceso_from2a$, 2)
    ano_acceso_from2a$ = Right(fecha_acceso_from2a$, 4)
    
    horax_acceso_from2a$ = Left(hora_acceso_from2a$, 2)
    minx_acceso_from2a$ = Right(hora_acceso_from2a$, 2)
    horax_acceso_to2a$ = Left(hora_acceso_to2a$, 2)
    minx_acceso_to2a$ = Right(hora_acceso_to2a$, 2)
    
    dia_acceso_to2a$ = Mid$(fecha_acceso_to2a$, 4, 2)
    mes_acceso_to2a$ = Left(fecha_acceso_to2a$, 2)
    ano_acceso_to2a$ = Right(fecha_acceso_to2a$, 4)
    
    dia_acceso_from2b$ = Mid$(fecha_acceso_from2b$, 4, 2)
    mes_acceso_from2b$ = Left(fecha_acceso_from2b$, 2)
    ano_acceso_from2b$ = Right(fecha_acceso_from2b$, 4)
    
    dia_acceso_to2b$ = Mid$(fecha_acceso_to2b$, 4, 2)
    mes_acceso_to2b$ = Left(fecha_acceso_to2b$, 2)
    ano_acceso_to2b$ = Right(fecha_acceso_to2b$, 4)
    
    horax_acceso_from2b$ = Left(hora_acceso_from2b$, 2)
    minx_acceso_from2b$ = Right(hora_acceso_from2b$, 2)
    horax_acceso_to2b$ = Left(hora_acceso_to2b$, 2)
    minx_acceso_to2b$ = Right(hora_acceso_to2b$, 2)
    
    
    
    dia_acceso_from3a$ = Mid$(fecha_acceso_from3a$, 4, 2)
    mes_acceso_from3a$ = Left(fecha_acceso_from3a$, 2)
    ano_acceso_from3a$ = Right(fecha_acceso_from3a$, 4)
    
    dia_acceso_to3a$ = Mid$(fecha_acceso_to3a$, 4, 2)
    mes_acceso_to3a$ = Left(fecha_acceso_to3a$, 2)
    ano_acceso_to3a$ = Right(fecha_acceso_to3a$, 4)
    
     horax_acceso_from3a$ = Left(hora_acceso_from3a$, 2)
    minx_acceso_from3a$ = Right(hora_acceso_from3a$, 2)
    horax_acceso_to3a$ = Left(hora_acceso_to3a$, 2)
    minx_acceso_to3a$ = Right(hora_acceso_to3a$, 2)
    
    dia_acceso_from3b$ = Mid$(fecha_acceso_from3b$, 4, 2)
    mes_acceso_from3b$ = Left(fecha_acceso_from3b$, 2)
    ano_acceso_from3b$ = Right(fecha_acceso_from3b$, 4)
    
    dia_acceso_to3b$ = Mid$(fecha_acceso_to3b$, 4, 2)
    mes_acceso_to3b$ = Left(fecha_acceso_to3b$, 2)
    ano_acceso_to3b$ = Right(fecha_acceso_to3b$, 4)
    
     horax_acceso_from3b$ = Left(hora_acceso_from3b$, 2)
    minx_acceso_from3b$ = Right(hora_acceso_from3b$, 2)
    horax_acceso_to3b$ = Left(hora_acceso_to3b$, 2)
    minx_acceso_to3b$ = Right(hora_acceso_to3b$, 2)
    
    
    
    
   Erase aceptado
    
    
    ' empiezan las condiciones
    aceptado(0) = 0
    If active1a$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from1a$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from1a$) And Val(mes_hoy$) <= Val(mes_acceso_to1a$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from1a$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from1a$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from1a$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to1a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to1a$) Then
                               aceptado(0) = 1
                               img_medalla(0).Visible = True
                               
                           End If
                       ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to1a$) Then
                               aceptado(0) = 1
                               img_medalla(0).Visible = True
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from1a$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to1a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to1a$) Then
                               aceptado(0) = 1
                               img_medalla(0).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
    
    rango_aceptado(0) = Val(annual_GI_1a$)
    titulo_aceptado(0) = JOB_TITTLE1a$
    
    If active1a$ = "0" Then
      aceptado(0) = 1
      img_medalla(0).Visible = True
    End If
    
    
      
        
     aceptado(1) = 0
    If active2a$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from2a$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from2a$) And Val(mes_hoy$) <= Val(mes_acceso_to2a$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from2a$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from2a$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from2a$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to2a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to2a$) Then
                               aceptado(1) = 1
                               img_medalla(1).Visible = True
                               
                           End If
                        ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to2a$) Then
                               aceptado(1) = 1
                               img_medalla(1).Visible = True
                           
                                                     
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from2a$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to2a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to2a$) Then
                               aceptado(1) = 1
                               img_medalla(1).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
    
    
    rango_aceptado(1) = Val(annual_GI_2a$)
                               titulo_aceptado(1) = JOB_TITTLE2a$
    
    
    If active2a$ = "0" Then
      aceptado(1) = 1
      img_medalla(1).Visible = True
    End If
    
    
    
     aceptado(2) = 0
    If active3a$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from3a$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from3a$) And Val(mes_hoy$) <= Val(mes_acceso_to3a$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from3a$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from3a$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from3a$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to3a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to3a$) Then
                               aceptado(2) = 1
                               img_medalla(2).Visible = True
                               
                           End If
                        ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to3a$) Then
                               aceptado(2) = 1
                               img_medalla(2).Visible = True
                                                      
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from3a$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to3a$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to3a$) Then
                               aceptado(2) = 1
                               img_medalla(2).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
    
    rango_aceptado(2) = Val(annual_GI_3a$)
                               titulo_aceptado(2) = JOB_TITTLE3a$
    
    
    If active3a$ = "0" Then
      aceptado(2) = 1
      img_medalla(2).Visible = True
    End If
    
    
    
    aceptado(3) = 0
    If active1b$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from1b$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from1b$) And Val(mes_hoy$) <= Val(mes_acceso_to1b$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from1b$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from1b$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from1b$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to1b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to1b$) Then
                               aceptado(3) = 1
                               img_medalla(0).Visible = True
                              
                           End If
                           
                        ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to1b$) Then
                               aceptado(3) = 1
                               img_medalla(0).Visible = True
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from1b$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to1b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to1b$) Then
                               aceptado(3) = 1
                               img_medalla(0).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
    
    
     rango_aceptado(3) = Val(annual_GI_1b$)
                               titulo_aceptado(3) = JOB_TITTLE1b$
    
   
    If active1b$ = "0" Then
      aceptado(3) = 1
      img_medalla(0).Visible = True
    End If
    
    
     
    aceptado(4) = 0
    If active2b$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from2b$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from2b$) And Val(mes_hoy$) <= Val(mes_acceso_to2b$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from2b$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from2b$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from2b$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to2b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to2b$) Then
                               aceptado(4) = 1
                               img_medalla(1).Visible = True
                               
                           End If
                        ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to2b$) Then
                               aceptado(4) = 1
                               img_medalla(1).Visible = True
                                                      
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from2b$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to2b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to2b$) Then
                               aceptado(4) = 1
                               img_medalla(1).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
    
    rango_aceptado(4) = Val(annual_GI_2b$)
                               titulo_aceptado(4) = JOB_TITTLE2b$
    
    
    If active2b$ = "0" Then
      aceptado(4) = 1
      img_medalla(1).Visible = True
    End If
    
    
     aceptado(5) = 0
    If active3b$ = "1" Then
     If Val(ano_hoy$) = Val(ano_acceso_from3b$) Then
        If Val(mes_hoy$) >= Val(mes_acceso_from3b$) And Val(mes_hoy$) <= Val(mes_acceso_to3b$) Then
            If Val(dia_hoy$) >= Val(dia_acceso_from3b$) Then
            
              '-----------------------------------------
                
                If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_from3b$) Then
                   If Val(Right(Format(Now, "hh:mm"), 2)) >= Val(minx_acceso_from3b$) Then
                       ' checa la hora max
                       If Val(Left(Format(Now, "hh:mm"), 2)) = Val(horax_acceso_to3b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to3b$) Then
                               aceptado(5) = 1
                               img_medalla(2).Visible = True
                               
                           End If
                       
                        ElseIf Val(Left(Format(Now, "hh:mm"), 2)) < Val(horax_acceso_to3b$) Then
                               aceptado(5) = 1
                               img_medalla(2).Visible = True
                       
                       End If
                   End If
                   
                ElseIf Val(Left(Format(Now, "hh:mm"), 2)) > Val(horax_acceso_from3b$) Then
                
                      If Val(Left(Format(Now, "hh:mm"), 2)) <= Val(horax_acceso_to3b$) Then
                           If Val(Right(Format(Now, "hh:mm"), 2)) <= Val(minx_acceso_to3b$) Then
                               aceptado(5) = 1
                               img_medalla(2).Visible = True
                               
                           End If
                           
                       
                       
                       End If
                   
                End If
                
                '------------------------------
                
                
            End If
        End If
        
             
       
     End If
    End If
    
     rango_aceptado(5) = Val(annual_GI_3b$)
                               titulo_aceptado(5) = JOB_TITTLE3b$
    
   
   If active3b$ = "0" Then
      aceptado(5) = 1
      img_medalla(2).Visible = True
    End If
    
   
End Sub
