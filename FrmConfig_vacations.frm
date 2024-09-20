VERSION 5.00
Begin VB.Form FrmConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail account"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "FrmConfig_vacations.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.lvButtons_H btnsend 
      Height          =   615
      Left            =   1680
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Send email"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H btnsetup 
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "Load Config"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Text            =   "25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CheckBox chkAut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "This account requires authentication"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   5655
   End
   Begin VB.CheckBox chkSSL 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Use SSL - Some servers require this option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   6375
   End
   Begin VB.ComboBox cboServ 
      Height          =   315
      ItemData        =   "FrmConfig_vacations.frx":000C
      Left            =   5400
      List            =   "FrmConfig_vacations.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin Project1.lvButtons_H CmdAceptar 
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      Top             =   3600
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
      Image           =   "FrmConfig_vacations.frx":0031
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H CmdCancelar 
      Height          =   615
      Left            =   4680
      TabIndex        =   15
      Top             =   3600
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
      Image           =   "FrmConfig_vacations.frx":0DF4
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   1720
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1125
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server SMTP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   555
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private mcIni                       As clsIni               ' guardar los datos en un archivo .ini
Dim INI_PATH                        As String

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
      
      valido1 = 1974
      Load Form3
     
      valido1 = G

      Kill fuente$ + "nueva.htm"
      Name fuente$ + "nueva2.htm" As fuente$ + "nueva.htm"
      
      



If transfiere$ = "NO SEND" Then
  MsgBox "The message could not be sent", 64, "ERROR detected"
Else
  MsgBox "The message was sent correctly", 64, "Attention"
End If

transfiere$ = ""

End Sub



Private Sub btnsend_Click()
On Error Resume Next
Form1.envia_correo2
End Sub

Private Sub btnsetup_Click()
On Error Resume Next

txtServer.Text = "smtp.office365.com"
txtPort.Text = "25"
chkAut.Value = 1
chkSSL.Value = 1
txtUser.Text = "Vacations@justautoins.com"
txtPassword.Text = "Ju5t4ut02023!"

End Sub

Private Sub CmdAceptar_Click()
On Error Resume Next
    actualiza_registro


    With mcIni
        Call .writeValue(INI_PATH, "datos", "servicio Mail", cboServ.ListIndex)
        Call .writeValue(INI_PATH, "datos", "servidor", txtServer.Text)
        Call .writeValue(INI_PATH, "datos", "usuario", txtUser.Text)
        Call .writeValue(INI_PATH, "datos", "password", .Encriptar(App.EXEName, txtPassword.Text, 1))
        Call .writeValue(INI_PATH, "datos", "puerto", txtPort.Text)
        Call .writeValue(INI_PATH, "datos", "ssl", chkSSL.Value)
        Call .writeValue(INI_PATH, "datos", "Aut", chkAut.Value)
    End With
    
    
    
    
    
    'base.Close
    
    If transfiere$ = "888" Then
      Hide
      'base.Close
      Forma_main.Show
     
      Unload Me
      Exit Sub
    End If
    
    Forma_main.Show
    Unload Me
End Sub

Public Sub actualiza_registro()
On Error Resume Next
  
 
 

    ' Para la cadena de selección
    Dim sSelect As String
    Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
    Dim Rs As ADODB.Recordset
    
    
    
    ' DETECTA SI ESTA CREADO EL REGISTRO

    Set Rs = New ADODB.Recordset

   
    sSelect = "SELECT server From emailservernotification where origen='Vacations'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    server$ = Rs(0)
    
                         
    Rs.Close
    
    
    R$ = ""
    For t = 1 To Len(txtPassword.Text)
       R$ = R$ + Chr$(Asc(Mid$(txtPassword.Text, t, 1)) + 15)
    Next t
    
    Password$ = R$
       
       Set Rs = New ADODB.Recordset
    
    If server$ <> "" Then
                             
       sSelect = "update emailservernotification set origen='Vacations', service='2', server='" + txtServer.Text + "', port='" + txtPort.Text + "', aut='" + Format(chkAut.Value, "0") + _
    "', SSL='" + Format(chkSSL.Value, "0") + "', username='" + LTrim(RTrim(txtUser.Text)) + "', password='" + Password$ + "' where origen='Vacations'"
    
    Else
    
    sSelect = "INSERT INTO emailservernotification (origen, service, server, port, aut, ssl, username, password)  VALUES ('Vacations" + _
    "', '2', '" + txtServer.Text + "', '" + txtPort.Text + "', '" + Format(chkAut.Value, "0") + "', '" + Format(chkSSL.Value, "0") + _
    "', '" + LTrim(RTrim(txtUser.Text)) + "', '" + Password$ + "')"
    
    End If
    
    
       Rs.Open sSelect, base, adOpenUnspecified
    
       Rs.Close
    
    
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
Forma_main.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
'Me.BackColor = Form1.BackColor
'chkAut.BackColor = Form1.BackColor
'chkSSL.BackColor = Form1.BackColor


FrmConfig.top = 0
FrmConfig.Left = (Screen.Width - Width) / 2

Me.top = 0

If transfiere$ = "888" Then Exit Sub






 Dim sSelect As String
 Dim ultimo_id As Integer 'Long
    
    ' El recordset para acceder a los datos
 Dim Rs As ADODB.Recordset
    
   
    
cboServ.ListIndex = 2
    
    
 
    Set Rs = New ADODB.Recordset
  
     sSelect = "SELECT server From emailservernotification where origen='Vacations'"
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    server$ = Rs(0)
                         
    Rs.Close
    
    txtServer.Text = server$





    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT port From emailservernotification where origen='Vacations'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    puerto$ = Rs(0)
                         
    Rs.Close

    txtPort.Text = puerto$
         
         
         
         
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT username From emailservernotification where origen='Vacations'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    usuario$ = Rs(0)
                         
    Rs.Close
         
    txtUser.Text = usuario$
    
    
    
    
    
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT password From emailservernotification where origen='Vacations'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    Password$ = Rs(0)
                         
    Rs.Close
    
    
    
    
    R$ = ""
    For t = 1 To Len(Password$)
       R$ = R$ + Chr$(Asc(Mid$(Password$, t, 1)) - 15)
    Next t
    
    Password$ = R$
    txtPassword.Text = Password$
    
    
    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT aut From emailservernotification where origen='Vacations'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    auten$ = Rs(0)
                         
    Rs.Close
    
    
       
         chkAut.Value = Val(auten$)




    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT ssl From emailservernotification where origen='Vacations'"

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ssl$ = Rs(0)
                         
    Rs.Close
    
     chkSSL.Value = Val(ssl$)



    Set Rs = New ADODB.Recordset
  
    sSelect = "SELECT service From emailservernotification where origen='Vacations'"
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    servicio$ = Rs(0)
                         
    Rs.Close
    
     cboServ.ListIndex = Val(servicio$)





    Set mcIni = New clsIni
    ' INI_PATH = App.Path & "\config.ini"
   INI_PATH = "c:\vacations\config.ini"



cboServ.ListIndex = 2
cboServ.Enabled = False
cboServ.Visible = False


    With mcIni
        Call .writeValue(INI_PATH, "datos", "servicio Mail", cboServ.ListIndex)
        Call .writeValue(INI_PATH, "datos", "servidor", txtServer.Text)
        Call .writeValue(INI_PATH, "datos", "usuario", txtUser.Text)
        Call .writeValue(INI_PATH, "datos", "password", .Encriptar(App.EXEName, txtPassword.Text, 1))
        Call .writeValue(INI_PATH, "datos", "puerto", txtPort.Text)
        Call .writeValue(INI_PATH, "datos", "ssl", chkSSL.Value)
        Call .writeValue(INI_PATH, "datos", "Aut", chkAut.Value)
    End With






    
    With mcIni
         cboServ.ListIndex = .getValue(INI_PATH, "datos", "servicio Mail", -1)
         txtServer.Text = .getValue(INI_PATH, "datos", "servidor")
         txtPort.Text = .getValue(INI_PATH, "datos", "puerto")
         txtUser.Text = .getValue(INI_PATH, "datos", "usuario")
         txtPassword.Text = .Encriptar(App.EXEName, .getValue(INI_PATH, "datos", "password"), 2)
         chkSSL.Value = .getValue(INI_PATH, "datos", "ssl", 0)
         chkAut.Value = .getValue(INI_PATH, "datos", "Aut", 0)
     End With
     
     
     If transfiere$ = "777" Then
        CmdAceptar_Click
        
     End If
     
     
   
     
     
End Sub



Private Sub cboServ_Click()
    Dim idMail  As String
    
    
    If Me.Visible = False Then
        Exit Sub
    End If
    
    With cboServ
        Select Case .ListIndex
            ' Yahoo
            Case 0
                txtPort.Text = "465"
                chkAut.Value = 1
                chkSSL.Value = 1
                txtServer.Text = "smtp.mail.yahoo.com"
                
                idMail = InputBox("Ingrese el Id de su cuenta de yahoo. Por ejemplo si su cuenta es 'maria@yahoo.com', puede ingresar 'maria@yahoo.com' , o solo 'maria'")
                
                If idMail <> "" Then
                    'txtFrom.Text = idMail
                    txtUser.Text = idMail
                    MsgBox "Para poder utilizar el acceso pop y Smtp de Yahoo, deberá estar activada la opción 'Acceso web y Pop', desde las opciones generales de la cuenta de Yahoo", vbInformation
                End If
            ' Gmail
            Case 1
                txtPort.Text = "465"
                chkAut.Value = 1
                chkSSL.Value = 1
                txtServer.Text = "smtp.gmail.com"
                
                idMail = InputBox("Ingrese el Id de su cuenta de Gmail. Por ejemplo si su cuenta es 'maria@gmail.com', puede ingresar 'maria@gmail.com' , o solo 'maria'")
                If idMail <> "" Then
                    'txtFrom.Text = idMail
                    txtUser.Text = idMail
                End If
            ' otro
            Case 2
                chkAut.Value = 1
                chkSSL.Value = 0
                txtServer.Text = ""
                txtPort.Text = "25"
                txtPassword.Text = ""
                'txtFrom.Text = ""
                txtUser.Text = ""
        End Select
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mcIni = Nothing
End Sub

Private Sub Image1_Click()
On Error Resume Next
txtPassword.PasswordChar = ""
End Sub


Private Sub Image2_Click()

End Sub

Private Sub img_mail_down_Click()

End Sub


