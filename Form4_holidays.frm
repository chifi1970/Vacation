VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Block day"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4_holidays.frx":0000
   ScaleHeight     =   5505
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   4875
      Begin VB.Frame marco_dias 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame marco_dias 
         Height          =   1215
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPEndDate 
            Height          =   375
            Left            =   120
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   240
            Width           =   825
         End
      End
      Begin MSComCtl2.DTPicker DTPStartTime 
         Height          =   375
         Left            =   480
         TabIndex        =   10
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
         TabIndex        =   11
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
   Begin Project1.lvButtons_H btnunblock 
      Height          =   1180
      Left            =   1380
      TabIndex        =   4
      Top             =   3100
      Width           =   920
      _ExtentX        =   1614
      _ExtentY        =   2090
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
      ImgAlign        =   4
      Image           =   "Form4_holidays.frx":5CAA
      ImgSize         =   40
      cBack           =   33023
   End
   Begin Project1.lvButtons_H btnbloquear 
      Height          =   1000
      Left            =   2720
      TabIndex        =   1
      Top             =   2980
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1773
      CapAlign        =   2
      BackStyle       =   7
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
      ImgAlign        =   4
      Image           =   "Form4_holidays.frx":6BF3
      ImgSize         =   40
      cBack           =   49344
   End
   Begin Project1.lvButtons_H CmdCancel 
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   4800
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
      Image           =   "Form4_holidays.frx":7823
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12632319
      CalendarTitleBackColor=   12632319
      Format          =   213778433
      CurrentDate     =   44676
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   390
      Left            =   -360
      TabIndex        =   12
      Top             =   5040
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
   End
   Begin VB.Label lblid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   600
      TabIndex        =   14
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Start date:"
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
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EventKey As Long, id_empleado$, seg As Integer
Private Sub btnbloquear_Click()
On Error Resume Next

   Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
     
     
    
   ' verifica que no este esa fecha ya grabada
   
   
   existe = 0
   ID_vacaciones$ = ""
   sSelect = "select idvacation from vacationsprogram where daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "'  and active='1'"      ' + "' and idemployee='" + datos(0) + "'"
   Rs.Open sSelect, base, adOpenUnspecified
   If Err Then
        Conecta_SQL
   End If
   ID_vacaciones$ = Rs(0)
   Rs.Close
      
   
   
   
   
      ' obtiene el ID de GABY
      sSelect = "select idemployee from employeeinfo where username='gjimenez'"
      Rs.Open sSelect, base, adOpenUnspecified
      If Err Then
        Conecta_SQL
      End If
      ID_gaby$ = Rs(0)
      Rs.Close
      
      
         
      aprobado_status = True
      
   
   
   NUEVO = 0
   
   
   
   
   
   If ID_vacaciones$ = "" Then
      ' graba registro
      
      
       sSelect = "insert into vacationsprogram (idemployee, idmanager, approvedby, daterequested, hours, approved, notes, datecreated, lastupdated, active) " & _
       "VALUES ('" + ID_gaby$ + "', '" + ID_gaby$ + "', '" + ID_gaby$ + "' , '" + Format(DTPStartDate, "mm/dd/yyyy") + "', '8', '1', '" & _
       "BLOCK DAY" + "', '" + Format(Now, "mm/dd/yyyy hh:mm") + "', '" + Format(Now, "mm/dd/yyyy hh:mm") + "', '1')"
       
       PicColor.BackColor = vbRed
       colorx = vbRed
       'transfiere$ = "ENVIA CORREO1"
       transfiere$ = "NONE"
       
       NUEVO = 1
           
  Else
       
        Exit Sub
        
       
        sSelect = "update vacationsprogram set idemployee='" + ID_gaby$ + "', idmanager='" + ID_gaby$ + "', approvedby=" + ID_gaby$ + ", " & _
        "daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "', hours='8', approved='1', " & _
         "notes='BLOCK DAY', lastupdated='" + Format(Now, "mm/dd/yyyy hh:mm") + "', active='1' " & _
        "where idvacation='" + ID_vacaciones$ + "'"
       
        
         PicColor.BackColor = vbRed
         colorx = vbRed
            
         transfiere$ = "NONE"
        
       
  End If
  
  
  Rs.Open sSelect, base, adOpenUnspecified
   
    If Err Then
      Conecta_SQL
    End If
    
   
   Rs.Close
   
  
  
  
  ' actualiza el ID de vacaciones
   
   If ID_vacaciones$ = "" Then
    
     sSelect = "select idvacation from vacationsprogram where daterequested='" + Format(DTPStartDate, "mm/dd/yyyy") + "' and idemployee='" + ID_gaby$ + "' and idmanager='" + ID_gaby$ + "'  and active='1'"
   
     Rs.Open sSelect, base, adOpenUnspecified
   
     If Err Then
       Conecta_SQL
     End If
    
    ID_vacaciones$ = RTrim(LTrim(Rs(0)))
    Rs.Close
   
   End If
  
   
   
  


    Dim StartDate As Date
    Dim EndDate As Date
    
    StartDate = Format(DateValue(DTPStartDate.Value), "mm/dd/yy") & " 00:00"
    EndDate = Format(DateValue(DTPEndDate.Value), "mm/dd/yy") & " 23:59"
    
    f1$ = StartDate & " 00:00:00 AM"
    f2$ = EndDate
    
    
                        
    
    
      ubicacion_de_trabajo$ = "JA-HAVEN"
      
      Form1.ucCalendar1.AddEvents "GABRIELA JIMENEZ", StartDate, EndDate, vbRed, True, "BLOCK DAY", , False, False, True, , "JA-HAVEN"
     
         
    
  
   
    fecha_accesible = 0
    
    Unload Me
End Sub

Private Sub btnunblock_Click()
 On Error Resume Next
    
    
    Dim sSelect As String
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
    If ID_vacaciones$ = "" Then
       Exit Sub
    End If
    
    
    R$ = MsgBox("Are you sure you want to delete this event?", 4, "Attention")
    If R$ = "7" Then Exit Sub
    
    
    
    Form1.ucCalendar1.RemoveEvent evento  ' EventKey
       
    
    
   
   
    ' borra la reservacion del dia de vacaciones
    'sSelect = "delete from VacationsProgram where idvacation='" + ID_vacaciones$ + "'"
    sSelect = "update VacationsProgram set active='0' where idvacation='" + ID_vacaciones$ + "'"
      
   
    Rs.Open sSelect, base, adOpenUnspecified
    If Err Then
      Conecta_SQL
    End If
    
     Rs.Close
      Unload Me
      
End Sub

Private Sub CmdCancel_Click()
On Error Resume Next


Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
top = 0
Left = (Screen.Width - Width) / 2
Form4.Visible = False



Verifica_existencia


If id_empleado$ <> "76" Then
     
     Unload Me
End If
        

End Sub



Public Sub Verifica_existencia()
On Error Resume Next

Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
     
     
   
     
     
   If transfiere$ = "" Then
     transfiere$ = DTPStartDate.Value
   End If
    
   ' verifica que no este esa fecha ya grabada
   
   
   existe = 0
   ID_vacaciones$ = ""
   id_empleado$ = ""
   sSelect = "select idvacation, idemployee from vacationsprogram where daterequested='" + transfiere$ + "'  and active='1'"
   Rs.Open sSelect, base, adOpenUnspecified
   If Err Then
        Conecta_SQL
   End If
   ID_vacaciones$ = Rs(0)
   id_empleado$ = Rs(1)
   Rs.Close
   
   
   
   
   
   lblid.Caption = ID_vacaciones$
   
   If ID_vacaciones$ = "" Then
     id_empleado$ = "76"
     btnunblock.Enabled = False
     
   Else
     btnbloquear.Enabled = False
   End If
   
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
If seg >= 1 Then
 
 If Val(Format(DTPStartDate.Value, "yyyy")) < 2000 Then
   
   Unload Me
   Exit Sub
 Else
   Form4.Visible = True
   Form4.Refresh
   
 End If
 Timer1.Enabled = False
 
End If
End Sub


