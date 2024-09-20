Attribute VB_Name = "Module1"
'Global AI(50, 6)
'Global JA(50, 6)
Global ruta$, ruta_pdf$
Global tipomsg
Global IDdisc$
Global cliente_ID$
Global poliza_buscada$
Global user$, cargo$
Global transfiere$
Global oficina_guardada$(5)
Global valido1
Global asunto$
Global bloqueado As Integer
Global usuario_DMV$
Global pantalla As Integer
Global fecha1$, fecha2$
Global seg_end As Integer
Global ubicacion_de_trabajo$
Global ID_vacaciones$
Global nota$
Global colorx As Long
Global fecha_accesible As Integer
Global reservado As Boolean
Global usuario$
Global administrador As Integer, name_admon$
Global evento As Long
Global contador As Integer
Global mes_selecto_en_calendario$
Global dia_bloqueado As Integer
Global titulo_del_dia$, dia_thanks As Integer, dia_thanks2 As Integer


Global fecha_rango1$, fecha_rango2$, op_searchx As Integer

Global id_agente$, ID_manager$, correo_agente$, correo_manager$, correo_admin$
Global Indice_del_evento As Long
Global llavero(900) As Long



Global Const vborange = &H50D0&                        '&H80FF&
Global Const vbverde = &H4000&
Global Const vbverde_claro = &HC000&
Global Const vbamarillo = &HDFFF&


Global Const blanco = &HFFFFFF
Global Const negro1 = &H595959
Global Const gris = &HE0E0E0
Global Const naranja = &H3CC9FF
Global Const azul = &HF2B48A
Global Const naranja_fuerte = &H80FF&
Global Const rojo = &HC0C0FF
Global Const rosa = &HFF00FF
Global Const rosa_claro = &HFF80FF
  
  
  
Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer ' As Long para que funcione en Windows XP con VB6
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type


Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SILENT = &H4


Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long


Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
    
    ' esto es para obtener el IP

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
' Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
   


Public Sub carga_aniversarios()
On Error Resume Next
Dim sSelect As String
   Dim Rs As ADODB.Recordset
    
   Set Rs = New ADODB.Recordset
   
   If valido1 = 777 Then
     Exit Sub
   End If
   
   
   Form1.mensaje.Visible = True
   Form1.mensaje.Refresh
   
   
   sSelect = "select CONCAT( FirstName+' ',MiddleName,+' '+LastName1) as Agent, CONVERT(varchar,HireDate,32) as Date from EmployeeInfo where active=1 and month (hiredate)='" + mes_selecto_en_calendario$ + "' and year(hiredate)<YEAR(CURRENT_TIMESTAMP)  order by day(HireDate) "
   
      ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Form1.grid4.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Form1.grid4.Cols = Rs.Fields.Count + 1
    
    
    
   Form1.grid4.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Form1.grid4.TextMatrix(0, i + 1) = Rs.Fields(i).Name
   Next

   Form1.grid4.Row = 1
   Form1.grid4.Col = 1

   ' Set range of cells in the grid
   Form1.grid4.RowSel = Form1.grid4.Rows - 1
   Form1.grid4.ColSel = Form1.grid4.Cols - 1
   Form1.grid4.clip = rsVar

   ' Reset the grid's selected range of cells
   Form1.grid4.RowSel = Form1.grid4.Row
   Form1.grid4.ColSel = Form1.grid4.Col

   Rs.Close
   
   
Form1.grid4.Cols = Form1.grid4.Cols + 1

   ' configura anchos de celdas
   
Form1.grid4.ColWidth(0) = 500
Form1.grid4.ColWidth(1) = 2550 'account
Form1.grid4.ColAlignment(1) = flexAlignLeftCenter

Form1.grid4.ColWidth(2) = 1300   ' chkref
Form1.grid4.ColAlignment(2) = flexAlignLeftCenter

Form1.grid4.ColWidth(3) = 700   ' debit
Form1.grid4.ColAlignment(3) = flexAlignCenterCenter


Form1.grid4.Row = 0

Form1.grid4.Col = 1
Form1.grid4.Text = "Agent"

Form1.grid4.Col = 2
Form1.grid4.Text = "Date"

Form1.grid4.Col = 3
Form1.grid4.Text = "Years"


Form1.grid4.FixedRows = 1
Form1.grid4.FixedCols = 1



' calcula los años que tiene
Dim fecActual As Date
Dim fecNac As Date
Dim anos, meses, dias As Long
fecActual = Now

For t = 1 To Form1.grid4.Rows - 1
   Form1.grid4.Row = t
   Form1.grid4.Col = 0
   Form1.grid4.Text = t
   
   Form1.grid4.Col = 2
   f$ = Format(Form1.grid4.Text, "mm/dd/yyyy")
   
   fecNac = CDate(f$)
   anos = DateDiff("yyyy", fecNac, fecActual)
   meses = DateDiff("m", fecNac, fecActual) - (anos * 12)
   dias = DateDiff("d", CDate(Day(fecNac) & "/" & (Month(fecActual) - IIf(Day(fecNac) >= Day(fecActual), 1, 0)) & "/" & Year(fecActual)), Now)
   R$ = Format(anos, "#0") ' & "A " & meses & "M " & dias & "D"
   
   Form1.grid4.Col = 3
   Form1.grid4.Text = R$
   
Next t
   
   


   
 Form1.mensaje.Visible = False
   
End Sub

Public Function ShellDelete(ParamArray vntFileName() As Variant) As Long

    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT

    For i = LBound(vntFileName) To UBound(vntFileName)
    sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION

 'FOF_ALLOWUNDO
    End With

    ShellDelete = SHFileOperation(SHFileOp)

End Function

Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
       On Error Resume Next
       
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For i = 0 To .Count - 1
         If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).top = .Controls(i).top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).top * SFY, _
            .Controls(i).Width * SFX, _
            .Controls(i).Height * SFY
         End If
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
End Sub

