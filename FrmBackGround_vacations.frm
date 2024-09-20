VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmBackGround 
   Caption         =   "Imagen de fondo"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "FrmBackGround_vacations.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1485
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton BtnExaminar 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "FrmBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type



Private Sub BtnExaminar_Click()
    Dim OFName As OPENFILENAME
    
    With OFName
        .lStructSize = Len(OFName)
        .hwndOwner = Me.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = "Todos los archivos de imágenes" & Chr(0) & "*.gif;*.jpg;*.jpe;*.bmp;*.png" & Chr(0) & "Imágenes GIF (*.gif)" & Chr(0) & "*.gif" & Chr(0) & "Imágenes JPG (*.jpg, *.jpe)" & Chr(0) & "*.jpg;*.jpe" & Chr(0) & "Imágenes de mapas de bits (*.bmp)" & Chr(0) & "*.bmp" & Chr(0) & "Imágenes PNG (*.png)" & Chr(0) & "*.png" & Chr(0)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255

    End With

    If GetOpenFileName(OFName) Then
        Text1 = Trim$(OFName.lpstrFile)
    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    'Form3.HTML.All.tags("BODY")(0).Background = Text1
    'Unload Me
End Sub

Private Sub FrmBackGround_Click()

End Sub

