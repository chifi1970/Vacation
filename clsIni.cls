VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

DefInt A-Z

' Apis para leer y escibir en archivos ini
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Const ENCRYPT = 1
Const DECRYPT = 2

Const clave  As String = "Clave"
Const SECCION  As String = "Clave"

'---------------------------------------------------------------------
'  -  para encriptar y desencriptar cadenas
'---------------------------------------------------------------------

Function Encriptar( _
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim N        As Integer
    Dim rtn      As String
    
    If Text = vbNullString Or UserKey = vbNullString Then
       Encriptar = vbNullString
       Exit Function
    End If
    '//Get UserKey characters
    N = Len(UserKey)
    ReDim UserKeyASCIIS(1 To N)
    For i = 1 To N
        UserKeyASCIIS(i) = Asc(Mid$(UserKey, i, 1))
    Next
        
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(Mid$(Text, i, 1))
    Next
    
    '//Encryption/Decryption
    If Action = ENCRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= N, 1, j + 1)
           Temp = TextASCIIS(i) + UserKeyASCIIS(j)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= N, 1, j + 1)
           Temp = TextASCIIS(i) - UserKeyASCIIS(j)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
    
    '//Return
    Encriptar = rtn
End Function

Function getValue(path_ini As String, SECCION As String, key As String, Optional Default As Variant = "") As String

On Error GoTo error_handler

    Dim buffer As String * 256
    Dim ret As Long

    ret = GetPrivateProfileString(SECCION, key, Default, buffer, Len(buffer), path_ini)
    getValue = Left$(buffer, ret)
    
Exit Function
error_handler:
MsgBox Err.Description, vbCritical

End Function

Function writeValue(path_ini As String, SECCION As String, key As String, VALOR As Variant) As String
On Error GoTo error_handler

    Dim ret As Long
    ret = WritePrivateProfileString(SECCION, key, VALOR, path_ini)
   
Exit Function
error_handler:
MsgBox Err.Description, vbCritical
   
End Function
