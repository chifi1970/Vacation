Attribute VB_Name = "SQL_UW"
Global base As New ADODB.Connection
Global rsmensaje As New ADODB.Recordset

Global agente As String, tipo_acceso As Integer


Public Sub Conecta_SQL()
On Error Resume Next
   
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com" ' LAE ORIGINAL
 'server_ini$ = "ec2-54-193-90-176.us-west-1.compute.amazonaws.com" '  LAE PRUEBA



 With base
   .CursorLocation = adUseClient

    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
 
 
 
 
End Sub
