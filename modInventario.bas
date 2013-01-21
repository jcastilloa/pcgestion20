Attribute VB_Name = "modInventario"
'---------------------------------------------------------------------------------------
' Modulo      : modInventario
' Fecha/Hora  : 06/03/2004 20:39
' Autor       : JCASTILLO
' Propósito   : Rutinas para hacer el inventario de almacen
'---------------------------------------------------------------------------------------
Option Explicit

'directorio para ir guardando las bases de datos de inventarios.
Const dir_inventario = "c:\INVENTARIOS\"

'---------------------------------------------------------------------------------------
' Subrutina   : CreateDatabaseInventario
' Fecha/Hora  : 06/03/2004 21:38
' Autor       : JCASTILLO
' Propósito   : Devuelve el nombre del fichero creado,
'---------------------------------------------------------------------------------------
Private Function CreateDatabaseInventario() As String
'On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(7) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Dim fichero_inventario As String

fichero_inventario = "Inventario" & Format(Now, "dd-mm-yy-hh-mm-ss") & ".inv"

sCnn = strCnnMdb & dir_inventario & fichero_inventario

Cat.Create sCnn

  '----------* Table Definition of CONF_INVEN *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "CONF_INVEN"
    .Columns.Append "CODALM", adInteger
      .Columns("CODALM").Properties("Default").Value = "0"
    .Columns.Append "CODUSR", adInteger
      .Columns("CODUSR").Properties("Default").Value = "0"
    .Columns.Append "FMODI", adDate
      .Columns("FMODI").Properties("Default").Value = "now()"
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    .Columns.Append "IMP_A", adCurrency
      .Columns("IMP_A").Properties("Default").Value = "0"
    .Columns.Append "IMP_B", adCurrency
      .Columns("IMP_B").Properties("Default").Value = "0"
    .Columns.Append "IMP_TOT", adCurrency
      .Columns("IMP_TOT").Properties("Default").Value = "0"
    .Columns.Append "NUMPREN", adInteger
      .Columns("NUMPREN").Properties("Default").Value = "0"
  End With
  '----------* Index Definitions of CONF_INVEN *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Tbl(0).Indexes.Append Idx(0)

  Cat.tables.Append Tbl(0)

  '----------* Table Definition of INVENTARIO *----------
  Set Tbl(1) = New ADOX.Table
  Tbl(1).ParentCatalog = Cat
  With Tbl(1)
    .Name = "INVENTARIO"
    .Columns.Append "CASILLA", adInteger
      .Columns("CASILLA").Properties("Default").Value = "0"
    .Columns.Append "CODART", adInteger
      .Columns("CODART").Properties("Default").Value = "0"
    .Columns.Append "CODCOL", adSmallInt
      .Columns("CODCOL").Properties("Default").Value = "0"
    .Columns.Append "CODTALLA", adSmallInt
      .Columns("CODTALLA").Properties("Default").Value = "0"
    .Columns.Append "ESTANTE", adSmallInt
      .Columns("ESTANTE").Properties("Default").Value = "0"
    .Columns.Append "FMODI", adDate
      .Columns("FMODI").Properties("Default").Value = "now()"
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    .Columns.Append "PERCHERO", adInteger
      .Columns("PERCHERO").Properties("Default").Value = "0"
    .Columns.Append "TEMPOR", adUnsignedTinyInt
      .Columns("TEMPOR").Properties("Default").Value = "0"
  End With
  '----------* Index Definitions of INVENTARIO *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Tbl(1).Indexes.Append Idx(0)

  Cat.tables.Append Tbl(1)

  Set Cat = Nothing
  
  CreateDatabaseInventario = dir_inventario & fichero_inventario
  
  Exit Function

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Code Gen Error")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Function
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select
    
    'devolver error
    CreateDatabaseInventario = "@"

End Function


'---------------------------------------------------------------------------------------
' Subrutina   : inicia_inventario
' Fecha/Hora  : 06/03/2004 21:37
' Autor       : JCASTILLO
' Propósito   : Crea la base de datos de inventario y devuelve el nombre la db creada para
'               trabajar posteriormente con ella. Devuelve @ si hubo algun error
'---------------------------------------------------------------------------------------
Public Function inicia_inventario() As String

'primero crear la base de datos
   On Error GoTo inicia_inventario_Error

 inicia_inventario = CreateDatabaseInventario

   On Error GoTo 0
   Exit Function

inicia_inventario_Error:

    inicia_inventario = "@"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento inicia_inventario de Módulo modInventario"

End Function
