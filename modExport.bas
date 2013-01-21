Attribute VB_Name = "modExport"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Enum CZErrors
[Insufficient Buffer] = -5
End Enum

'Custom compressed file header
Type CompressionHeader
    OriginalExt As String * 3
    OriginalSize As Long
End Type
'Actual header variable
Dim FileHeader As CompressionHeader
'Used to compare compression ratios
Dim OriginalSize As Long, CompressedSize As Long

Public Sub CompressFile(ByVal SrcFilename As String, DstFilename As String)
'Used to strip the extension from the original filename
Dim fExtension As String * 3
   On Error GoTo CompressFile_Error

fExtension = Right(SrcFilename, 3)

'Allocate an array to receive the data from a file
Dim DataBytes() As Byte
ReDim DataBytes(FileLen(SrcFilename) - 1)

'Copy the data from the source into a numerical array
Open SrcFilename For Binary Access Read As #1
    Get #1, , DataBytes()
Close #1

'Track the original size
OriginalSize = UBound(DataBytes) + 1

'Allocate memory for a temporary compression array
Dim BufferSize As Long
Dim TempBuffer() As Byte
BufferSize = UBound(DataBytes) + 1
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'Compress the data using zLib
Dim result As Long
result = compress(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1)

'Copy the compressed data back into our first array
ReDim DataBytes(BufferSize - 1)
CopyMemory DataBytes(0), TempBuffer(0), BufferSize

'Kill the now useless buffer
Erase TempBuffer

'Some very simple error handling
If result = 0 Then
    CompressedSize = UBound(DataBytes) + 1
Else
    Err.Raise 1, "CompressFile", "Se produjo un error al comprimir " & SrcFilename
    Exit Sub
End If

On Error Resume Next
Kill DstFilename

'Build our custom compressed file header
FileHeader.OriginalExt = fExtension
FileHeader.OriginalSize = OriginalSize
'Write the header and then the compressed data
Open DstFilename For Binary Access Write As #1
    Put #1, 1, FileHeader
    Put #1, , DataBytes()
Close #1

'Kill the now unnecessary compressed data
Erase DataBytes

   On Error GoTo 0
   Exit Sub

CompressFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CompressFile de Módulo modExport"

End Sub

Public Sub DecompressFile(ByVal SrcFilename As String, Optional ByRef DstFilename As String = 0)
'Allocate a temporary array for receiving the compressed data
Dim DataBytes() As Byte
Dim result As Long
   On Error GoTo DecompressFile_Error

ReDim DataBytes(FileLen(SrcFilename) - Len(FileHeader) - 1)
'Copy out the header and then the compressed data
Open SrcFilename For Binary Access Read As #1
    Get #1, 1, FileHeader
    Get #1, , DataBytes()
Close #1

'Get the compressed size
OriginalSize = UBound(DataBytes) + 1

'Allocate memory for buffers
Dim BufferSize As Long
Dim TempBuffer() As Byte
BufferSize = FileHeader.OriginalSize
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'Decompress the data using zLib
result = uncompress(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1)

'Copy the uncompressed data back into our first array
ReDim DataBytes(BufferSize - 1)
CopyMemory DataBytes(0), TempBuffer(0), BufferSize

'Some very simple error handling
If result = 0 Then
    CompressedSize = UBound(DataBytes) + 1
Else
    Err.Raise 2, "DeCompressFile", "Se produjo un error al descomprimir " & SrcFilename
    Exit Sub
End If

'Kill the now unnecessary buffer
Erase TempBuffer

'Build the output path using the original filename
If Len(DstFilename) = 0 Then
    DstFilename = Left(SrcFilename, Len(SrcFilename) - 3)
    DstFilename = DstFilename & FileHeader.OriginalExt
End If
On Error Resume Next
Kill DstFilename

'Write the uncompressed data back into its original format
Open DstFilename For Binary Access Write As #1
    Put #1, , DataBytes()
Close #1

'Kill the now unnecessary data array
Erase DataBytes

   On Error GoTo 0
   Exit Sub

DecompressFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento DecompressFile de Módulo modExport"

End Sub

Public Function ExportDB(srcFile As String, Optional zipFile As String = Empty) As String
Dim tables(2, 48) As String
Dim i As Integer
Dim j As Integer
Dim Count As Long
Dim nFile As Long
Dim strFields As String

'srcFile = App.Path + "\export.dat"
If Len(zipFile) = 0 Then
    zipFile = App.path + "\export.zip"
    'Kill zipFile
End If
tables(0, 0) = "CENTROS"
tables(1, 0) = "SELECT CODIGO, DESCRIPCION, DIRECC, PROVIN, LOCALI, CODPOS, TELFNO, FAX, COMENTARIO, MBAJA, FALTA, FMODI , ROWGUID FROM CENTROS"
tables(0, 1) = "ALMACENES"
tables(1, 1) = "SELECT CODIGO,CODCEN,DESCRIPCION,UBICACION,TELEFONO,MBAJA,FBAJA, FALTA, FMODI , ROWGUID FROM ALMACENES"
tables(0, 2) = "BANCOS"
tables(1, 2) = "SELECT CODIGO,NOMBRE, DOMICILIO, LOCALIDAD, PROVINCIA, CODPOS, TELEFONO, FAX, PERCON, TELEFPCON, MBAJA, FALTA, FMODI , ROWGUID FROM BANCOS"
tables(0, 3) = "CAJAS"
tables(1, 3) = "SELECT CODIGO, CODCEN, CODALM, DESCRIPCION, UBICACION, TELEFONO, SALDOINI, CAJA_A, CAJA_B, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM CAJAS"
tables(0, 4) = "CATTALL"
tables(1, 4) = "SELECT CODIGO, DESCRIPCION, MBAJA,FBAJA, FALTA, FMODI , ROWGUID FROM CATTALL"
tables(0, 5) = "CLIENTES"
tables(1, 5) = "SELECT CODIGO, RAZO, TITULAR, DIRECCION, CODPOS, POBLACION, PROVINCIA, PAIS, TELEFONO1, TELEFONO2, FAX, EMAIL, NIF, WEB, PERCONTA, TELCONTA, REPRESEN, COMENTARIO, DCTO, DCTOPP, IMPUESTOS, FCOBRO, DIAPAGO1, DIAPAGO2, CODBAN, ENTIDAD, SUCURSAL, DC, CUENTA, DIRECENVIO, CPENVIO, PROVENVIO, PAISENVIO, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM CLIENTES"
tables(0, 6) = "COLORES"
tables(1, 6) = "SELECT CODIGO, DESCRIPCION, CODCOL, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM COLORES"
tables(0, 7) = "DCTOS"
tables(1, 7) = "SELECT C11,C12,C13,C14,C15,C21,C22,C23,C24,C25,C31,C32,C33,C34,C35,C41,C42,C43,C44,C45,C51,C52,C53,C54,C55, C00, C01, C02, C03, C04, C05, C06, C07, C08, C09, C10, C16, C17, C18, C19, C20, C26, C27, C28, C29, C30, C46, C47, C48, C49, C50, C56, C57, C58, C59, C60, C61, C62, C63, C64, C65, C66, C67, C68, C69, C70, C71, C72, C73, C74, C75, C76, C77, C78, C79, C80, C81, C82, C83, C84, C85, C86, C87, C88, C89, C90, C91, C92, C93, C94, C95, C96, C97, C98, C99, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM DCTOS"
tables(0, 8) = "TEMPOR"
tables(1, 8) = "SELECT IDTEM, AÑO, TEMPORADA, ABREVIA, ACTUAL, HIST, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM TEMPOR"
tables(0, 9) = "DETPEDPRO"
tables(1, 9) = "SELECT NUMERO, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, DCTO, IVA, RE, METIDO, FMODI , ROWGUID FROM DETPEDPRO"
tables(0, 10) = "DEUDCLI"
tables(1, 10) = "SELECT CODIGO, CODCAJA, CODPER, CODCLI, CAJACLI, CODVEN, IMPORTE, FACTURA, FECHA, ESTADO, DESCRIPCION, MBAJA, FALTA, FMODI, COMENTARIO , ROWGUID FROM DEUDCLI"
tables(0, 11) = "FAMILIAS"
tables(1, 11) = "SELECT CODIGO, CODSEC, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM FAMILIAS"
tables(0, 12) = "FCOBRO"
tables(1, 12) = "SELECT CODIGO, DESCRIPCION, PRIMERA, SEGUNDA, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM FCOBRO"
tables(0, 13) = "FPAGO"
tables(1, 13) = "SELECT CODIGO, DESCRIPCION, DIAS, MBAJA, FBAJA, FMODI , ROWGUID FROM FPAGO"
tables(0, 14) = "COSTURE"
tables(1, 14) = "SELECT CODIGO, NOMBRE, DIRECCION, CODPOS, POBLACION, PROVINCIA, PAIS, TELEFONO1, TELEFONO2, FAX, EMAIL, NIF, FPAGO, CODBAN, ENTIDAD, SUCURSAL, DC, CUENTA, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM COSTURE"
tables(0, 15) = "IVA"
tables(1, 15) = "SELECT CODIGO, IVA, RE, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM IVA"
tables(0, 16) = "MAARTIC"
tables(1, 16) = "SELECT CODIGO, SECCION, FAMILIA, SUBFAM, MODELO, ABREVIA, REF, PRECOM, DCTO, PREVEN, TARIFA, STOCK, STOCKMIN, STOCKMAX, PEDIR, TIPOIVA, CODPROV, TEMPOR, HIST, COMEN, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM MAARTIC"
tables(0, 17) = "MAING"
tables(1, 17) = "SELECT CODIGO, DESCRIPCION, COMENTARIO, MBAJA, FBAJA, FMODI , ROWGUID FROM MAING"
tables(0, 18) = "MAPAG"
tables(1, 18) = "SELECT CODIGO, DESCRIPCION, COMENTARIO, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM MAPAG"
tables(0, 19) = "MAPROV"
tables(1, 19) = "SELECT CODIGO, CIF, NOMBRE, SECTOR, DIRECC, PROVIN, LOCALI, CODPOS, TELFNO, PERCON1, PERCON2, FAX, DCTO, DCTOPP, RE, FPAGO, EXENTO, CODBAN, CCENTI, CCOFICI, CCDC, CCCUEN, COMEN, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM MAPROV"
tables(0, 20) = "PDEUDCLI"
tables(1, 20) = "SELECT CODIGO, CODDEU, CODPER, CODCLI, CODVEN, IMPORTE, CODCAJA, FACTURA, FECHA, DESCRIPCION, FMODI , ROWGUID FROM PDEUDCLI"
tables(0, 21) = "PERSONAL"
tables(1, 21) = "SELECT CODIGO, NOMBRE, CODCEN, CODCAJA, DIRECCION, CODPOS, POBLACION, PROVINCIA, PAIS, TELEFONO1, TELEFONO2, FAX, EMAIL, NIF, AFILIA, ANTIGU, CODBAN, ENTIDAD, SUCURSAL, DC, CUENTA, FPAGO, COMENTARIO, CLAVE, MBAJA, FBAJA, FALTA, TIPPERM, FMODI , ROWGUID FROM PERSONAL"
tables(0, 22) = "INGRESOS"
tables(1, 22) = "SELECT CODIGO, TIPOING, CODPER, IMPORTE, IVA, CODCAJA, FACTURA, FECHA, ESTADO, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI, COMENTARIO , ROWGUID FROM INGRESOS"
tables(0, 23) = "ARREGLOS"
tables(1, 23) = "SELECT ID, CODCOST,CODART, TEMPOR, CODVEN, DESCRIPCION, COSTE, PVP, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM ARREGLOS"
tables(0, 24) = "PLAZOE"
tables(1, 24) = "SELECT CODIGO, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM PLAZOE"
tables(0, 25) = "CABPEDPRO"
tables(1, 25) = "SELECT NUMERO, FECHA, CODPROV, CODALM, CODUSR, COMENTARIO, FPAGO,ESTADO, TRNSPORTI, DCTOPP, GASTOS, PLAZOE, PORTES, MBAJA, FBAJA, TOTALIVA, TOTALNET, SUCODIGO, ALBARAN, FACTURA, FMODI , ROWGUID FROM CABPEDPRO"
tables(0, 26) = "CABVENTA"
tables(1, 26) = "SELECT CODIGO, CODPER, CODCAJA, CODCLI, CAJADES, SUBTOT, IVATOT, RETOT, IMP_PRIMERA, IMP_SEGUNDA, ESTADO, FCOBRO, COMEN, FHORA, FMODI , ROWGUID FROM CABVENTA"
tables(0, 27) = "DETVENTA"
tables(1, 27) = "SELECT CODVEN, CODCAJA, LINEA, CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES, PREVEN, DCTO, IVA, RE, OFERTA , ROWGUID, FMODI FROM DETVENTA"
tables(0, 28) = "PTRANS"
tables(1, 28) = "SELECT CODIGO, CODALMORIG, CODALMDEST, ESTADO, CODUSR, FMODI , ROWGUID FROM PTRANS"
tables(0, 29) = "PTRANSMSG"
tables(1, 29) = "SELECT ID, CODIGO, CODALMORIG, CODUSR, CODALM, MSG, FMODI , ROWGUID FROM PTRANSMSG"
tables(0, 30) = "RCABPEDPRO"
tables(1, 30) = "SELECT NUMERO,ALMORIG,FECHA, CODPROV, CODALM, CODUSR, COMENTARIO, FPAGO, ESTADO, TRNSPORTI, DCTOPP, GASTOS, PLAZOE, PORTES, MBAJA, FBAJA, FALTA, TOTALIVA, TOTALNET, CODPTRN, SUCODIGO, ALBARAN,FACTURA, FMODI,DESTINO , ROWGUID FROM RCABPEDPRO"
tables(0, 31) = "PAGOS"
tables(1, 31) = "SELECT CODIGO, CODCAJA,TIPOPAGO, CODPROV, CODPER, IMPORTE, MENSUAL, IVA,  NUMPED, FACTURA, FPAGO,ESTADO,  DESCRIPCION, CODBAN, ENTIDAD, SUCURSAL, DC, CUENTA, MBAJA, FBAJA, FMODI, COMENTARIO , ROWGUID FROM PAGOS"
tables(0, 32) = "RDETPEDPRO"
tables(1, 32) = "SELECT NUMERO,ALMORIG, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, DCTO, IVA, RE,METIDO, FMODI , ROWGUID FROM RDETPEDPRO"
tables(0, 33) = "REPRESEN"
tables(1, 33) = "SELECT CODIGO, NOMBRE, DIRECCION, CODPOS, LOCALIDAD, PROVINCIA, PAIS, TELEFONO1, TELEFONO2, FAX, EMAIL, NIF, COMISION, APLICARCOM, COMENTARIO, CODBAN, ENTIDAD, SUCURSAL, DC, CUENTA, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM REPRESEN"
tables(0, 34) = "CABPEDCLI"
tables(1, 34) = "SELECT NUMERO, FECHA, CODCLI, COMENTARIO, CODREP, COMISION, FPAGO, CODUSR, ESTADO, TRNSPORTI, DCTOPP, GASTOS, PLAZOE, PORTES, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM CABPEDCLI"
tables(0, 35) = "DETPEDCLI"
tables(1, 35) = "SELECT NUMERO, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, DCTO, IVA, RE, FMODI , ROWGUID FROM DETPEDCLI"
tables(0, 36) = "CABPRES"
tables(1, 36) = "SELECT NUMERO,ALMORIG, FECHA, CODPROV,CODALM,CODUSR, COMENTARIO, FPAGO, ESTADO, TRNSPORTI, DCTOPP, GASTOS, PLAZOE, PORTES, MBAJA, FBAJA,FALTA,TOTALIVA,TOTALNET,CODPTRN,SUCODIGO,ALBARAN,FACTURA, FMODI ,DESTINO, ROWGUID FROM CABPRES"
tables(0, 37) = "DETPRES"
tables(1, 37) = "SELECT NUMERO, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, DCTO, IVA, RE, FMODI , ROWGUID FROM DETPRES"
tables(0, 38) = "ECABPEDCLI"
tables(1, 38) = "SELECT NUMERO, FECHA, CODCLI, COMENTARIO, CODREP, COMISION, FPAGO, CODUSR, ESTADO, TRNSPORTI, DCTOPP, GASTOS, PLAZOE, PORTES, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM ECABPEDCLI"
tables(0, 39) = "SECCIONES"
tables(1, 39) = "SELECT CODIGO, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM SECCIONES"
tables(0, 40) = "SECTORES"
tables(1, 40) = "SELECT CODST, SECTOR, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM SECTORES"
tables(0, 41) = "SUBFAM"
tables(1, 41) = "SELECT CODIGO, CODFAM, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM SUBFAM"
tables(0, 42) = "TALLAS"
tables(1, 42) = "SELECT CODIGO, DESCRIPCION, CATTALL, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM TALLAS"
tables(0, 43) = "STOCK"
tables(1, 43) = "SELECT CODART, TALLA, COLOR, TEMPOR, CODALM, STOCK, FMODI , ROWGUID FROM STOCK"
tables(0, 44) = "DETTRANS"
tables(1, 44) = "SELECT CODIGO, CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES, FMODI , ROWGUID FROM DETTRANS"
tables(0, 45) = "EDETPEDCLI"
tables(1, 45) = "SELECT NUMERO, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, DCTO, IVA, RE, FMODI , ROWGUID FROM EDETPEDCLI"
tables(0, 46) = "TARIFAS"
tables(1, 46) = "SELECT CODIGO, ACTIVA, PORCEN, DESCRIPCION, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM TARIFAS"
tables(0, 47) = "TARJETAS"
tables(1, 47) = "SELECT CODIGO, DESCRIPCION, TASA, MBAJA, FBAJA, FALTA, FMODI , ROWGUID FROM TARJETAS"

Dim rs As New ADODB.Recordset
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
On Error GoTo Error
    nFile = FreeFile
    Open srcFile For Output As #nFile
    For i = 0 To (UBound(tables, 2) - LBound(tables, 2) - 1)
        rs.Open "SELECT * FROM sysobjects WHERE id = object_id(N'[dbo].[" & tables(0, i) & "]')", locCnn, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            rs.Close
            rs.Open tables(1, i), locCnn, adOpenStatic, adLockReadOnly
            'rs.Fields(0).
            If Not rs.EOF Then
                Print #nFile, "[" & tables(0, i) & "]"
                strFields = ""
                For j = 0 To rs.fields.Count - 1
                    strFields = strFields & CStr(Trim(rs.fields.Item(j).Name)) & vbTab
                Next j
                Print #nFile, strFields
                Print #nFile, rs.GetString(adClipString, , vbTab, vbCr + vbLf)
                Print #nFile, "[FIN " & tables(0, i) & "]"
                Count = Count + 1
            End If
        End If
        rs.Close
    Next i
    Close nFile
    Debug.Print "-------->" & Count
    
    CompressFile srcFile, zipFile
    ExportDB = zipFile
    Exit Function
Error:
    ExportDB = Empty
    Close nFile
    MsgBox Err.Description
End Function

Public Function ImportDB(filename As String) As Boolean
Dim i As Integer
Dim j As Integer
Dim Count As Long
Dim nFile As Long
Dim fields() As String
Dim Values() As String
Dim cFields As New Collection
Dim cValues As New Collection
Dim line As String
Dim cTableName As String
Dim currentSection As String
Dim passedSection As Boolean
Dim passedFields As Boolean
Dim Value As Variant
Dim rGuid As String
Dim origName As String

DecompressFile filename, origName

Dim rs As New ADODB.Recordset
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
On Error GoTo Error
    nFile = FreeFile
    Open origName For Input As #nFile
    passedFields = False
    passedSection = False
    locCnn.BeginTrans
    While Not EOF(nFile)
InitLoop:
        Line Input #nFile, line
        If Len(line) = 0 Then GoTo InitLoop
        If Left(line, 1) = "[" Then
            If Len(currentSection) > 0 Then
                currentSection = ""
                passedSection = False
                passedFields = False
                Debug.Print "-------->Cerre " & line
                rs.Close
            Else
                currentSection = line
                passedSection = True
                Debug.Print "-------->Abri " & line
                cTableName = line
                rs.Open "SELECT * FROM " & cTableName, locCnn, adOpenDynamic, adLockPessimistic
            End If
        Else
            If Not passedFields Then
                passedFields = True
                fields = Split(line, vbTab)
                Debug.Print "-------->Lei Campos " & line
            Else
                Values = Split(line, vbTab)
                ArrayToCollection fields, Values, cValues
                If Not rs.BOF Then
                    rs.MoveFirst
                End If
                rGuid = cValues("rowguid")
                If Len(rGuid) = 0 Then
                    rGuid = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"
                End If
                rs.Find "rowguid=" & rGuid
                If rs.EOF Then
                    rs.AddNew
                Else
                    If Not IsNull(rs("FMODI")) And Len(cValues("FMODI")) > 0 Then
                        If rs("FMODI") >= CDate(cValues("FMODI")) Then
                            GoTo InitLoop
                        End If
                    Else
                            GoTo InitLoop
                    End If
                End If
                For i = 1 To cValues.Count
                    Select Case rs.fields(fields(i - 1)).Type
                        Case adVarChar, adChar, adVarWChar, adWChar
                            Value = CStr(cValues(i))
                        Case adInteger, adSmallInt, adTinyInt, adSingle, adUnsignedInt, adUnsignedTinyInt, adUnsignedSmallInt, adBoolean
                            If cValues(i) = "" Then
                                Value = Null
                            Else
                                Value = CLng(cValues(i))
                            End If
                        Case adDouble, adNumeric
                            If cValues(i) = "" Then
                                Value = Null
                            Else
                                Value = CDbl(cValues(i))
                            End If
                        Case adDate, adDBDate, adDBTime, adDBTimeStamp
                            If cValues(i) = "" Then
                                Value = Null
                            Else
                                Value = CDate(cValues(i))
                            End If
                        Case adGUID
                            If cValues(i) = "" Then
                                Value = Null
                            Else
                                Value = CVar(cValues(i))
                            End If
                    End Select
                    rs(fields(i - 1)) = Value
                Next i
                rs.Update
            End If
        End If
    Wend
    locCnn.CommitTrans
    Close nFile
    Kill origName
    ImportDB = True
    Exit Function
Error:
    ImportDB = False
    Close nFile
    Kill origName
    locCnn.RollbackTrans
    MsgBox Err.Description
End Function

Private Sub ArrayToCollection(ByRef fields() As String, Values() As String, ByRef Col As Collection)
    Dim i As Integer
    
    For i = 1 To Col.Count
        Col.Remove 1
    Next i
    For i = 0 To (UBound(fields) - LBound(fields) - 1)
        Col.Add Values(i), fields(i)
    Next i
End Sub

