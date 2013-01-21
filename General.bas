Attribute VB_Name = "General"
'---------------------------------------------------------------------------------------
' Module    : General
' DateTime  : 12/10/2003 02:16
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

'La conexión al servidor local
Public locCnn As New ADODB.Connection
Public strLocCnn As String

'La conexión al servidor local para el proveedor MSDATASHAPE
Public locCnnSP As New ADODB.Connection
Public strLocCnnSP As String

'La conexión al servidor local, a la base de datos de configuración
Public EmpCnn As New ADODB.Connection
Public strEmpCnn As String

'La conexión al servidor CENTRAL (REMOTO)
Public SrvCnn As New ADODB.Connection
Public strSrvCnn As String

Public UsarTimerIP As Long

Private Const Version = "2.0.0"
Public Const titulo = "PC Gestion " & Version

Public Const Separacion_MDIForm = 10

Private crystalr As Object 'para almacenar el crystal report

Const dir_rpts = "\Reports\"

Const rpt_etiquetas = "codbar.rpt"
Const rpt_totTrn = "tottrn.rpt"
Const rpt_totPed = "totped.rpt"
Const rpt_totPedT = "totpedt.rpt"

'Almacena la temporada actual de trabajo.
Public TemporadaActual As Byte

'Almacena el almacen actual de trabajo (para transferencias). Este
'valor es obtenido al autentificar el usuario.
Public AlmacenActual As Byte

'Almacena el CODIGO de usuario actual
Public UsuarioActual As Integer

'Almacena la CAJA asignada al usuario
Public CajaActual As Byte

'Almacena el CENTRO asignado al usuario actual
Public CentroActual As Byte

'Almacena las lineas del ticket
Public LineasTicket(5) As String

'define el codigo X en las ventas para crear un arreglo rapidamente
Public Const CodigoArreglos = 5


'Aqui se define el tipo de servidor
' 1 - MSSQL
' 2 - MySQL
Public Const TipoServer = 1




'_____________________________________________________________
'Almacena el TIPO DE PERMISO para el usuario actual
Public TipoPermiso As Byte
' 0 -> dependiente comun (minimo permiso)
' 1 -> supervisor (permiso total)
'_____________________________________________________________


'Tamaño en caracteres del código de barras:
'5 codigo articulo
'3 temporada
'2 talla
'3 color
Public Const LenCodBar = 13

Public Type MiCodBar
    CODIGO_ART As String * 5
    TEMPORADA_ART As String * 3
    TALLA_ART As String * 2
    COLOR_ART As String * 3
End Type

'Caracter de moneda (€, o el caracter que sea)
Public SimboloMoneda As String * 1

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Enum MakeAsModal
    Modal = 0
    Modalless = 1
End Enum

'cadena de conexion para los mdb de access 2000.
Public Const strCnnMdb = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source="


'la conexión al servidor remoto
'Public srvCnn As New ADODB.Connection
'Public Const sRmtCnn = "PROVIDER=MSDASQL;dsn=SQLSER;uid=;pwd=;"

'Devuelve un campo pedido en la consulta SQL. (Devuelve el campo
'0 del recordset). Si no existe el registro (o error) devuelve un "@"
Public Function devuelve_campo(strSQL As String, Optional conexion As ADODB.Connection) As Variant
Dim rc As New ADODB.Recordset
On Error GoTo devuelve_campo_Error

If conexion Is Nothing Then
rc.Open strSQL, locCnn, adOpenDynamic, adLockReadOnly
Else
rc.Open strSQL, conexion, adOpenDynamic, adLockReadOnly
End If

If rc.EOF Then
    devuelve_campo = "@"
Else
    
    'si es nulo (no hay registros) devolver 1
    If IsNull(rc.fields(0).Value) Then
    
    devuelve_campo = "@"
    
    Else
    
    devuelve_campo = rc.fields(0).Value
    
    End If
    
End If

rc.Close
Set rc = Nothing

On Error GoTo 0
Exit Function

devuelve_campo_Error:
    devuelve_campo = "@"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure devuelve_campo of Módulo General", vbExclamation, titulo
End Function


'---------------------------------------------------------------------------------------
' Subrutina   : Descompone_CBAR
' Fecha/Hora  : 11/01/2004 17:38
' Autor       : JCASTILLO
' Propósito   : Descompone el codigo de barras en los diversos codigos y de
'               vuelve un tipo MiCodBar
'---------------------------------------------------------------------------------------
Public Function Descompone_CBAR(codigo As String) As MiCodBar

   On Error GoTo Descompone_CBAR_Error

    With Descompone_CBAR
        .CODIGO_ART = Mid(codigo, 1, 5)
        .TEMPORADA_ART = Mid(codigo, 6, 3)
        '12345 678 90 123
        '00000-000-00-000
        .TALLA_ART = Mid(codigo, 9, 2)
        .COLOR_ART = Mid(codigo, 11, 3)
    End With

   Exit Function

Descompone_CBAR_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Descompone_CBAR de Módulo General"
End Function


'---------------------------------------------------------------------------------------
' Procedure : carga_cadenas_conexion
' DateTime  : 12/10/2003 01:47
' Author    : Administrador
' Purpose   : Carga las cadenas de conexión para el proyecto. Lee de la unica base de datos
'             Fija: CONFIG.
'---------------------------------------------------------------------------------------
Public Function carga_cadenas_conexion(idEmpresa As Long) As Boolean
On Error GoTo carga_cadenas_conexion_Error

Const fichero_cfg = "\Config.pcg"

Dim rc As New ADODB.Recordset


'//////////////////////////////////////////
'/////////////  Connection String para enlazar a un SQLServer sin autentificación windows
'/////////////  cambiar el user, passw y DB por el correspondiente.
'//////////////////////////////////////////
'Provider=SQLOLEDB.1;Password=serv01;Persist Security Info=True;User ID=server;Initial Catalog=LOCAL;Data Source=81.33.28.52
'//////////////////////////////////////////

'La conexión a la base de datos,de configuración (local en mdb)
strEmpCnn = strCnnMdb & App.Path & fichero_cfg
'strEmpCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & App.Path & fichero_cfg

'si no existe la DB config ... generar y presentar los formularios de Configuración de Empresa
'y configuración de puesto, y salir.
If Dir(App.Path & fichero_cfg) = "" Then
  
    Call CreateDatabaseConfig
    
    EmpCnn.Open strEmpCnn
    
    frmMntEmp.Configuracion_Inicial = True
    frmMntEmp.Show
    
    DoEvents
    
    carga_cadenas_conexion = True
    Exit Function
    
Else

    EmpCnn.Open strEmpCnn
    
End If

rc.Open "select * from EMPRESAS where ID = " & idEmpresa, EmpCnn, adOpenStatic, adLockReadOnly

'si no hay registros
If rc.RecordCount <= 0 Then
    With frmMntEmp
        .Configuracion_Inicial = True
        .Show
        DoEvents
    End With
    
    carga_cadenas_conexion = True
    Exit Function
End If

'La conexión al servidor local
strLocCnn = rc.fields("CONSTRING")
'strLocCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Trim(rc.fields("BBDDCLI").Value) & ";Data Source=" & Trim(rc.fields("IPCLI").Value)

'Conexión al servidor local para el proveedor msdatashape
strLocCnnSP = "PROVIDER=MSDataShape;Data " & strLocCnn 'PROVIDER=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & rc.Fields("BBDDCLI").Value & ";Data Source=" & rc.Fields("IPCLI").Value

'La conexión al servidor CENTRAL (REMOTO)
'strSrvCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & rc.fields("BBDDSRV").Value & ";Data Source=" & rc.fields("IPSRV").Value

rc.Close
Set rc = Nothing

On Error GoTo 0
Exit Function

carga_cadenas_conexion_Error:

    carga_cadenas_conexion = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure carga_cadenas_conexion of Módulo General", vbExclamation, titulo

End Function


'---------------------------------------------------------------------------------------
' Procedure : habilita
' DateTime  : 12/10/2003 12:41
' Author    : Administrador
' Purpose   : Habilitar o deshabilitar todos los mitext
'             que estan enlazados a datos
'---------------------------------------------------------------------------------------
'
Public Function habilita(formulario As Form, SiNo As Boolean)
Dim tmpm As Object

For Each tmpm In formulario.Controls

    If TypeName(tmpm) = "miText" Then
    
        If tmpm.DataField <> "" Then
    
            tmpm.Enabled = SiNo
            
        End If
    
    End If
    
Next

Set tmpm = Nothing
End Function



'---------------------------------------------------------------------------------------
' Procedure : main
' DateTime  : 12/10/2003 02:16
' Author    : Administrador
' Purpose   : Iniciar el programa
'---------------------------------------------------------------------------------------
'
Sub Main()
 
'Si ya se esta ejecutando
If App.PrevInstance Then

        If MsgBox("Ya esta ejecutando el programa. ¿Dese abrirlo otra vez?", vbQuestion + vbYesNo, titulo) = vbNo Then
        'Activar la otra instancia
        On Error Resume Next
        AppActivate titulo
        End
        
        End If
        
End If
 
 
If carga_cadenas_conexion(1) = False Then
  If leer_configuracion Then
      frmInicSesion.Show
    End If
  'FrmTMPInicio.Show
End If

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : escribe_txt_config
' Fecha/Hora  : 13/12/2004 20:35
' Autor       : JCASTILLO
' Propósito   : Escribir TXT de configuración del puesto, para enlazar desde los
'               módulos creados con genexus. Forzar = true para eliminar un posible
'               fichero previo.
'
'               Tambien escribe el fichero con el id del usuario que abrio pcgestion
'
'---------------------------------------------------------------------------------------
Public Sub escribe_txt_config(forzar As Boolean)
Const ftxt = "pcg20.dat"
Const ftxtu = "PCG20U.DAT"

   On Error GoTo escribe_txt_config_Error
   
        
        If forzar Then
            If Dir(ftxt) <> "" Then Kill ftxt
        End If
                        
        'escribir configuración del puesto
        If Dir(ftxt) = "" Then
            Open ftxt For Output As #100
            Print #100, CentroActual
            Print #100, CajaActual
            Print #100, AlmacenActual
            Close #100
        End If
        
        'escribir id usuario q abrio el programa
        If Dir(ftxtu) <> "" Then Kill ftxtu
        
        Open ftxtu For Output As #100
        Print #100, TipoPermiso
        Close #100
        

   On Error GoTo 0
   Exit Sub

escribe_txt_config_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento escribe_txt_config de Módulo General"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : leer_configuracion
' DateTime  : 01/11/2003 21:52
' Author    : Administrador
' Purpose   : Leer la configuración y cargar las variables.
'             devuelve TRUE si no ha ocurrido ningun error.
'---------------------------------------------------------------------------------------
'

Public Function leer_configuracion() As Boolean
Dim EmpresaActual As Long
Dim m_matriz As Variant
Dim tmpconn As New ADODB.Connection

  On Error GoTo leer_configuracion_Error

  tmpconn.Open strEmpCnn
  m_matriz = devuelve_matriz("SELECT CODEMP, CODALM, CODCAJA, CODCEN FROM PUESTCNF", tmpconn)
  
  If Not IsArray(m_matriz) Then
    tmpconn.Close
    Set tmpconn = Nothing
    MsgBox "Revise la configuración del PUESTO (CONFIG)", vbCritical, titulo
    leer_configuracion = False
    Exit Function
  End If
  
  EmpresaActual = m_matriz(0)
  AlmacenActual = m_matriz(1)
  CajaActual = m_matriz(2)
  CentroActual = m_matriz(3)
  
  ReDim m_matriz(0)
  'obtener lineas del ticket
  m_matriz = devuelve_matriz("SELECT CL1, CL2, CL3, PL1, PL2 FROM EMPRESAS WHERE ID = " & EmpresaActual, tmpconn)
  
  LineasTicket(0) = m_matriz(0)
  LineasTicket(1) = m_matriz(1)
  LineasTicket(2) = m_matriz(2)
  LineasTicket(3) = m_matriz(3)
  LineasTicket(4) = m_matriz(4)
  
  tmpconn.Close
  Set tmpconn = Nothing
     
  ReDim m_matriz(0)
  
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
   'obtener la temporada de trabajo actual
   TemporadaActual = devuelve_campo("SELECT IDTEM FROM TEMPOR WHERE ACTUAL = 1")
        
   'obtener el ultimo caracter de la expresión, para obtener el simbolo
   'de la moneda actual del sistema (€ u otro). Se usa en miText
   SimboloMoneda = Right(Format(1, "Currency"), 1)
   
   leer_configuracion = True
   
   'crear el txt de configuracion para genexus (solo si no existe)
   Call escribe_txt_config(False)
   
   On Error GoTo 0
   Exit Function
leer_configuracion_Error:
   
    leer_configuracion = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure leer_configuracion of Módulo General"
    
    frmMntEmp.Show
End Function



Public Function ValidEmail(ByVal strCheck As String) As Boolean
'Created by Chad M. Kovac
'Tech Knowledgey, Inc.
'http://www.TechKnowledgeyInc.com

Dim bCK As Boolean
Dim strDomainType As String
Dim strDomainName As String
Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
Dim i As Integer

bCK = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
If Not bCK Then GoTo ExitFunction

bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
If Not bCK Then GoTo ExitFunction

' Check for invalid characters.
If Len(strCheck) > Len(sInvalidChars) Then
    For i = 1 To Len(sInvalidChars)
        If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
Else
    For i = 1 To Len(strCheck)
        If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
End If

If InStr(1, strCheck, "@") > 1 Then 'Check for an @ symbol
    bCK = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
Else
    bCK = False
End If
If Not bCK Then GoTo ExitFunction

strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
If Not bCK Then GoTo ExitFunction

strDomainType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
If Not bCK Then GoTo ExitFunction

strCheck = Left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
Do Until InStr(1, strCheck, ".") <= 1
    If Len(strCheck) >= InStr(1, strCheck, ".") Then
        strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
    Else
        bCK = False
        GoTo ExitFunction
    End If
Loop
If strCheck = "." Or Len(strCheck) = 0 Then bCK = False

ExitFunction:
ValidEmail = bCK

strDomainType = ""
strDomainName = ""
End Function



'*==========================================
'PARA HACER EL FORMULARIO MODAL SI/NO POR CODIGO
'*==========================================

'In A Module:
'


Public Sub MakeModal(ByRef frmParent As Form, ByVal isModal As MakeAsModal)
    On Error Resume Next
    Dim RetVal As Long
    RetVal = EnableWindow(frmParent.hwnd, isModal)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : procesa_informes
' DateTime  : 10/11/2003 21:27
' Author    : Administrador
' Purpose   : Procesar los informes de la aplicación. Recibe el parametro numinf
'---------------------------------------------------------------------------------------
'
Public Sub procesa_informes(numinf As Integer, orden_asc As Boolean, Optional formula As String)

   On Error GoTo procesa_informes_Error
   
    Set crystalr = Nothing
    Set crystalr = CreateObject("Crystal.CrystalReport")
 
DoEvents

With crystalr

If Trim(formula) <> "" Then .SelectionFormula = formula


Select Case numinf
'etiquetas pedidos

Case 1

    
            .ReportFileName = App.Path & dir_rpts & rpt_etiquetas  '.DataFiles(0) = "c:\TempEtiquetasDB.mdb"
            
            'si dice orden descendente
            If Not orden_asc Then
                .SortFields(0) = "-{ETIQUETAS.ID}"
            End If
            
            
            
    

Case 2

            .ReportFileName = App.Path & dir_rpts & rpt_totTrn


Case 3

'rpt_totPed

            .ReportFileName = App.Path & dir_rpts & rpt_totPed
            

Case 4

            .ReportFileName = App.Path & dir_rpts & rpt_totPedT
            
            

End Select

            
            .WindowState = 2
            .Action = 1
            
End With


   On Error GoTo 0
   Exit Sub

procesa_informes_Error:

    Set crystalr = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure procesa_informes of Módulo General"

End Sub

'In A Form:
'
'Private Sub Command1_Click()
'    MakeModal frmMain, Modal
'End Sub

'Private Sub Command2_Click()
'    MakeModal frmMain, Modalless
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : stock
' DateTime  : 11/11/2003 12:08
' Author    : Administrador
' Purpose   : Rutina para el manejo del stock (adiciones o eliminaciones de artículos de stock. Devuelve:
'                 0 -> no hay ningun error ni mensaje
'                 1 -> stock por debajo de 0
'                 2 -> stock por debajo del minimo
'                 3 -> stock por encima del máximo
'                 4 -> ha ocurrido algun error
'---------------------------------------------------------------------------------------
Public Function stock(CODIGO_ARTICULO As Integer, Temporada As Byte, talla As Integer, Color As Integer, almacen As Byte, unidades As Double, añadir As Boolean, Optional conexion As ADODB.Connection) As Byte
Dim rc As New ADODB.Recordset
Dim rcart As New ADODB.Recordset

Dim tmpsql As String
Dim tmpartsql As String
On Error GoTo stock_Error

'conformamos la selección para el Artículo
tmpsql = "SELECT CODART, TEMPOR, TALLA, COLOR, CODALM, STOCK FROM STOCK WHERE CODART =" & CODIGO_ARTICULO & " AND TEMPOR = " & Temporada & " AND TALLA = " & talla & " AND COLOR = " & Color & " AND CODALM = " & almacen
tmpartsql = "SELECT STOCKMAX, STOCKMIN FROM MAARTIC WHERE CODIGO= " & CODIGO_ARTICULO & " AND TEMPOR= " & Temporada

With rc

'Set rc = New ADODB.Recordset
'Set rcart = New ADODB.Recordset
    
If conexion Is Nothing Then
    .Open tmpsql, locCnn, adOpenStatic, adLockOptimistic
    rcart.Open tmpartsql, locCnn, adOpenDynamic, adLockReadOnly
Else
    .Open tmpsql, conexion, adOpenStatic, adLockOptimistic
    rcart.Open tmpartsql, conexion, adOpenDynamic, adLockReadOnly
End If

'si NO existe el registro, crear ...
If .RecordCount <= 0 Then

        .AddNew
        
        .fields("CODART") = CODIGO_ARTICULO
        .fields("TEMPOR") = Temporada
        .fields("TALLA") = talla
        .fields("COLOR") = Color
        .fields("CODALM") = almacen
        
        If añadir Then
        .fields("STOCK") = unidades  'sumar unidades
        Else
        .fields("STOCK") = -unidades 'restar unidades
        End If
        
        'si es menor de 0, salir con condicion 1
        If .fields("STOCK") < 0 Then stock = 1
        'si es menor del minimo, salir con condicion 2
        If .fields("STOCK") < rcart.fields("STOCKMIN") Then stock = 2
        'si es por encima del maximo, salir con condicion 2
        If .fields("STOCK") > rcart.fields("STOCKMAX") Then stock = 3
        
        

        .Update

Else  'si existe el registro, actualizar el campo stock
    
        
        
        If añadir Then
        .fields("STOCK") = .fields("STOCK") + unidades 'sumar unidades
        Else
        .fields("STOCK") = .fields("STOCK") - unidades 'restar unidades
        End If
        
        'si es menor de 0, salir con condicion 1
        If .fields("STOCK") < 0 Then stock = 1
        'si es menor del minimo, salir con condicion 2
        If .fields("STOCK") < rcart.fields("STOCKMIN") Then stock = 2
        'si es por encima del maximo, salir con condicion 2
        If .fields("STOCK") > rcart.fields("STOCKMAX") Then stock = 3
        
        .Update

End If

DoEvents

rcart.Close
Set rcart = Nothing

.Close
End With

tmpsql = ""
tmpartsql = ""
    
   Set rc = Nothing
    
   On Error GoTo 0
   Exit Function

stock_Error:
    
    stock = 4
    Set rc = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stock of Módulo General"

End Function
'---------------------------------------------------------------------------------------



'---------------------------------------------------------------------------------------
' Procedure : Habilita_miTextNum
' DateTime  : 11/11/2003 22:18
' Author    : Administrador
' Purpose   : Habilita o deshabilita los controles miTextNum. Colocar
'             esta rutina en SetButtons:
'             call Habilita_miTextNum (me, not bval)
'---------------------------------------------------------------------------------------
'
'Public Sub Habilita_miTextNum(formulario As Form, habilita As Boolean)

'Dim mitmp As Object

   'On error GoTo Habilita_miTextNum_Error

 '   With formulario
  '
   '     For Each mitmp In formulario
    '
     '   If TypeOf mitmp Is miTextNum Then
 '           mitmp.Enabled = habilita
  '      End If
   '
   '     Next
  '
 '   End With

'Set mitmp = Nothing
 '  On Error GoTo 0
 '  Exit Sub

'Habilita_miTextNum_Error:

 '   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Habilita_miTextNum of Módulo General"
'End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : cerrar_conexiones
' Fecha/Hora  : 19/11/2003 23:46
' Autor       : JCASTILLO
' Propósito   : Cerrar las conexiones al servidor (local y remoto) que hayan podido quedar
'               abiertas.
'---------------------------------------------------------------------------------------
Public Sub cerrar_conexiones()
Dim frmX As Form

   On Error GoTo cerrar_conexiones_Error

   On Error Resume Next
          
   With locCnn
    If .State <> 0 Then
        .Close
    End If
   Set locCnn = Nothing
   End With
                
   With locCnnSP
    If .State <> 0 Then
        .Close
    End If
   Set locCnnSP = Nothing
   End With
   
   With SrvCnn
      If .State <> 0 Then
        .Close
    End If
   Set SrvCnn = Nothing
   End With
   
    For Each frmX In Forms
        Unload frmX
        Set frmX = Nothing
    Next
   
   'On Error GoTo 0
   Exit Sub

cerrar_conexiones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cerrar_conexiones de Módulo General"
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : crear_nueva_transferencia
' Fecha/Hora     : 24/11/2003 11:12
' Autor             : JCastillo
' Propósito       : Añade un registro a PTRANS con los datos de la transferencia. Devuelve el ID de transferencia
'                       asignado.
'---------------------------------------------------------------------------------------
Public Function crear_nueva_transferencia(codart As Long, tempor As Byte, Color As Integer, talla As Integer, Almacen_Origen As Byte, Almacen_Destino As Byte, es_entrada As Boolean, es_cancelacion As Boolean, unidades As Double, conexion As ADODB.Connection) As Long
Dim rc As New ADODB.Recordset
Dim tmpcodigo As Variant

   On Error GoTo crear_nueva_transferencia_Error
    
    With rc
        .Open "SELECT * FROM PTRANS", conexion, adOpenDynamic, adLockOptimistic
        
         tmpcodigo = devuelve_campo("select MAX(IDTRANS) + 1 from PTRANS", conexion)
          
        .AddNew
        
        .fields("IDTRANS") = tmpcodigo
        .fields("CODART") = codart
        .fields("TEMPOR") = tempor
        .fields("CODCOL") = Color
        .fields("CODTALLA") = talla
        
        'si es una entrada, almacén de origen = 0
        If Not es_entrada Then
        
            .fields("CODALMORIG") = 0
            
        Else
        
            .fields("CODALMORIG") = Almacen_Origen
        
        End If
        
        .fields("CODALMDEST") = Almacen_Destino
        .fields("ENTRADA") = es_entrada
        .fields("UNIDADES") = unidades
        .fields("CANCEL") = es_cancelacion
        
        .Update
    
    End With
    
   'devolver el código
   crear_nueva_transferencia = tmpcodigo
   
   Set tmpcodigo = Nothing
   On Error GoTo 0
   Exit Function

crear_nueva_transferencia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento crear_nueva_transferencia de Módulo General"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : CodigoSeguridad_TRN
' Fecha/Hora    : 13/01/2004 11:22
' Autor         : JCastillo
'                                                    COD TRNS -ALM
' Propósito     : Enmascara un numero 000000000-000 en otro
'                     9 digitos codigo de transferencia
'                     3 digitos codigo de almacen
'                     Es un codigo de seguridad para poder aceptar la transferencia
'                     incluso si no coinciden las prendas en la comprobación.
'---------------------------------------------------------------------------------------
Public Function CodigoSeguridad_TRN(codigo As String) As Double
Dim tmpvalor As Double
Dim tmpcodalm As Double
Dim tmpcodigo As String
Dim tmpcodigo2 As String
Dim car As String * 1
Dim var As Long

   On Error GoTo CodigoSeguridad_TRN_Error

   If Not IsNumeric(codigo) Then Exit Function
   
   tmpcodalm = CDbl(Right(codigo, 3))
   tmpvalor = CDbl(codigo)

   'multiplicamos el codigo completo por la parte correspondiente al codigo
   'de almacén
   tmpvalor = CDbl((((tmpvalor * tmpcodalm) + 1) * 28.6) / 30.5)
   
   tmpcodigo = CStr(tmpvalor)
   
   'si tiene mas de 10 digitos (codigo muy grande), aplicar una rutina de
   'reducción
   Do Until Len(tmpcodigo) <= 10
           
        'coger un caracter si, y otro no
        For var = 1 To Len(tmpcodigo) Step 2
        
            car = Mid(tmpcodigo, var, 1)
            
            If IsNumeric(car) Then
                tmpcodigo2 = tmpcodigo2 & Mid(tmpcodigo, var, 1)
            End If
            
            car = " "
            
        Next var
        
        'asignar nuevo codigo
        tmpcodigo = tmpcodigo2
        tmpcodigo2 = ""
        
   Loop

   tmpvalor = CDbl(tmpcodigo)

   'devolver el valor
   CodigoSeguridad_TRN = tmpvalor
            
   On Error GoTo 0
   Exit Function

CodigoSeguridad_TRN_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CodigoSeguridad_TRN de Modulo General"
 
End Function



'---------------------------------------------------------------------------------------
' Subrutina   : devuelve_matriz
' Fecha/Hora  : 15/01/2004 20:54
' Autor       : JCASTILLO
' Propósito   : Devuelve una matriz de campos, para cuando es necesario hacer mas de una
'               consulta a la misma tabla, en vez de usar un devuelve_campo para cada
'               consulta
'---------------------------------------------------------------------------------------
Public Function devuelve_matriz(ConsultaSQL As String, conexion As ADODB.Connection) As Variant
Dim rc As New ADODB.Recordset
Dim var As Integer
Dim miMat As Variant

   On Error GoTo devuelve_matriz_Error

        rc.Open ConsultaSQL, conexion, adOpenStatic, adLockReadOnly
    
        If rc.RecordCount > 0 Then
        
            ReDim miMat(rc.fields.Count)
    
            For var = 0 To rc.fields.Count - 1
               
                miMat(var) = rc.fields(var).Value
        
            Next
        
        Else
        
            miMat = "@"
        
        End If
    
    rc.Close
    Set rc = Nothing
    
    devuelve_matriz = miMat

   On Error GoTo 0
   Exit Function

devuelve_matriz_Error:

    'devuelve_matriz(0) = "@"
    Set rc = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento devuelve_matriz de Módulo General"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : Obtiene_Precom_Pedido
' Fecha/Hora     : 22/01/2004 11:07
' Autor             : JCastillo
' Propósito       :  Obtener el ultimo precio de compra de los pedidos
'---------------------------------------------------------------------------------------
Public Function Obtiene_Precom_Pedido(p_Codart As Long, p_tempor As Byte, Optional p_talla As Integer, Optional p_color As Integer, Optional conexion As ADODB.Connection) As Single
Dim Tmp  As Variant
Dim tmppcom As Double

   On Error GoTo Obtiene_Precom_Pedido_Error
   'obtener el ultimo pedido en el q aparece el artículo ...
    
    If p_talla > 0 And p_color > 0 Then
        Tmp = devuelve_matriz("SELECT TOP 1 PRECOM, DCTO,  IVA, RE from DETPEDPRO where CODART = " & p_Codart & " AND TEMPOR = " & p_tempor & " AND CODTALLA = " & p_talla & " AND CODCOL  = " & p_color & " ORDER BY NUMERO DESC, LINEA DESC", conexion)
    Else
        Tmp = devuelve_matriz("SELECT TOP 1 PRECOM, DCTO,  IVA, RE from DETPEDPRO where CODART = " & p_Codart & " AND TEMPOR = " & p_tempor & " ORDER BY NUMERO DESC, LINEA DESC", conexion)
    End If
    
    If IsArray(Tmp) Then
        If UBound(Tmp) > 1 Then
        
        'obtiene el precio de compra
        tmppcom = Tmp(0) - ((Tmp(0) * Tmp(1)) / 100)
        Obtiene_Precom_Pedido = tmppcom + (tmppcom * Tmp(2) / 100)
        
        End If
    
    Else
    
        Obtiene_Precom_Pedido = 0
    
    End If
 
   On Error GoTo 0
   Exit Function

Obtiene_Precom_Pedido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Obtiene_Precom_Pedido de Formulario frmEtiqLibre"
End Function


Public Sub Espera(nSeg As Single)
   Dim nIni As Single
   Dim nFin As Single
   nIni = Timer
   nFin = nIni + nSeg
   Do While nFin > Timer
      DoEvents
   Loop
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : Obtiene_IVA_Pedido
' Fecha/Hora     : 22/01/2004 11:07
' Autor             : JCastillo
' Propósito       : Obtener el ultimo IVA de los pedidos
'---------------------------------------------------------------------------------------
Public Function Obtiene_IVA_Pedido(p_Codart As Long, p_tempor As Byte, Optional p_talla As Integer, Optional p_color As Integer, Optional conexion As ADODB.Connection) As Single
Dim Tmp  As Variant

   On Error GoTo Obtiene_IVA_Pedido
   'obtener el ultimo pedido en el q aparece el artículo ...
    
    If p_talla > 0 And p_color > 0 Then
        Tmp = devuelve_campo("SELECT TOP 1 IVA from DETPEDPRO where CODART = " & p_Codart & " AND TEMPOR = " & p_tempor & " AND CODTALLA = " & p_talla & " AND CODCOL  = " & p_color & " ORDER BY NUMERO DESC, LINEA DESC", conexion)
    Else
        Tmp = devuelve_campo("SELECT TOP 1 IVA from DETPEDPRO where CODART = " & p_Codart & " AND TEMPOR = " & p_tempor & " ORDER BY NUMERO DESC, LINEA DESC", conexion)
    End If
    
    'si existe ...
    If Tmp <> "@" Then
        Obtiene_IVA_Pedido = Tmp
    'si no existe, poner iva 0
    Else
        Obtiene_IVA_Pedido = 0
    End If
 
   On Error GoTo 0
   Exit Function

Obtiene_IVA_Pedido:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Obtiene_IVA_Pedido de Formulario frmEtiqLibre"
End Function

Private Sub CreateDatabaseConfig()
On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Dim rc As Recordset

sCnn = strEmpCnn

Cat.Create sCnn

  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "EMPRESAS"
    .Columns.Append "BBDDCLI", adVarWChar, 10
      .Columns("BBDDCLI").Properties("Default").Value = "LOCAL"
    .Columns.Append "BBDDSRV", adVarWChar, 10
      .Columns("BBDDSRV").Properties("Default").Value = "LOCAL"
    .Columns.Append "CIF", adVarWChar, 12
    .Columns.Append "CL1", adVarWChar, 30
      .Columns("CL1").Properties("Nullable").Value = True
    .Columns.Append "CL2", adVarWChar, 30
      .Columns("CL2").Properties("Nullable").Value = True
    .Columns.Append "CL3", adVarWChar, 30
      .Columns("CL3").Properties("Nullable").Value = True
    .Columns.Append "CODPOS", adVarWChar, 5
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMENT", adLongVarWChar
      .Columns("COMENT").Properties("Nullable").Value = True
    .Columns.Append "CONSTRING", adVarWChar, 250
    .Columns.Append "CREADACLI", adBoolean
      .Columns("CREADACLI").Properties("Nullable").Value = False
    .Columns.Append "CREADASRV", adBoolean
      .Columns("CREADASRV").Properties("Nullable").Value = False
    .Columns.Append "DIRECC", adVarWChar, 50
      .Columns("DIRECC").Properties("Nullable").Value = True
    .Columns.Append "FAX", adVarWChar, 12
      .Columns("FAX").Properties("Nullable").Value = True
    .Columns.Append "ID", adInteger
      .Columns("ID").Properties("AutoIncrement").Value = True
      .Columns("ID").Properties("Nullable").Value = True
    .Columns.Append "IPCLI", adVarWChar, 15
      .Columns("IPCLI").Properties("Default").Value = "(local)"
    .Columns.Append "IPSRV", adVarWChar, 15
      .Columns("IPSRV").Properties("Default").Value = "(local)"
    .Columns.Append "LOCALI", adVarWChar, 20
      .Columns("LOCALI").Properties("Nullable").Value = True
    .Columns.Append "LOGO", adLongVarBinary
      .Columns("LOGO").Properties("Nullable").Value = True
    .Columns.Append "LOGOF", adLongVarBinary
      .Columns("LOGOF").Properties("Nullable").Value = True
    .Columns.Append "NOMBRE", adVarWChar, 50
    .Columns.Append "PL1", adVarWChar, 30
      .Columns("PL1").Properties("Nullable").Value = True
    .Columns.Append "PL2", adVarWChar, 30
      .Columns("PL2").Properties("Nullable").Value = True
    .Columns.Append "PROVIN", adVarWChar, 20
      .Columns("PROVIN").Properties("Nullable").Value = True
    .Columns.Append "RAZO", adVarWChar, 50
    .Columns.Append "TELEF", adVarWChar, 12
      .Columns("TELEF").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(0)


  '----------* Table Definition of PUESTCNF *----------
  Set Tbl(6) = New ADOX.Table
  Tbl(6).ParentCatalog = Cat
  With Tbl(6)
    .Name = "PUESTCNF"
    .Columns.Append "CODALM", adUnsignedTinyInt
      .Columns("CODALM").Properties("Nullable").Value = False
    .Columns.Append "CODCAJA", adUnsignedTinyInt
      .Columns("CODCAJA").Properties("Nullable").Value = False
    .Columns.Append "CODCEN", adUnsignedTinyInt
      .Columns("CODCEN").Properties("Nullable").Value = False
    .Columns.Append "CODEMP", adUnsignedTinyInt
      .Columns("CODEMP").Properties("Nullable").Value = False
    .Columns.Append "ID", adSmallInt
      .Columns("ID").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(6)

  Set Cat = Nothing
  
  Set rc = New ADODB.Recordset
  
  'Escribir una configuración inicial para la empresa
  rc.Open "SELECT * FROM EMPRESAS", strEmpCnn, adOpenDynamic, adLockOptimistic
  
  With rc
    .AddNew
    .fields("ID").Value = 1
    .fields("RAZO") = "Empresa Inicial (cambiar)"
    .fields("CIF") = "0"
    .fields("DIRECC") = "C/ Inicial"
    .fields("CODPOS") = "00000"
    .fields("LOCALI") = "Localidad Inicial"
    .fields("PROVIN") = "Provincia Inicial"
    .fields("TELEF") = "0"
    .fields("FAX") = "0"
    .fields("NOMBRE") = "Titular Inicial (cambiar)"
    .fields("BBDDSRV") = "LOCAL"
    .fields("BBDDCLI") = "LOCAL"
    .fields("IPCLI") = "(local)"
    .fields("IPSRV") = "(local)"
    .fields("CREADASRV") = True
    .fields("CREADACLI") = True
    .fields("CL1") = "Empresa Inicial (cambiar)"
    .fields("CL2") = "------------------------------"
    .fields("CL3") = "    Telef. 900 00 00 00"
    .fields("PL1") = "- Gracias por su compra -"
    .fields("PL2") = "------------------------------"
    .fields("CONSTRING") = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & .fields("BBDDCLI").Value & ";Data Source=" & .fields("IPCLI").Value
    .Update
    .Close
  End With
  
  'Escribir una configuración inicial para el puesto
  rc.Open "SELECT * FROM PUESTCNF", strEmpCnn, adOpenDynamic, adLockOptimistic
  
  With rc
    .AddNew
    .fields("ID").Value = 1
    .fields("CODEMP").Value = 1
    .fields("CODALM").Value = 1
    .fields("CODCAJA").Value = 1
    .fields("CODCEN").Value = 1
    .Update
  End With
  
  rc.Close
  Set rc = Nothing
  
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Code Gen Error")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select

End Sub
