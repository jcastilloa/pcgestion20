Attribute VB_Name = "SINC_TRANS"
'---------------------------------------------------------------------------------------
' Modulo      : SINC_TRANS
' Fecha/Hora  : 19/11/2003 21:10
' Autor       : JCASTILLO
' Propósito   : Rutinas para la sincronización y transferencias.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' PROCESO DE LA TRANSFERENCIA:
'---------------------------------------------------------------------------------------
' Transferencias (Tabla PTRANS). Estados posibles para la transferencia:

' 0 -> EN CREACION: la transferencia esta en creación, todavia no le sale a los puestos, solo desde el
'        central o el puesto desde donde se crea. En este estado es posible borrar la transferencia.
'
' 1 -> PENDIENTE: la transferencia ya no esta en creación, significa que ya le sale a los puestos que
'        intervienen en la misma. desde el puesto de destino de la mercancía, se puede ACEPTAR. y desde
'        el puesto de origen se puede anular.
'
' 2 -> ACEPTADA. La transferencia ha sido aceptada y añadida en el Destino.
'
' 3 -> CANCELADA: La transferencia ha sido cancelada. Solo se puede cancelar desde el origen, pues hay
'        que devolver las unidades previamente descontadas al pasar a ESTADO=1 (pendiente).

'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const STILL_ACTIVE = &H103
Const PROCESS_QUERY_INFORMATION = &H400


Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal Options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const FORMAT_FULL = &H1

Public Type Cabecera_Grid_Print
    Lineas() As String
    cuantos As Long
End Type

Private Type Rect
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type TFormatRange
  hdc         As Long
  hdcTarget   As Long
  rc          As Rect
  rcPage      As Rect
End Type

Const WM_USER = &H400
Const VP_FORMATRANGE = WM_USER + 125
Const VP_YESIDO = 456654


'---------------------------------------------------------------------------------------
' Procedimiento : iniciar_transferencia
' Fecha/Hora     : 11/12/2003 10:32
' Autor             : JCastillo
' Propósito       : Pasar una transferencia de estado 0 a estado 1 (se pone estado 1 en
'                       el registro de PTRANS y se descuentan las unidades de stock.
'                       Devuelve 0 si se ha realizado correctamente y 1 si ha habido algun
'                       error y 2 si no hay registros en el detalle para esa transferencia
'---------------------------------------------------------------------------------------
Public Function iniciar_transferencia(CODIGO_TRANSF As Long, codigo_almacen As Byte, conexion As ADODB.Connection) As Byte
Dim rc As ADODB.Recordset
Dim entrans As Boolean

Dim miConn As New ADODB.Connection

   On Error GoTo iniciar_transferencia_Error
   
   Set rc = New ADODB.Recordset

   With miConn
        .ConnectionString = conexion.ConnectionString
        .CursorLocation = adUseServer
        .Open
        .BeginTrans
   End With

    
        'seleccionar transferencias
        rc.Open "SELECT CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES FROM DETTRANS WHERE CODIGO = " & CODIGO_TRANSF & " AND CODALM = " & codigo_almacen, conexion, adOpenStatic, adLockReadOnly
    
        'salida por no hay registros
        If rc.RecordCount <= 0 Then
            rc.Close
            Set rc = Nothing
            iniciar_transferencia = 2
            Exit Function
        End If
   

        entrans = True
   
   'quitar de stock
   Do Until rc.EOF
   
        'descontar las unidades para el almacen
        stock rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), codigo_almacen, rc.fields("UNIDADES"), False, miConn
        rc.MoveNext
        
   Loop
      
   rc.Close
   Set rc = Nothing
   
   'pasar a pendiente
   miConn.Execute "UPDATE PTRANS SET ESTADO = 1 WHERE CODIGO = " & CODIGO_TRANSF & " AND CODALMORIG = " & codigo_almacen
   
   
   With miConn
    .CommitTrans
    .Close
   End With
   
   entrans = False

   Set miConn = Nothing

   On Error GoTo 0
   Exit Function
      
iniciar_transferencia_Error:
    
    Set rc = Nothing
    
    If entrans Then
    
    With miConn
      If .State = 1 Then
        .RollbackTrans
        .Close
      End If
     End With
    
    End If
        
    Set miConn = Nothing
        
    iniciar_transferencia = 1

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento iniciar_transferencia de Módulo SINC_TRANS"
    
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : aceptar_transferencia
' Fecha/Hora     : 09/12/2003 12:54
' Autor             : JCastillo
' Propósito       : Acepta una transferencia PENDIENTE (ESTADO=1).
'                       Introduce las unidades en STOCK y marca la transferencia como ACEPTADA
'                       (ESTADO=2).
'                       Devuelve 0 si se ha realizado correctamente y 1 si ha habido algun error y
'                       2 si no hay registros en el detalle para esa transferencia
'---------------------------------------------------------------------------------------
Public Function aceptar_transferencia(codtrans As Long, codigo_almacen As Byte, conexion As ADODB.Connection) As Byte
Dim tmpstrconn As String
Dim tmprc As ADODB.Recordset

   On Error GoTo aceptar_transferencia_Error
    
    
    Set tmprc = New ADODB.Recordset
    'abrimos el rc antes de cambiar de cursor para q vaya mas rapido
    With tmprc
    
        .Open "SELECT CODART, TEMPOR, CODTALLA, CODCOL, CODALM, UNIDADES FROM DETTRANS WHERE CODIGO = " & codtrans & " AND CODALM = " & codigo_almacen, conexion, adOpenStatic, adLockReadOnly
        .ActiveConnection = Nothing  'lo desconectamosç
        'si no hay registros
        If .RecordCount = 0 Then
            tmprc.Close
            Set tmprc = Nothing
            'salida no hay registros
            aceptar_transferencia = 2
            Exit Function
        End If
    
        .MoveFirst
    End With

    With conexion
            tmpstrconn = .ConnectionString  'guardar el connection anterior por si acaso
            If .State <> 0 Then .Close
            .CursorLocation = adUseServer  'abrir con cursor para las transacciones
            .Open strLocCnn
            .BeginTrans
    End With
        
    'mientras haya registros
    Do Until tmprc.EOF
    
        'introducir unidades en stock
        stock tmprc.fields("CODART"), tmprc.fields("TEMPOR"), tmprc.fields("CODTALLA"), tmprc.fields("CODCOL"), AlmacenActual, tmprc.fields("UNIDADES"), True, conexion
        tmprc.MoveNext
        
    Loop
        
    With conexion
    
            'poner esa transferencia como ACEPTADA.
            .Execute "UPDATE PTRANS SET ESTADO=2  WHERE CODIGO = " & codtrans & " AND CODALMORIG = " & codigo_almacen
            
            DoEvents
    
            .CommitTrans  'aceptar todos los cambios
            If .State <> 0 Then .Close
            .CursorLocation = adUseClient
            .Open tmpstrconn
    End With
   
   tmpstrconn = ""
   
   'proceso terminado normalmente
   aceptar_transferencia = 0
    
   On Error GoTo 0
   Exit Function

aceptar_transferencia_Error:

  'On Error Resume Next
  
    With conexion
            If tmpstrconn <> .ConnectionString Then
                 .RollbackTrans
                If .State <> 0 Then .Close
                .CursorLocation = adUseClient
                .Open tmpstrconn
            End If
    End With
    
    tmpstrconn = ""
    'salida con error
    aceptar_transferencia = 1
    

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento aceptar_transferencia de Módulo SINC_TRANS"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : anular_transferencia_pendiente
' Fecha/Hora     : 11/12/2003 11:46
' Autor             : JCastillo
' Propósito        : Anula una transferencia pendiente. Solo usar con las que tengan ESTADO = 1 (pendiente).
'                      devuelve 0-> si se ha realizado correctamente
'                                   1-> si hubo algun error
'---------------------------------------------------------------------------------------
Public Function anular_transferencia_pendiente(codtrans As Long, codigo_almacen As Byte, conexion As ADODB.Connection) As Byte
Dim rc As ADODB.Recordset

   On Error GoTo anular_transferencia_pendiente_Error
   
   Set rc = New ADODB.Recordset
   
   'seleccionar transferencias
   rc.Open "SELECT CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES FROM DETTRANS WHERE CODIGO = " & codtrans & " AND CODALM = " & codigo_almacen, conexion, adOpenStatic, adLockReadOnly
    
   'salida por no hay registros
   If rc.RecordCount > 0 Then
  
        'quitar de stock
        Do Until rc.EOF
            'descontar las unidades para el almacen
            stock rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), codigo_almacen, rc.fields("UNIDADES"), True, conexion
            rc.MoveNext
        Loop
   
   End If
      
   rc.Close
   Set rc = Nothing
   
   'pasar a cancelada
   conexion.Execute "UPDATE PTRANS SET ESTADO = 3 WHERE CODIGO = " & codtrans & " AND CODALMORIG = " & codigo_almacen
  
   anular_transferencia_pendiente = 0
   On Error GoTo 0
   Exit Function

anular_transferencia_pendiente_Error:

   anular_transferencia_pendiente = 1
   Set rc = Nothing

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento anular_transferencia_pendiente de Módulo SINC_TRANS"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : crear_transferencia
' Fecha/Hora     : 09/12/2003 16:35
' Autor             : JCastillo
' Propósito       : Crea un nuevo registro de cabecera de transferencia (en PTRANS). Devuelve el numero de trans
'                       ferencia creado. Se crea una transferencia con estado 0 (Creación)
'---------------------------------------------------------------------------------------'
Public Function crear_transferencia(codigo_almacen_origen As Byte, codigo_almacen_destino As Byte, conexion As ADODB.Connection, DesdePedido As Boolean, Num_Ped As Double) As Long
Dim tmpcodigo As Variant

    On Error GoTo crear_transferencia_Error
    
    'si es desde almacen usar la codificacion 900000000 + el contador para almacen
    If DesdePedido Then
        tmpcodigo = devuelve_campo("select max(CODIGO) + 1 from PTRANS where (CODIGO > 900000000) AND (CODALMORIG = " & codigo_almacen_origen & ")", conexion)
        If tmpcodigo = "@" Then tmpcodigo = 900000000
    Else
        tmpcodigo = devuelve_campo("select max(CODIGO) + 1 from PTRANS where CODALMORIG = " & codigo_almacen_origen, conexion)
        If tmpcodigo = "@" Then tmpcodigo = 1
    End If
           
       'se inserta con el estado en creación
        conexion.Execute "INSERT INTO PTRANS (CODIGO, CODALMORIG, CODALMDEST, ESTADO, CODUSR, NUMPED) " & _
                              "VALUES (" & tmpcodigo & ", " & codigo_almacen_origen & ", " & codigo_almacen_destino & ", 0, " & UsuarioActual & ", " & Num_Ped & ")"
    
    crear_transferencia = CLng(tmpcodigo)
    Set tmpcodigo = Nothing
    
   On Error GoTo 0
   Exit Function

crear_transferencia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento crear_transferencia de Módulo SINC_TRANS"
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : crear_linea_transferencia
' Fecha/Hora     : 10/12/2003 12:19
' Autor             : JCastillo
' Propósito        : Crea una nueva linea de transferencia (en dettrans)
'---------------------------------------------------------------------------------------'
Public Function crear_linea_transferencia(CODIGO_TRANSF As Double, almacen As Byte, codart As Integer, tempor As Byte, talla As Integer, Color As Integer, unidades As Double, conexion As ADODB.Connection, DesdePedido As Boolean)
Dim tmpcodigo As Variant

    On Error GoTo crear_linea_transferencia_Error
    
    
        If DesdePedido Then
        tmpcodigo = devuelve_campo("select max(ID) + 1 from DETTRANS where (ID > 500000000) AND (CODALM = " & almacen & ")", conexion)
        If tmpcodigo = "@" Then tmpcodigo = 500000000
    Else
        tmpcodigo = devuelve_campo("select max(ID) + 1 from DETTRANS where CODALM = " & almacen, conexion)
        If tmpcodigo = "@" Then tmpcodigo = 1
    End If
    
    tmpcodigo = devuelve_campo("select max(ID) + 1 from DETTRANS where CODALM = " & almacen, conexion)
        
    If tmpcodigo = "@" Then tmpcodigo = 1
    
        conexion.Execute "INSERT INTO DETTRANS (CODIGO, CODART, ID, CODALM, TEMPOR, CODTALLA, CODCOL, UNIDADES) " & _
                              "VALUES (" & CODIGO_TRANSF & ", " & codart & ", " & tmpcodigo & ", " & almacen & ", " & tempor & ", " & talla & ", " & Color & "," & unidades & ")"
    
    'End If
    
    crear_linea_transferencia = CLng(tmpcodigo)
    Set tmpcodigo = Nothing
     

   On Error GoTo 0
   Exit Function

crear_linea_transferencia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento crear_linea_transferencia de Módulo SINC_TRANS"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : Abre_Conexion
' Fecha/Hora    : 02/12/2003 09:33
' Autor         : JCastillo
' Propósito     :Abrir el Acceso telefónico a Redes de Windows y ejecutar una conexión
'---------------------------------------------------------------------------------------
'Private Sub Abre_Conexion()
'Dim AbrirConexion As Long
'   On Error GoTo Abre_Conexion_Error
'
'AbrirConexion = Shell("rundll32.exe rnaui.dll,RnaDial " & "ConexiónInternet", 1)
'SendKeys "{ENTER}"
'
'   On Error GoTo 0
'   Exit Sub
'
'Abre_Conexion_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Abre_Conexion de Módulo SINC_TRANS"
'End Sub


Public Sub PrintGrid(Grid As VSFlexGrid, ByVal LeftMargin As Single, _
              ByVal TopMargin As Single, ByVal RightMargin As _
              Single, ByVal BottomMargin As Single, Titel As _
              String, Datum As String, Optional many As Integer)
              
  Dim tRange As TFormatRange
  Dim lReturn As Long
  Dim DName As String
  Dim DSchacht As Integer
  Dim gbeg As Long
  Dim CopyCW() As Long
  Dim GRef As Boolean
  Dim x%
  
    GRef = False
    If many > 0 Then
      ' Anzahl der zu druckenden Colums festlegen
      ' Alles > many wird auf colwidth = 0 gesetzt
      If Grid.Cols > many Then
        gbeg = Grid.Cols - many
        ReDim CopyCW(gbeg)
        Grid.Redraw = False
        For x = many To Grid.Cols - 1
          CopyCW(x - many) = Grid.ColWidth(x)
          Grid.ColWidth(x) = 0
        Next x
        GRef = True
      End If
    End If
    
    'mit wParam <> 0 kann überprüft werden
    'ob das Control OPP unterstützt, wenn ja wird
    '456654 (VP_YESIDO) zurückgeliefert
    lReturn = SendMessage(Grid.hwnd, VP_FORMATRANGE, 1, 0)
    
    If lReturn = VP_YESIDO Then
      
      'Struktur mit Formatierungsinformationen füllen
      Printer.ScaleMode = vbPixels
      
      With tRange
        .hdc = Printer.hdc
        
        'Höhe und Breite einer Seite (in Pixel)
        .rcPage.Right = Printer.ScaleWidth
        .rcPage.Bottom = Printer.ScaleHeight
        
        'Lage und Abmessungen des Bereichs auf den
        'gedruckt werden soll (in Pixel)
        .rc.Left = Printer.ScaleX(LeftMargin, vbMillimeters)
        .rc.Top = Printer.ScaleY(TopMargin, vbMillimeters)
        .rc.Right = .rcPage.Right - Printer.ScaleX(RightMargin, _
                                                   vbMillimeters)
                                                   
        .rc.Bottom = .rcPage.Bottom - Printer.ScaleY(BottomMargin, _
                                                     vbMillimeters)
      End With
  
      'Drucker initialisieren
      Printer.Print vbNullString
      
      'Seite(n) drucken
      Do
        Printer.CurrentX = Printer.ScaleX(LeftMargin, vbMillimeters)
        Printer.CurrentY = Printer.ScaleY(10, vbMillimeters)
        If Titel <> "" Then Printer.Print Titel
  
        Printer.CurrentX = Printer.ScaleX(LeftMargin, vbMillimeters)
        Printer.CurrentY = Printer.ScaleY(16, vbMillimeters)
        
        If Datum <> "" Then
          Printer.Print Datum
        Else
          Printer.Print Format(Date, "DD.MM.YYYY")
        End If
        lReturn = SendMessage(Grid.hwnd, VP_FORMATRANGE, 0, _
                              VarPtr(tRange))
        
        If lReturn < 0 Then
          Exit Do
        Else
          Printer.NewPage
        End If
      Loop
      Printer.EndDoc
  
      'Reset
      lReturn = SendMessage(Grid.hwnd, VP_FORMATRANGE, 0, 0)
    End If
    
    If GRef Then
      'Alle Colums wieder in richtiger Breite darstellen
      For x = many To Grid.Cols - 1
        Grid.ColWidth(x) = CopyCW(x - many)
      Next x
      Grid.Redraw = True
    End If
End Sub



Public Sub PrintFlexGrid(flxdata As _
Object, xmin As Single, ymin As Single, orientacion As Byte, CabeceraL1 As String, CabeceraL2 As String, Cabecera_Fontsize As Byte, Linea_Totales As Long, Optional fuente_tamano As Byte)
Const GAP = 60

' Orientacion
' vbPRORPortrait     1
' vbPRORLandscape 2

Dim xmax As Single
Dim ymax As Single
Dim x As Single
Dim C As Integer
Dim R As Integer
Dim ptr As Object
Dim guardar_cy As Long
Dim guardar_cx As Long
Dim tmppag As Long

   On Error GoTo PrintFlexGrid_Error

    Set ptr = Printer
    
    ptr.Orientation = orientacion
    
    
    With ptr.Font
        .Name = flxdata.Font.Name
    
        'Imprimir la Cabecera
        If CabeceraL1 <> "" Then
        
            If Cabecera_Fontsize > 0 Then
                .Size = Cabecera_Fontsize
            Else
                .Size = flxdata.Font.Size
            End If
              
       ptr.CurrentX = xmax = xmin + GAP
       ptr.Print CabeceraL1
       ptr.Print CabeceraL2
       ptr.Print
       
       ymin = ymin + ptr.CurrentY + 20
       
       End If
            
       If fuente_tamano > 0 Then
        .Size = fuente_tamano
       Else
        .Size = flxdata.Font.Size
       End If
        
    End With
       
       guardar_cy = ptr.CurrentY
       guardar_cx = ptr.CurrentX
       
       ptr.CurrentY = ptr.Height - (ptr.TextHeight(ptr.Page) * 4)
       ptr.CurrentX = ptr.Width - 1200 '(ptr.TextWidth(ptr.Page) * 4)
       
       ptr.Print ptr.Page
       
       ptr.CurrentY = guardar_cy
       ptr.CurrentX = guardar_cx
       
       tmppag = 1
           



    With flxdata
        ' See how wide the whole thing is.
        xmax = xmin + GAP
        For C = 0 To .Cols - 1
          
          If Not .ColHidden(C) Then
            
            
            'Select Case .ColFormat(c)
                    
            'Case "Currency"
            
            xmax = xmax + .ColWidth(C) + 2 * GAP
                    
            'Case Else
            
            'xmax = xmax + .ColWidth(c) + 2 * GAP
                        
            'End Select
                      
          End If
          
        Next C

        ' Print each row.
        ptr.CurrentY = ymin
        For R = 0 To .Rows - 1
            'Draw a line above this row.
            
            'si es igual a la linea de totales
            If R = Linea_Totales - 1 Then
                ptr.FontUnderline = True
                ptr.FontBold = True
            Else
                ptr.FontUnderline = False
                ptr.FontBold = False
            End If
            
            If R > 0 Then ptr.Line (xmin, _
                ptr.CurrentY)-(xmax, ptr.CurrentY)
            ptr.CurrentY = ptr.CurrentY + GAP

            ' Print the entries on this row.
            x = xmin + GAP
            
            For C = 0 To .Cols - 1
                
                If Not .ColHidden(C) Then
                    
                    ptr.CurrentX = x
                    
                   'Select Case .ColFormat(c)
                    
                    'Case "Currency"
                    
                   ' ptr.Print Format(BoundedText(ptr, .TextMatrix(r, _
                        c), .ColWidth(c)), "Currency")
                   '
                  ' Case Else
                   
                    ptr.Print BoundedText(ptr, .TextMatrix(R, C), .ColWidth(C));
                    
                  ' End Select
                        
                    x = x + .ColWidth(C) + 2 * GAP
                End If
                
            Next C
           
            ptr.CurrentY = ptr.CurrentY + GAP

            ' Move to the next line.
            ptr.Print
                
        'cuando cambie de página, imprimir el numero otra vez
        If tmppag <> ptr.Page Then
        
        
               guardar_cy = ptr.CurrentY
               guardar_cx = ptr.CurrentX
       
               ptr.CurrentY = ptr.Height - (ptr.TextHeight(ptr.Page) * 4)
               ptr.CurrentX = ptr.Width - 1200 '(ptr.TextWidth(ptr.Page) * 4)
       
               ptr.Print ptr.Page
       
               ptr.CurrentY = guardar_cy
               ptr.CurrentX = guardar_cx
               
               tmppag = ptr.Page
               
               'guardar_cy = ptr.CurrentY
               'ptr.CurrentY = ptr.Height - (ptr.TextHeight(ptr.Page) * 4)
               'ptr.Print ptr.Page
               'ptr.CurrentY = guardar_cy
               
        End If
        
        
        Next R
        ymax = ptr.CurrentY

        ' Draw a box around everything.
        'ptr.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
        x = xmin
      '  For c = 0 To .Cols - 2
      '      If Not .ColHidden(c) Then
      '          X = X + .ColWidth(c) + 2 * GAP
      '          ptr.Line (X, ymin)-(X, ymax)
      '      End If
      '  Next c
    End With
    
    ptr.EndDoc
    Set ptr = Nothing

   On Error GoTo 0
   Exit Sub

PrintFlexGrid_Error:

    ptr.EndDoc
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento PrintFlexGrid de Módulo SINC_TRANS"
    
End Sub

' Truncate the string so it fits within the width.
Private Function BoundedText(ByVal ptr As Object, ByVal txt _
    As String, ByVal max_wid As Single) As String
    Do While ptr.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
End Function

Private Sub copia_rc(origen As ADODB.Recordset, destino As ADODB.Recordset, comprueba_duplicado As Boolean)
Dim var As Long
Dim añadir As Boolean

'meter registros ...
   On Error GoTo copia_rc_Error

añadir = True

Do Until origen.EOF
    
    'primero buscar si existe el registro en el destino ...
    If comprueba_duplicado Then
        destino.Find "ROWGUID = '" & origen.fields("ROWGUID") & "'", , adSearchForward, 1
            'si no se encuentra el registro, añadirlo nuevo
            If destino.EOF Then
                añadir = True
            Else
                añadir = False
            End If
    End If
    
    If añadir Then
    
        destino.AddNew
    
        For var = 0 To origen.fields.Count - 1
        
                If Not IsNull(origen.fields(var)) Then _
                destino.fields(origen.fields(var).Name) = origen.fields(var)
    
        Next var
    
        destino.Update
    
    End If
    
    origen.MoveNext
    
Loop

   On Error GoTo 0
   Exit Sub

copia_rc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento copia_rc de Módulo SINC_TRANS"

End Sub

Private Sub Actualiza_rc(origen As ADODB.Recordset, destino As ADODB.Recordset, esPtrans As Boolean)
Dim var As Long
Dim añadir As Boolean

'meter registros ...
   On Error GoTo Actualiza_rc_Error

añadir = True

Do Until origen.EOF

            destino.Find "ROWGUID = " & origen.fields("ROWGUID"), , adSearchForward, 1
            
            'destino.Filter = "ROWGUID = " & origen.fields("ROWGUID")  ', , adSearchForward, 1"
            
            
            'si no se encuentra el registro, añadirlo nuevo
            If destino.EOF Then
                añadir = True
            Else
                añadir = False
            End If
            
            'destino.Filter = ""
    
    
    If añadir Then
    
        destino.AddNew
    
        For var = 0 To origen.fields.Count - 1
        
        Debug.Print origen.Source
        
                If ver_tipo_campo(origen.fields(var)) = 2 Then
                'si es de texto, y nulo, meter un " "
                    
                    If IsNull(origen.fields(var)) Then
                    destino.fields(origen.fields(var).Name) = " "
                    Else
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
                    End If
 
                Else
                
                    If Not IsNull(origen.fields(var)) Then _
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
                
                End If
             
        Next var
    
    'si ya existe, actualizar el registro
    Else
    
        For var = 0 To origen.fields.Count - 1
        
                'si es Ptrans tener cuidado con no actualizar el campo
                'estado
                If esPtrans And origen.fields(var).Name = "ESTADO" Then
                
                'no hacer nada ...
                
                'de lo contrario
                Else
                    
                    If Not IsNull(origen.fields(var)) Then _
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
    
                
                End If
                
        Next var
    
    End If
    
    destino.Update
    destino.Requery
    
    origen.MoveNext
    
Loop

   On Error GoTo 0
   Exit Sub

Actualiza_rc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Actualiza_rc de Módulo SINC_TRANS"

End Sub




'---------------------------------------------------------------------------------------
' Procedimiento : Actualiza_rc_Clave
' Fecha/Hora     : 26/03/2004 16:40
' Autor             : JCastillo
' Propósito       : Actualiza recordset buscando por clave(s) ...
'---------------------------------------------------------------------------------------
Private Sub Actualiza_rc_Clave(origen As ADODB.Recordset, conexion As ADODB.Connection, NomTablaDest As String, Clave1 As String, Clave2 As String, Clave3 As String, Clave4 As String, WhereEspecial As String, OrdenBY As String, esPtrans As Boolean)
Dim var As Long
Dim añadir As Boolean
Dim destino As ADODB.Recordset
Dim tmpwhere As String

'meter registros ...
   On Error GoTo Actualiza_rc_Clave_Error

Set destino = New ADODB.Recordset

Do Until origen.EOF

'abrir el registro seleccionado en destino ...

If Trim(Clave1) <> "" Then
    tmpwhere = Clave1 & " = " & origen.fields(Clave1).Value
End If

If Trim(Clave2) <> "" Then
    tmpwhere = tmpwhere & " AND " & Clave2 & " = " & origen.fields(Clave2).Value
End If

If Trim(Clave3) <> "" Then
    tmpwhere = tmpwhere & " AND " & Clave3 & " = " & origen.fields(Clave3).Value
End If

If Trim(Clave4) <> "" Then
    tmpwhere = tmpwhere & " AND " & Clave4 & " = " & origen.fields(Clave4).Value
End If


'If OrdenBY <> "" Then OrdenBY = " ORDER BY " & OrdenBY
'If WhereEspecial <> "" Then WhereEspecial = " AND (" & WhereEspecial & ")"
 
If destino.State = 1 Then destino.Close

Debug.Print "select * from " & NomTablaDest & " WHERE " & tmpwhere
destino.Open "select * from " & NomTablaDest & " WHERE " & tmpwhere, conexion, adOpenStatic, adLockOptimistic

'If Trim(WhereEspecial) = "" Then
    'Debug.Print "select * from " & NomTablaDest & " WHERE " & tmpwhere & " " & OrdenBY,
'    destino.Open "select * from " & NomTablaDest & " WHERE " & tmpwhere & " " & OrdenBY, conexion, adOpenStatic, adLockOptimistic
'Else
'    destino.Open "select * from " & NomTablaDest & " WHERE " & tmpwhere & " AND (" & WhereEspecial & ") " & OrdenBY, conexion, adOpenStatic, adLockOptimistic
'End If


'si no existe el registro en destino ...
If destino.RecordCount <= 0 Then
 añadir = True
'si existe, actualizar ...
Else
 añadir = False
End If

    If añadir Then
    
        destino.AddNew
    
        For var = 0 To origen.fields.Count - 1
        
        Debug.Print origen.Source
        
                If ver_tipo_campo(origen.fields(var)) = 2 Then
                'si es de texto, y nulo, meter un " "
                    
                    If IsNull(origen.fields(var)) Then
                    destino.fields(origen.fields(var).Name) = " "
                    Else
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
                    End If
 
                Else
                
                    If Not IsNull(origen.fields(var)) Then _
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
                
                End If
             
        Next var
    
    'si ya existe, actualizar el registro
    Else
    
        For var = 0 To origen.fields.Count - 1
        
                'si es Ptrans tener cuidado con no actualizar el campo
                'estado
                If esPtrans And origen.fields(var).Name = "ESTADO" Then
                
                'no hacer nada ...
                
                'de lo contrario
                Else
                    
                    If Not IsNull(origen.fields(var)) Then _
                    destino.fields(origen.fields(var).Name) = origen.fields(var)
    
                
                End If
                
        Next var
    
    End If
    
    destino.Update
        
    origen.MoveNext
    
Loop

If destino.State = 1 Then destino.Close
Set destino = Nothing

   On Error GoTo 0
   Exit Sub

Actualiza_rc_Clave_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Actualiza_rc_Clave de Módulo SINC_TRANS"
 
End Sub


'


'---------------------------------------------------------------------------------------
' Procedimiento : Crea_TRN_Datos
' Fecha/Hora     : 19/02/2004 12:55
' Autor             : JCastillo
' Propósito       : Crear una mdb con los datos relativos a la transferencia dada ...
'---------------------------------------------------------------------------------------
Public Sub Crea_TRN_Datos(num_trans As Long, codalm As Byte, codalmdest As Byte, crear_mdb As Boolean, Comprimir As Boolean, Nombre_Ruta As String, conexion As ADODB.Connection, IncluirArt As Boolean)
Dim rc As New ADODB.Recordset
Dim rcm As New ADODB.Recordset
Dim var As Long
Dim tmpin As String  'para almacenar los codigos para formar IN (de select .. IN)
Dim sCnn As String
Dim total_trans As Double
Dim crear_registro As Boolean

'cadena de conexión
sCnn = strCnnMdb & Nombre_Ruta
'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & Nombre_Ruta

'se crea fisicamente el fichero
If crear_mdb Then Call CreateDatabaseTRN(Nombre_Ruta)

Debug.Print "SELECT * FROM PTRANS WHERE ((CODIGO = " & num_trans & ") AND (CODALMORIG = " & codalm & ") AND (CODALMDEST = " & codalmdest & "))"

'--------------- PTRANS ----------------------------------------------------------------------------
'leer del origen PTRANS
rc.Open "SELECT * FROM PTRANS WHERE ((CODIGO = " & num_trans & ") AND (CODALMORIG = " & codalm & ") AND (CODALMDEST = " & codalmdest & "))", conexion, adOpenDynamic, adLockReadOnly
    
    
    
'total = total + (total_trn - dcto)
total_trans = (rc.fields("TOTAL") - ((rc.fields("TOTAL") * rc.fields("DCTO")) / 100))

'introducir registros de PTRANS:
rcm.Open "SELECT * FROM PTRANS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, False)

'--------------- DETTRANS ----------------------------------------------------------------------------
'leer del origen DETTRANS
rc.Close
rc.Open "SELECT * FROM DETTRANS WHERE CODIGO = " & num_trans & " AND CODALM = " & codalm & " AND CAST(CODART AS VARCHAR(8)) + CAST(TEMPOR AS VARCHAR(3)) IN (SELECT CAST(CODART AS VARCHAR(8)) + CAST(TEMPOR AS VARCHAR(3)) FROM MAARTIC WHERE MBAJA = 0)", conexion, adOpenDynamic, adLockReadOnly

'introducir registros de PTRANS:
rcm.Close
rcm.Open "SELECT * FROM DETTRANS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, False)

'--------------- PTRANSMSG ----------------------------------------------------------------------------
'leer del origen DETTRANS
rc.Close
rc.Open "SELECT * FROM PTRANSMSG WHERE CODIGO = " & num_trans & " AND CODALMORIG = " & codalm, conexion, adOpenDynamic, adLockReadOnly
'introducir registros de PTRANS:
rcm.Close
rcm.Open "SELECT * FROM PTRANSMSG", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, False)

'--------------- MAARTIC ----------------------------------------------------------------------------

'leer del origen MAARTIC (los que existan en DETTRANS para esta transferencia)
rc.Close

If IncluirArt = False Then
rc.Open "SELECT * FROM MAARTIC WHERE (CAST(CODIGO AS varchar(5)) + CAST(TEMPOR AS varchar(3))) in (SELECT (CAST(CODART AS varchar(5)) + CAST(TEMPOR AS varchar(3))) FROM DETTRANS WHERE CODIGO = " & num_trans & " AND CODALM = " & codalm & ")", conexion, adOpenDynamic, adLockReadOnly
Else
rc.Open "SELECT * FROM MAARTIC WHERE TEMPOR IN (" & TemporadaActual & "," & TemporadaActual + 1 & ")", conexion, adOpenDynamic, adLockReadOnly
End If
'introducir registros de MAARTIC:
rcm.Close
rcm.Open "SELECT * FROM MAARTIC", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- MAPROV  (todo) ----------------------------------------------------------------------------
'leer del origen MAPROV (los que existan en DETTRANS para esta transferencia)

rc.Close
rc.Open "SELECT * FROM MAPROV WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
rcm.Close
rcm.Open "SELECT * FROM MAPROV", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'---------------  CLIENTES ----------------------------------------------------

'leer del origen CLIENTES (los que existan en DETTRANS para esta transferencia)
rc.Close
rc.Open "SELECT * FROM CLIENTES WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
rcm.Close
rcm.Open "SELECT * FROM CLIENTES", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'---------------  TALLAS ----------------------------------------------------

'leer del origen TALLAS
rc.Close
rc.Open "SELECT * FROM TALLAS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de TALLAS:
rcm.Close
rcm.Open "SELECT * FROM TALLAS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- COLORES ----------------------------------------------------

'leer del origen COLORES
rc.Close
rc.Open "SELECT * FROM COLORES WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de COLORES:
rcm.Close
rcm.Open "SELECT * FROM COLORES", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)


'--------------- CATEGORIAS TALLAS ----------------------------------------------------

'leer del origen CATEGORIAS TALLAS
rc.Close
rc.Open "SELECT * FROM CATTALL WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de CATEGORIAS TALLAS:
rcm.Close
rcm.Open "SELECT * FROM CATTALL", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- CENTROS ----------------------------------------------------

'leer del origen CENTROS
rc.Close
rc.Open "SELECT * FROM CENTROS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de CENTROS:
rcm.Close
rcm.Open "SELECT * FROM CENTROS", sCnn, adOpenDynamic, adLockOptimistic
Call copia_rc(rc, rcm, True)

'--------------- ALMACENES ----------------------------------------------------
'leer del origen ALMACENES
rc.Close
rc.Open "SELECT * FROM ALMACENES WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de ALMACENES:
rcm.Close
rcm.Open "SELECT * FROM ALMACENES", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- CAJAS ----------------------------------------------------
'leer del origen CAJAS
rc.Close
rc.Open "SELECT * FROM CAJAS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de CAJAS:
rcm.Close
rcm.Open "SELECT * FROM CAJAS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)


'--------------- SECCION ----------------------------------------------------
'leer del origen SECCION
rc.Close
rc.Open "SELECT * FROM SECCIONES WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de SECCION:
rcm.Close
rcm.Open "SELECT * FROM SECCIONES", sCnn, adOpenDynamic, adLockOptimistic
Call copia_rc(rc, rcm, True)

'--------------- FAMILIA ----------------------------------------------------
'leer del origen FAMILIA
rc.Close
rc.Open "SELECT * FROM FAMILIAS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcm.Close
rcm.Open "SELECT * FROM FAMILIAS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- SUBFAMILIA ----------------------------------------------------
'leer del origen SUBFAM
rc.Close
rc.Open "SELECT * FROM SUBFAM WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcm.Close
rcm.Open "SELECT * FROM SUBFAM", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)


'--------------- COSTURE ----------------------------------------------------
'leer del origen COSTURE
rc.Close
rc.Open "SELECT * FROM COSTURE WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcm.Close
rcm.Open "SELECT * FROM COSTURE", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- PERSONAL ----------------------------------------------------
'leer del origen PERSONAL
rc.Close
rc.Open "SELECT * FROM PERSONAL WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcm.Close
rcm.Open "SELECT * FROM PERSONAL", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- BANCOS ----------------------------------------------------
'leer del origen BANCOS
rc.Close
rc.Open "SELECT * FROM BANCOS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcm.Close
rcm.Open "SELECT * FROM BANCOS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- MAING ----------------------------------------------------
'leer del origen MAING
rc.Close
rc.Open "SELECT * FROM MAING WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
rcm.Close
rcm.Open "SELECT * FROM MAING", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- MAPAG ----------------------------------------------------
'leer del origen MAPAG
rc.Close
rc.Open "SELECT * FROM MAPAG WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de MAPAG
rcm.Close
rcm.Open "SELECT * FROM MAPAG", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- OFERTAS ----------------------------------------------------
'leer del origen OFERTAS
rc.Close
rc.Open "SELECT * FROM OFERTAS WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de OFERTAS
rcm.Close
rcm.Open "SELECT * FROM OFERTAS", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'--------------- TEMPOR ----------------------------------------------------
'leer del origen TEMPOR
rc.Close
rc.Open "SELECT * FROM TEMPOR WHERE MBAJA = 0", conexion, adOpenDynamic, adLockReadOnly
'introducir registros de TEMPOR
rcm.Close
rcm.Open "SELECT * FROM TEMPOR", sCnn, adOpenDynamic, adLockOptimistic

Call copia_rc(rc, rcm, True)

'Datos de la transferencia

rcm.Close

With rcm
    .Open "SELECT count(CODUSR) FROM CONF_TRN", sCnn, adOpenDynamic, adLockOptimistic
    
    'si es nulo, es 0
    If IsNull(.fields(0)) Then
    
    crear_registro = True
    
    Else
    
        If .fields(0) = 0 Then
            crear_registro = True
        ElseIf .fields(0) > 0 Then
            crear_registro = False
        End If
                        
    End If
    
    .Close
    .Open "SELECT * FROM CONF_TRN", sCnn, adOpenDynamic, adLockOptimistic
    
    If crear_registro Then
        .AddNew
        .fields("TOTAL") = 0
        .fields("NUMTRANS") = 0
    End If
    
    .fields("FHORA") = Now
    .fields("CODUSR") = UsuarioActual
    .fields("CODALMORIG") = codalm
    .fields("CODALMDEST") = codalmdest
    .fields("TOTAL") = .fields("TOTAL") + total_trans
    .fields("NUMTRANS") = .fields("NUMTRANS") + 1
    .Update
End With

rc.Close
rcm.Close
Set rc = Nothing
Set rcm = Nothing

'comprimir el fichero en el mismo pero con extension .trz
If Comprimir Then
    Call CompressFile(Nombre_Ruta, Left(Nombre_Ruta, Len(Nombre_Ruta) - 3) & "trnz")
    DoEvents
    'borrar el fichero original ...
    Kill (Nombre_Ruta)
End If



End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Leer_TRN_Datos
' Fecha/Hora    : 20/02/2004 18:10
' Autor         : JCastillo
' Propósito     :    Leer transferencia desde fichero
'                        Devuelve  0 -> todo OK
'                                       1 -> cancelada
'---------------------------------------------------------------------------------------
Public Function Leer_TRN_Datos(Nombre_Ruta As String, conexion As ADODB.Connection, mostrar_pregunta As Boolean) As Byte
Dim rct As New ADODB.Recordset
Dim rcl As New ADODB.Recordset
Dim var As Long
Dim tmpin As String  'para almacenar los codigos para formar IN (de select .. IN)
Dim sCnn As String

'carga cadena de conexión
   On Error GoTo Leer_TRN_Datos_Error

sCnn = strCnnMdb & Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"
'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"

'descomprimir el fichero
Call DecompressFile(Nombre_Ruta, Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb")
DoEvents

'--------------- CENTROS ----------------------------------------------------

rct.Open "SELECT * FROM CONF_TRN", sCnn, adOpenDynamic, adLockReadOnly

If mostrar_pregunta Then

    If rct.fields("CODALMDEST") <> AlmacenActual Then
        MsgBox "La transferencia no corresponde al almacen actual. Corresponde a: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMDEST"), locCnn)), vbInformation, titulo
        Leer_TRN_Datos = 1
        Exit Function
    End If

    'preguntar al usuario si las desea incorporar en el sistema
    If MsgBox("¿Desea incorporar al sistema la/s transferencia (" & rct.fields("NUMTRANS") & ") ?" & Chr(13) & _
      "Origen: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMORIG"), locCnn)) & Chr(13) & _
      "Destino: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMDEST"), locCnn)) & Chr(13) & _
      "Total(" & rct.fields("NUMTRANS") & "): " & Format(rct.fields("TOTAL"), "Currency") & Chr(13) & _
      "Fecha: " & rct.fields("FHORA") & Chr(13) & _
      "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rct.fields("CODUSR"), locCnn)), vbQuestion + vbYesNo, titulo) = vbNo Then
            
      rct.Close
      
      Set rct = Nothing
      Set rcl = Nothing
      
      'borrar el mdb
      Kill Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"
      Leer_TRN_Datos = 1
      Exit Function
      
    End If

End If

'borrar el fichero original (comprimido)...
Kill (Nombre_Ruta)

rct.Close
 
'leer del origen CENTROS
'rct.Close
rct.Open "SELECT * FROM CENTROS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CENTROS:
'rcm.Close
rcl.Open "SELECT * FROM CENTROS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
Call Actualiza_rc(rct, rcl, False)

'--------------- ALMACENES ----------------------------------------------------
'leer del origen ALMACENES
rct.Close
rct.Open "SELECT * FROM ALMACENES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de ALMACENES:
rcl.Close
rcl.Open "SELECT * FROM ALMACENES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- CAJAS ----------------------------------------------------
'leer del origen CAJAS
rct.Close
rct.Open "SELECT * FROM CAJAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CAJAS:
rcl.Close
rcl.Open "SELECT * FROM CAJAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)


'--------------- SECCION ----------------------------------------------------
'leer del origen SECCION
'rct.Close
rct.Close
rct.Open "SELECT * FROM SECCIONES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de SECCION:
'rcl.Close
rcl.Close
rcl.Open "SELECT * FROM SECCIONES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
Call Actualiza_rc(rct, rcl, False)

'--------------- FAMILIA ----------------------------------------------------
'leer del origen FAMILIA
rct.Close
rct.Open "SELECT * FROM FAMILIAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcl.Close
rcl.Open "SELECT * FROM FAMILIAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- SUBFAMILIA ----------------------------------------------------
'leer del origen SUBFAM
rct.Close
rct.Open "SELECT * FROM SUBFAM ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcl.Close
rcl.Open "SELECT * FROM SUBFAM ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- BANCOS ----------------------------------------------------
'leer del origen BANCOS
rct.Close
rct.Open "SELECT * FROM BANCOS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcl.Close
rcl.Open "SELECT * FROM BANCOS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- PERSONAL ----------------------------------------------------
'leer del origen PERSONAL
rct.Close
rct.Open "SELECT * FROM PERSONAL ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcl.Close
rcl.Open "SELECT * FROM PERSONAL ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)


'--------------- MAPROV  (todo) ----------------------------------------------------------------------------
'leer del origen MAPROV (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM MAPROV ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
rcl.Close
rcl.Open "SELECT * FROM MAPROV ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- MAARTIC ----------------------------------------------------------------------------

'leer del origen MAARTIC (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM MAARTIC ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAARTIC (solo en articulos de la temporada actual y siguiente:
rcl.Close
rcl.Open "SELECT * FROM MAARTIC WHERE TEMPOR IN (" & TemporadaActual - 1 & ", " & TemporadaActual & ", " & TemporadaActual + 1 & ") ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- CATEGORIAS TALLAS ----------------------------------------------------

'leer del origen CATEGORIAS TALLAS
rct.Close
rct.Open "SELECT * FROM CATTALL ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CATEGORIAS TALLAS:
rcl.Close
rcl.Open "SELECT * FROM CATTALL ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'---------------  TALLAS ----------------------------------------------------

'leer del origen TALLAS
rct.Close
rct.Open "SELECT * FROM TALLAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de TALLAS:
rcl.Close
rcl.Open "SELECT * FROM TALLAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- COLORES ----------------------------------------------------

'leer del origen COLORES
rct.Close
rct.Open "SELECT * FROM COLORES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de COLORES:
rcl.Close
rcl.Open "SELECT * FROM COLORES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- PTRANS ----------------------------------------------------------------------------
rct.Close
rcl.Close

'leer del origen PTRANS
rct.Open "SELECT * FROM PTRANS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
    
'introducir registros de PTRANS:
rcl.Open "SELECT * FROM PTRANS WHERE CODALMDEST = " & AlmacenActual & " ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'esPtrans = true
Call Actualiza_rc(rct, rcl, True)

'--------------- DETTRANS ----------------------------------------------------------------------------
'leer del origen DETTRANS
rct.Close
rct.Open "SELECT * FROM DETTRANS ORDER BY ID DESC", sCnn, adOpenDynamic, adLockReadOnly

'introducir registros de DETTRANS:
rcl.Close
rcl.Open "SELECT * FROM DETTRANS ORDER BY ID DESC", conexion, adOpenDynamic, adLockOptimistic
Call Actualiza_rc(rct, rcl, False)

'--------------- PTRANSMSG ----------------------------------------------------------------------------
'leer del origen DETTRANS
rct.Close
rct.Open "SELECT * FROM PTRANSMSG ORDER BY ID DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de PTRANS:
rcl.Close
rcl.Open "SELECT * FROM PTRANSMSG WHERE CODALMORIG = " & AlmacenActual & " ORDER BY ID DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'---------------  CLIENTES ----------------------------------------------------

'leer del origen CLIENTES (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM CLIENTES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
rcl.Close
rcl.Open "SELECT * FROM CLIENTES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- COSTURE ----------------------------------------------------
'leer del origen COSTURE
rct.Close
rct.Open "SELECT * FROM COSTURE ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
rcl.Close
rcl.Open "SELECT * FROM COSTURE ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- MAING ----------------------------------------------------
'leer del origen MAING
rct.Close
rct.Open "SELECT * FROM MAING ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
rcl.Close
rcl.Open "SELECT * FROM MAING ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- MAPAG ----------------------------------------------------
'leer del origen MAPAG
rct.Close
rct.Open "SELECT * FROM MAPAG ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPAG
rcl.Close
rcl.Open "SELECT * FROM MAPAG ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- OFERTAS ----------------------------------------------------
'leer del origen OFERTAS
rct.Close
rct.Open "SELECT * FROM OFERTAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de OFERTAS
rcl.Close
rcl.Open "SELECT * FROM OFERTAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)

'--------------- TEMPOR ----------------------------------------------------
'leer del origen TEMPOR
rct.Close
rct.Open "SELECT * FROM TEMPOR ORDER BY IDTEM DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de TEMPOR
rcl.Close
rcl.Open "SELECT * FROM TEMPOR ORDER BY IDTEM DESC", conexion, adOpenDynamic, adLockOptimistic

Call Actualiza_rc(rct, rcl, False)
'Datos de la transferencia

rct.Close
rcl.Close
Set rct = Nothing
Set rcl = Nothing


'Kill Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"

   On Error GoTo 0
   Exit Function

Leer_TRN_Datos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Leer_TRN_Datos de Módulo SINC_TRANS"

End Function



'---------------------------------------------------------------------------------------
' Procedimiento : Leer_TRN_Datos_Clave
' Fecha/Hora     : 26/03/2004 16:55
' Autor             : JCastillo
' Propósito       :  Leer transferencia desde fichero, lee los datos por clave
'                        Devuelve  0 -> todo OK
'                                       1 -> cancelada
'
Public Function Leer_TRN_Datos_Clave(Nombre_Ruta As String, conexion As ADODB.Connection, mostrar_pregunta As Boolean) As Byte
Dim rct As New ADODB.Recordset
'Dim rcl As New ADODB.Recordset
Dim var As Long
Dim tmpin As String  'para almacenar los codigos para formar IN (de select .. IN)
Dim sCnn As String

'carga cadena de conexión
 
   On Error GoTo Leer_TRN_Datos_Clave_Error

sCnn = strCnnMdb & Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"
'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"

'descomprimir el fichero
Call DecompressFile(Nombre_Ruta, Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb")
DoEvents

'--------------- CENTROS ----------------------------------------------------

rct.Open "SELECT * FROM CONF_TRN", sCnn, adOpenDynamic, adLockReadOnly

If mostrar_pregunta Then

    If rct.fields("CODALMDEST") <> AlmacenActual Then
        MsgBox "La transferencia no corresponde al almacen actual. Corresponde a: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMDEST"), locCnn)), vbInformation, titulo
        Leer_TRN_Datos_Clave = 1
        Exit Function
    End If

    'preguntar al usuario si las desea incorporar en el sistema
    If MsgBox("¿Desea incorporar al sistema la/s transferencia (" & rct.fields("NUMTRANS") & ") ?" & Chr(13) & _
      "Origen: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMORIG"), locCnn)) & Chr(13) & _
      "Destino: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rct.fields("CODALMDEST"), locCnn)) & Chr(13) & _
      "Total(" & rct.fields("NUMTRANS") & "): " & Format(rct.fields("TOTAL"), "Currency") & Chr(13) & _
      "Fecha: " & rct.fields("FHORA") & Chr(13) & _
      "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rct.fields("CODUSR"), locCnn)), vbQuestion + vbYesNo, titulo) = vbNo Then
            
      rct.Close
      
      Set rct = Nothing
      'Set rcl = Nothing
      
      'borrar el mdb
      Kill Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"
      Leer_TRN_Datos_Clave = 1
      Exit Function
      
    End If

End If

'borrar el fichero original (comprimido)...
Kill (Nombre_Ruta)

rct.Close
 
'leer del origen CENTROS
'rct.Close
rct.Open "SELECT * FROM CENTROS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CENTROS:
'rcm.Close
'rcl.Open "SELECT * FROM CENTROS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "CENTROS", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- ALMACENES ----------------------------------------------------
'leer del origen ALMACENES
rct.Close
rct.Open "SELECT * FROM ALMACENES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de ALMACENES:
'rcl.Close
'rcl.Open "SELECT * FROM ALMACENES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "ALMACENES", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- CAJAS ----------------------------------------------------
'leer del origen CAJAS
rct.Close
rct.Open "SELECT * FROM CAJAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CAJAS:

'rcl.Close
'rcl.Open "SELECT * FROM CAJAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "CAJAS", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- SECCION ----------------------------------------------------
'leer del origen SECCION
'rct.Close
rct.Close
rct.Open "SELECT * FROM SECCIONES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de SECCION:
'rcl.Close
'rcl.Close
'rcl.Open "SELECT * FROM SECCIONES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "SECCIONES", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- FAMILIA ----------------------------------------------------
'leer del origen FAMILIA
rct.Close
rct.Open "SELECT * FROM FAMILIAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
'rcl.Close
'rcl.Open "SELECT * FROM FAMILIAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "FAMILIAS", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- SUBFAMILIA ----------------------------------------------------
'leer del origen SUBFAM
rct.Close
rct.Open "SELECT * FROM SUBFAM ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
'rcl.Close
'rcl.Open "SELECT * FROM SUBFAM ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "SUBFAM", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- BANCOS ----------------------------------------------------
'leer del origen BANCOS
rct.Close
rct.Open "SELECT * FROM BANCOS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
'rcl.Close
'rcl.Open "SELECT * FROM BANCOS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "BANCOS", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- PERSONAL ----------------------------------------------------
'leer del origen PERSONAL
rct.Close
rct.Open "SELECT * FROM PERSONAL ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
'rcl.Close
'rcl.Open "SELECT * FROM PERSONAL ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "PERSONAL", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- MAPROV  (todo) ----------------------------------------------------------------------------
'leer del origen MAPROV (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM MAPROV ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
'rcl.Close
'rcl.Open "SELECT * FROM MAPROV ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "MAPROV", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- MAARTIC ----------------------------------------------------------------------------

'leer del origen MAARTIC (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM MAARTIC ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAARTIC (solo en articulos de la temporada actual y siguiente:
'rcl.Close
'rcl.Open "SELECT * FROM MAARTIC WHERE TEMPOR IN (" & TemporadaActual - 1 & ", " & TemporadaActual & ", " & TemporadaActual + 1 & ") ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "MAARTIC", "CODIGO", "TEMPOR", "", "", "TEMPOR IN (" & TemporadaActual - 1 & ", " & TemporadaActual & ", " & TemporadaActual + 1 & ")", "CODIGO DESC", False)


'--------------- CATEGORIAS TALLAS ----------------------------------------------------

'leer del origen CATEGORIAS TALLAS
rct.Close
rct.Open "SELECT * FROM CATTALL ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de CATEGORIAS TALLAS:
'rcl.Close
'rcl.Open "SELECT * FROM CATTALL ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "CATTALL", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'---------------  TALLAS ----------------------------------------------------

'leer del origen TALLAS
rct.Close
rct.Open "SELECT * FROM TALLAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de TALLAS:
'rcl.Close
'rcl.Open "SELECT * FROM TALLAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "TALLAS", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- COLORES ----------------------------------------------------

'leer del origen COLORES
rct.Close
rct.Open "SELECT * FROM COLORES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de COLORES:
'rcl.Close
'rcl.Open "SELECT * FROM COLORES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic
'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "COLORES", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- PTRANS ----------------------------------------------------------------------------
rct.Close
'rcl.Close

'leer del origen PTRANS
rct.Open "SELECT * FROM PTRANS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
    
'introducir registros de PTRANS:
'rcl.Open "SELECT * FROM PTRANS WHERE CODALMDEST = " & AlmacenActual & " ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, True)

'esPtrans = true
Call Actualiza_rc_Clave(rct, conexion, "PTRANS", "CODIGO", "CODALMORIG", "", "", "CODALMDEST = " & AlmacenActual, "CODIGO DESC", True)

'--------------- DETTRANS ----------------------------------------------------------------------------
'leer del origen DETTRANS
rct.Close
rct.Open "SELECT * FROM DETTRANS ORDER BY ID DESC", sCnn, adOpenDynamic, adLockReadOnly

'introducir registros de DETTRANS:
'rcl.Close
'rcl.Open "SELECT * FROM DETTRANS ORDER BY ID DESC", conexion, adOpenDynamic, adLockOptimistic
'Call Actualiza_rc(rct, rcl, False)

Call Actualiza_rc_Clave(rct, conexion, "DETTRANS", "ID", "CODALM", "", "", "", "ID DESC", False)

'--------------- PTRANSMSG ----------------------------------------------------------------------------
'leer del origen DETTRANS
rct.Close
rct.Open "SELECT * FROM PTRANSMSG ORDER BY ID DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de PTRANS:
'rcl.Close
'rcl.Open "SELECT * FROM PTRANSMSG WHERE CODALMORIG = " & AlmacenActual & " ORDER BY ID DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "PTRANSMSG", "ID", "CODALM", "", "", "CODALMORIG = " & AlmacenActual, "ID DESC", False)

'---------------  CLIENTES ----------------------------------------------------

'leer del origen CLIENTES (los que existan en DETTRANS para esta transferencia)
rct.Close
rct.Open "SELECT * FROM CLIENTES ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPROV:
'rcl.Close
'rcl.Open "SELECT * FROM CLIENTES ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "CLIENTES", "CODIGO", "CODCAJA", "", "", "", "CODIGO DESC", False)

'--------------- COSTURE ----------------------------------------------------
'leer del origen COSTURE
rct.Close
rct.Open "SELECT * FROM COSTURE ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de FAMILIA:
'rcl.Close
'rcl.Open "SELECT * FROM COSTURE ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "COSTURE", "CODIGO", "", "", "", "", "CODIGO DESC", False)


'--------------- MAING ----------------------------------------------------
'leer del origen MAING
rct.Close
rct.Open "SELECT * FROM MAING ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'rcl.Close
'rcl.Open "SELECT * FROM MAING ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "MAING", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- MAPAG ----------------------------------------------------
'leer del origen MAPAG
rct.Close
rct.Open "SELECT * FROM MAPAG ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de MAPAG
'rcl.Close
'rcl.Open "SELECT * FROM MAPAG ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "MAPAG", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- OFERTAS ----------------------------------------------------
'leer del origen OFERTAS
rct.Close
rct.Open "SELECT * FROM OFERTAS ORDER BY CODIGO DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de OFERTAS
'rcl.Close
'rcl.Open "SELECT * FROM OFERTAS ORDER BY CODIGO DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "OFERTAS", "CODIGO", "", "", "", "", "CODIGO DESC", False)

'--------------- TEMPOR ----------------------------------------------------
'leer del origen TEMPOR
rct.Close
rct.Open "SELECT * FROM TEMPOR ORDER BY IDTEM DESC", sCnn, adOpenDynamic, adLockReadOnly
'introducir registros de TEMPOR
'rcl.Close
'rcl.Open "SELECT * FROM TEMPOR ORDER BY IDTEM DESC", conexion, adOpenDynamic, adLockOptimistic

'Call Actualiza_rc(rct, rcl, False)
Call Actualiza_rc_Clave(rct, conexion, "TEMPOR", "IDTEM", "", "", "", "", "IDTEM DESC", False)
'Datos de la transferencia

rct.Close
'rcl.Close
Set rct = Nothing
'Set rcl = Nothing


'Kill Left(Nombre_Ruta, Len(Nombre_Ruta) - 4) & "mdb"

   


   On Error GoTo 0
   Exit Function

Leer_TRN_Datos_Clave_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Leer_TRN_Datos_Clave de Módulo SINC_TRANS"

End Function


'rct.Close

'With rct
  '  .Open "SELECT * FROM CONF_TRN", sCnn, adOpenDynamic, adLockOptimistic
  '
 '   If .RecordCount <= 0 Then
 '       .AddNew
 '       .fields("TOTAL") = 0
 '       .fields("NUMTRANS") = 0
 '   End If
    
   ' .fields("FHORA") = Now
   ' .fields("CODUSR") = UsuarioActual
   ' .fields("CODALMORIG") = codalm
   ' .fields("CODALMDEST") = codalmdest
   '' .fields("TOTAL") = .fields("TOTAL") + total_trans
   ' .fields("NUMTRANS") = .fields("NUMTRANS") + 1
   ' .Update
'End With

'Crea base de datos de TRANSFERENCIAS
'Obtiene el parametro NOMBRE y PATH del fichero
Public Sub CreateDatabaseTRN(Nombre_Ruta As String)
'On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(27) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String

'cadena de conexión
sCnn = strCnnMdb & Nombre_Ruta
'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & Nombre_Ruta

'si existe borrar ...
If Dir(Nombre_Ruta) <> "" Then Kill Nombre_Ruta

Cat.Create sCnn

  '----------* Table Definition of ALMACENES *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "ALMACENES"
    .Columns.Append "CODCEN", adUnsignedTinyInt
      .Columns("CODCEN").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adUnsignedTinyInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 25
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
      
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
      
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
      
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TELEFONO", adVarWChar, 12
      .Columns("TELEFONO").Properties("Nullable").Value = True
    .Columns.Append "UBICACION", adVarWChar, 25
      .Columns("UBICACION").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(0)

  '----------* Table Definition of BANCOS *----------
  Set Tbl(1) = New ADOX.Table
  Tbl(1).ParentCatalog = Cat
  With Tbl(1)
    .Name = "BANCOS"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adVarWChar, 5
      .Columns("CODPOS").Properties("Nullable").Value = False
    .Columns.Append "DOMICILIO", adVarWChar, 50
      .Columns("DOMICILIO").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAX", adVarWChar, 12
      .Columns("FAX").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "LOCALIDAD", adVarWChar, 40
      .Columns("LOCALIDAD").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "NOMBRE", adVarWChar, 50
      .Columns("NOMBRE").Properties("Nullable").Value = False
    .Columns.Append "PERCON", adVarWChar, 30
      .Columns("PERCON").Properties("Nullable").Value = False
    .Columns.Append "PROVINCIA", adVarWChar, 25
      .Columns("PROVINCIA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TELEFONO", adVarWChar, 12
      .Columns("TELEFONO").Properties("Nullable").Value = False
    .Columns.Append "TELEFPCON", adVarWChar, 12
      .Columns("TELEFPCON").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(1)

  '----------* Table Definition of CAJAS *----------
  Set Tbl(2) = New ADOX.Table
  Tbl(2).ParentCatalog = Cat
  With Tbl(2)
    .Name = "CAJAS"
    .Columns.Append "CAJA_A", adUnsignedTinyInt
      .Columns("CAJA_A").Properties("Nullable").Value = False
    .Columns.Append "CAJA_B", adUnsignedTinyInt
      .Columns("CAJA_B").Properties("Nullable").Value = False
    .Columns.Append "CODALM", adUnsignedTinyInt
      .Columns("CODALM").Properties("Nullable").Value = False
    .Columns.Append "CODCEN", adUnsignedTinyInt
      .Columns("CODCEN").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adUnsignedTinyInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 25
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SALDOINI", adSingle
      .Columns("SALDOINI").Properties("Nullable").Value = False
    .Columns.Append "TELEFONO", adVarWChar, 12
      .Columns("TELEFONO").Properties("Nullable").Value = True
    .Columns.Append "UBICACION", adVarWChar, 25
      .Columns("UBICACION").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(2)

  '----------* Table Definition of CATTALL *----------
  Set Tbl(3) = New ADOX.Table
  Tbl(3).ParentCatalog = Cat
  With Tbl(3)
    .Name = "CATTALL"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(3)

  '----------* Table Definition of CENTROS *----------
  Set Tbl(4) = New ADOX.Table
  Tbl(4).ParentCatalog = Cat
  With Tbl(4)
    .Name = "CENTROS"
    .Columns.Append "CODIGO", adUnsignedTinyInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adSmallInt
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMENTARIO", adVarWChar, 50
      .Columns("COMENTARIO").Properties("Nullable").Value = True
      
       .Columns.Append "INTERVALO", adSmallInt
      .Columns("INTERVALO").Properties("Nullable").Value = False
          
          .Columns.Append "HOSTDIR", adVarWChar, 50
      .Columns("HOSTDIR").Properties("Nullable").Value = True
      
      .Columns.Append "ACTIP", adBoolean
      .Columns("ACTIP").Properties("Nullable").Value = False
      
    .Columns.Append "DESCRIPCION", adVarWChar, 30
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "DIRECC", adVarWChar, 20
      .Columns("DIRECC").Properties("Nullable").Value = True
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAX", adVarWChar, 9
      .Columns("FAX").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "LOCALI", adVarWChar, 10
      .Columns("LOCALI").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "PROVIN", adVarWChar, 15
      .Columns("PROVIN").Properties("Nullable").Value = True
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TELFNO", adVarWChar, 9
      .Columns("TELFNO").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(4)

  '----------* Table Definition of CLIENTES *----------
  Set Tbl(5) = New ADOX.Table
  Tbl(5).ParentCatalog = Cat
  With Tbl(5)
    .Name = "CLIENTES"
    .Columns.Append "CODBAN", adSmallInt
      .Columns("CODBAN").Properties("Nullable").Value = True
    .Columns.Append "CODCAJA", adUnsignedTinyInt
      .Columns("CODCAJA").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adInteger
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adVarWChar, 5
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMENTARIO", adLongVarWChar
      .Columns("COMENTARIO").Properties("Nullable").Value = True
    .Columns.Append "CPENVIO", adVarWChar, 5
      .Columns("CPENVIO").Properties("Nullable").Value = True
    .Columns.Append "CUENTA", adVarWChar, 10
      .Columns("CUENTA").Properties("Nullable").Value = True
    .Columns.Append "DC", adVarWChar, 2
      .Columns("DC").Properties("Nullable").Value = True
    .Columns.Append "DCTO", adSingle
      .Columns("DCTO").Properties("Nullable").Value = False
    .Columns.Append "DCTOPP", adSingle
      .Columns("DCTOPP").Properties("Nullable").Value = False
    .Columns.Append "DIAPAGO1", adUnsignedTinyInt
      .Columns("DIAPAGO1").Properties("Nullable").Value = False
    .Columns.Append "DIAPAGO2", adUnsignedTinyInt
      .Columns("DIAPAGO2").Properties("Nullable").Value = False
    .Columns.Append "DIRECCION", adVarWChar, 40
      .Columns("DIRECCION").Properties("Nullable").Value = True
    .Columns.Append "DIRECENVIO", adVarWChar, 40
      .Columns("DIRECENVIO").Properties("Nullable").Value = True
    .Columns.Append "EMAIL", adVarWChar, 50
      .Columns("EMAIL").Properties("Nullable").Value = True
    .Columns.Append "ENTIDAD", adVarWChar, 4
      .Columns("ENTIDAD").Properties("Nullable").Value = True
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAX", adVarWChar, 17
      .Columns("FAX").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FCOBRO", adUnsignedTinyInt
      .Columns("FCOBRO").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "FOTO", adLongVarBinary
      .Columns("FOTO").Properties("Nullable").Value = True
    .Columns.Append "IMPUESTOS", adUnsignedTinyInt
      .Columns("IMPUESTOS").Properties("Nullable").Value = False
    .Columns.Append "LOCENVIO", adVarWChar, 40
      .Columns("LOCENVIO").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "NIF", adVarWChar, 15
      .Columns("NIF").Properties("Nullable").Value = False
    .Columns.Append "PAIS", adVarWChar, 25
      .Columns("PAIS").Properties("Nullable").Value = True
    .Columns.Append "PAISENVIO", adVarWChar, 25
      .Columns("PAISENVIO").Properties("Nullable").Value = True
    .Columns.Append "PERCONTA", adVarWChar, 40
      .Columns("PERCONTA").Properties("Nullable").Value = True
    .Columns.Append "POBLACION", adVarWChar, 40
      .Columns("POBLACION").Properties("Nullable").Value = True
    .Columns.Append "PROVENVIO", adVarWChar, 25
      .Columns("PROVENVIO").Properties("Nullable").Value = True
    .Columns.Append "PROVINCIA", adVarWChar, 25
      .Columns("PROVINCIA").Properties("Nullable").Value = True
    .Columns.Append "RAZO", adVarWChar, 40
      .Columns("RAZO").Properties("Nullable").Value = True
    .Columns.Append "REPRESEN", adInteger
      .Columns("REPRESEN").Properties("Nullable").Value = True
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SUCURSAL", adVarWChar, 4
      .Columns("SUCURSAL").Properties("Nullable").Value = True
    .Columns.Append "TELCONTA", adVarWChar, 17
      .Columns("TELCONTA").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO1", adVarWChar, 17
      .Columns("TELEFONO1").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO2", adVarWChar, 17
      .Columns("TELEFONO2").Properties("Nullable").Value = True
    .Columns.Append "TITULAR", adVarWChar, 40
      .Columns("TITULAR").Properties("Nullable").Value = True
    .Columns.Append "WEB", adVarWChar, 50
      .Columns("WEB").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(5)

  '----------* Table Definition of COLORES *----------
  Set Tbl(6) = New ADOX.Table
  Tbl(6).ParentCatalog = Cat
  With Tbl(6)
    .Name = "COLORES"
    .Columns.Append "CODCOL", adInteger
      .Columns("CODCOL").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(6)

  '----------* Table Definition of COSTURE *----------
  Set Tbl(7) = New ADOX.Table
  Tbl(7).ParentCatalog = Cat
  With Tbl(7)
    .Name = "COSTURE"
    .Columns.Append "CODBAN", adSmallInt
      .Columns("CODBAN").Properties("Nullable").Value = True
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adVarWChar, 5
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMENTARIO", adLongVarWChar
      .Columns("COMENTARIO").Properties("Nullable").Value = True
    .Columns.Append "CUENTA", adVarWChar, 10
      .Columns("CUENTA").Properties("Nullable").Value = True
    .Columns.Append "DC", adVarWChar, 2
      .Columns("DC").Properties("Nullable").Value = True
    .Columns.Append "DIRECCION", adVarWChar, 40
      .Columns("DIRECCION").Properties("Nullable").Value = True
    .Columns.Append "EMAIL", adVarWChar, 50
      .Columns("EMAIL").Properties("Nullable").Value = True
    .Columns.Append "ENTIDAD", adVarWChar, 4
      .Columns("ENTIDAD").Properties("Nullable").Value = True
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = True
    .Columns.Append "FAX", adVarWChar, 17
      .Columns("FAX").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "FPAGO", adUnsignedTinyInt
      .Columns("FPAGO").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "NIF", adVarWChar, 15
      .Columns("NIF").Properties("Nullable").Value = False
    .Columns.Append "NOMBRE", adVarWChar, 50
      .Columns("NOMBRE").Properties("Nullable").Value = False
    .Columns.Append "PAIS", adVarWChar, 25
      .Columns("PAIS").Properties("Nullable").Value = True
    .Columns.Append "POBLACION", adVarWChar, 40
      .Columns("POBLACION").Properties("Nullable").Value = True
    .Columns.Append "PROVINCIA", adVarWChar, 25
      .Columns("PROVINCIA").Properties("Nullable").Value = True
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SUCURSAL", adVarWChar, 4
      .Columns("SUCURSAL").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO1", adVarWChar, 17
      .Columns("TELEFONO1").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO2", adVarWChar, 17
      .Columns("TELEFONO2").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(7)

  '----------* Table Definition of DETTRANS *----------
  Set Tbl(8) = New ADOX.Table
  Tbl(8).ParentCatalog = Cat
  With Tbl(8)
    .Name = "DETTRANS"
    .Columns.Append "CODALM", adUnsignedTinyInt
      .Columns("CODALM").Properties("Nullable").Value = False
    .Columns.Append "CODART", adSmallInt
      .Columns("CODART").Properties("Nullable").Value = False
    .Columns.Append "CODCOL", adSmallInt
      .Columns("CODCOL").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adInteger
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODTALLA", adSmallInt
      .Columns("CODTALLA").Properties("Nullable").Value = False
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    
        .Columns.Append "RE", adSmallInt
      .Columns("RE").Properties("Nullable").Value = False
      
    .Columns.Append "ID", adInteger
      .Columns("ID").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
      .Columns("TEMPOR").Properties("Nullable").Value = False
    .Columns.Append "UNIDADES", adSingle
      .Columns("UNIDADES").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(8)

  '----------* Table Definition of FAMILIAS *----------
  Set Tbl(9) = New ADOX.Table
  Tbl(9).ParentCatalog = Cat
  With Tbl(9)
    .Name = "FAMILIAS"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODSEC", adSmallInt
      .Columns("CODSEC").Properties("Nullable").Value = True
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(9)

  '----------* Table Definition of MAARTIC *----------
  Set Tbl(10) = New ADOX.Table
  Tbl(10).ParentCatalog = Cat
  With Tbl(10)
    .Name = "MAARTIC"
    .Columns.Append "ABREVIA", adVarWChar, 20
      .Columns("ABREVIA").Properties("Nullable").Value = True
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPROV", adSmallInt
      .Columns("CODPROV").Properties("Nullable").Value = False
    .Columns.Append "COMEN", adLongVarWChar
      .Columns("COMEN").Properties("Nullable").Value = True
    .Columns.Append "DCTO", adSingle
      .Columns("DCTO").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAMILIA", adSmallInt
      .Columns("FAMILIA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "FOTO", adLongVarBinary
      .Columns("FOTO").Properties("Nullable").Value = True
    .Columns.Append "HIST", adBoolean
      .Columns("HIST").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "MODELO", adVarWChar, 30
      .Columns("MODELO").Properties("Nullable").Value = False
    .Columns.Append "PEDIR", adSingle
      .Columns("PEDIR").Properties("Nullable").Value = False
    .Columns.Append "PRECOM", adSingle
      .Columns("PRECOM").Properties("Nullable").Value = False
    .Columns.Append "PREVEN", adSingle
      .Columns("PREVEN").Properties("Nullable").Value = False
    .Columns.Append "REF", adVarWChar, 15
      .Columns("REF").Properties("Nullable").Value = True
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SECCION", adSmallInt
      .Columns("SECCION").Properties("Nullable").Value = True
    .Columns.Append "STOCK", adBoolean
      .Columns("STOCK").Properties("Nullable").Value = False
    .Columns.Append "STOCKMAX", adSingle
      .Columns("STOCKMAX").Properties("Nullable").Value = False
    .Columns.Append "STOCKMIN", adSingle
      .Columns("STOCKMIN").Properties("Nullable").Value = False
    .Columns.Append "SUBFAM", adSmallInt
      .Columns("SUBFAM").Properties("Nullable").Value = True
    .Columns.Append "TARIFA", adBoolean
      .Columns("TARIFA").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
      .Columns("TEMPOR").Properties("Nullable").Value = False
    .Columns.Append "TIPOIVA", adUnsignedTinyInt
      .Columns("TIPOIVA").Properties("Nullable").Value = False
          .Columns.Append "IVACOM", adUnsignedTinyInt
      .Columns("IVACOM").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(10)

  '----------* Table Definition of MAING *----------
  Set Tbl(11) = New ADOX.Table
  Tbl(11).ParentCatalog = Cat
  With Tbl(11)
    .Name = "MAING"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "COMENTARIO", adLongVarWChar
      .Columns("COMENTARIO").Properties("Nullable").Value = True
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(11)

  '----------* Table Definition of MAPAG *----------
  Set Tbl(12) = New ADOX.Table
  Tbl(12).ParentCatalog = Cat
  With Tbl(12)
    .Name = "MAPAG"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "COMENTARIO", adLongVarWChar
      .Columns("COMENTARIO").Properties("Nullable").Value = True
    .Columns.Append "DESCRIPCION", adVarWChar, 25
      .Columns("DESCRIPCION").Properties("Nullable").Value = True
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(12)

  '----------* Table Definition of MAPROV *----------
  Set Tbl(13) = New ADOX.Table
  Tbl(13).ParentCatalog = Cat
  With Tbl(13)
    .Name = "MAPROV"
    .Columns.Append "CCCUEN", adInteger
      .Columns("CCCUEN").Properties("Nullable").Value = False
    .Columns.Append "CCDC", adUnsignedTinyInt
      .Columns("CCDC").Properties("Nullable").Value = False
    .Columns.Append "CCENTI", adSmallInt
      .Columns("CCENTI").Properties("Nullable").Value = False
    .Columns.Append "CCOFICI", adSmallInt
      .Columns("CCOFICI").Properties("Nullable").Value = False
    .Columns.Append "CIF", adVarWChar, 12
      .Columns("CIF").Properties("Nullable").Value = False
    .Columns.Append "CODBAN", adSmallInt
      .Columns("CODBAN").Properties("Nullable").Value = True
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adSmallInt
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMEN", adLongVarWChar
      .Columns("COMEN").Properties("Nullable").Value = True
    .Columns.Append "DCTO", adSingle
      .Columns("DCTO").Properties("Nullable").Value = False
    .Columns.Append "DCTOPP", adSingle
      .Columns("DCTOPP").Properties("Nullable").Value = False
    .Columns.Append "DIRECC", adVarWChar, 20
      .Columns("DIRECC").Properties("Nullable").Value = True
    .Columns.Append "EXENTO", adBoolean
      .Columns("EXENTO").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAX", adVarWChar, 9
      .Columns("FAX").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "FPAGO", adUnsignedTinyInt
      .Columns("FPAGO").Properties("Nullable").Value = True
    .Columns.Append "LOCALI", adVarWChar, 10
      .Columns("LOCALI").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "NOMBRE", adVarWChar, 30
      .Columns("NOMBRE").Properties("Nullable").Value = False
    .Columns.Append "PERCON1", adVarWChar, 20
      .Columns("PERCON1").Properties("Nullable").Value = True
    .Columns.Append "PERCON2", adVarWChar, 20
      .Columns("PERCON2").Properties("Nullable").Value = True
    .Columns.Append "PROVIN", adVarWChar, 15
      .Columns("PROVIN").Properties("Nullable").Value = True
    .Columns.Append "RE", adBoolean
      .Columns("RE").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SECTOR", adUnsignedTinyInt
      .Columns("SECTOR").Properties("Nullable").Value = True
    .Columns.Append "TELFNO", adVarWChar, 9
      .Columns("TELFNO").Properties("Nullable").Value = True
  End With
  Cat.tables.Append Tbl(13)

  '----------* Table Definition of OFERTAS *----------
  Set Tbl(19) = New ADOX.Table
  Tbl(19).ParentCatalog = Cat
  With Tbl(19)
    .Name = "OFERTAS"
    .Columns.Append "CODART", adSmallInt
      .Columns("CODART").Properties("Nullable").Value = False
    .Columns.Append "CODCAJA", adUnsignedTinyInt
      .Columns("CODCAJA").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adInteger
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPROV", adInteger
      .Columns("CODPROV").Properties("Nullable").Value = False
    .Columns.Append "DCTO", adUnsignedTinyInt
      .Columns("DCTO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 30
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FAMILIA", adSmallInt
      .Columns("FAMILIA").Properties("Nullable").Value = False
    .Columns.Append "FFIN", adDate
      .Columns("FFIN").Properties("Nullable").Value = False
    .Columns.Append "FINICIO", adDate
      .Columns("FINICIO").Properties("Nullable").Value = False
    .Columns.Append "IMPORTE", adSingle
      .Columns("IMPORTE").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "SECCION", adSmallInt
      .Columns("SECCION").Properties("Nullable").Value = False
    .Columns.Append "SUBFAM", adSmallInt
      .Columns("SUBFAM").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
      .Columns("TEMPOR").Properties("Nullable").Value = False
    .Columns.Append "TIPO", adUnsignedTinyInt
      .Columns("TIPO").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(19)

  '----------* Table Definition of PERSONAL *----------
  Set Tbl(20) = New ADOX.Table
  Tbl(20).ParentCatalog = Cat
  With Tbl(20)
    .Name = "PERSONAL"
    .Columns.Append "AFILIA", adVarWChar, 15
      .Columns("AFILIA").Properties("Nullable").Value = True
    .Columns.Append "ANTIGU", adDate
      .Columns("ANTIGU").Properties("Nullable").Value = True
    .Columns.Append "CLAVE", adVarWChar, 10
      .Columns("CLAVE").Properties("Nullable").Value = False
    .Columns.Append "CODBAN", adVarWChar, 40
      .Columns("CODBAN").Properties("Nullable").Value = True
    .Columns.Append "CODCAJA", adUnsignedTinyInt
      .Columns("CODCAJA").Properties("Nullable").Value = False
    .Columns.Append "CODCEN", adUnsignedTinyInt
      .Columns("CODCEN").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODPOS", adInteger
      .Columns("CODPOS").Properties("Nullable").Value = True
    .Columns.Append "COMENTARIO", adLongVarWChar
      .Columns("COMENTARIO").Properties("Nullable").Value = True
    .Columns.Append "CUENTA", adVarWChar, 10
      .Columns("CUENTA").Properties("Nullable").Value = True
    .Columns.Append "DC", adVarWChar, 2
      .Columns("DC").Properties("Nullable").Value = True
    .Columns.Append "DIRECCION", adVarWChar, 40
      .Columns("DIRECCION").Properties("Nullable").Value = True
    .Columns.Append "EMAIL", adVarWChar, 50
      .Columns("EMAIL").Properties("Nullable").Value = True
    .Columns.Append "ENTIDAD", adVarWChar, 4
      .Columns("ENTIDAD").Properties("Nullable").Value = True
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
    .Columns.Append "FAX", adVarWChar, 17
      .Columns("FAX").Properties("Nullable").Value = True
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "FOTO", adLongVarBinary
      .Columns("FOTO").Properties("Nullable").Value = True
    .Columns.Append "FPAGO", adUnsignedTinyInt
      .Columns("FPAGO").Properties("Nullable").Value = True
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "NIF", adVarWChar, 15
      .Columns("NIF").Properties("Nullable").Value = False
    .Columns.Append "NOMBRE", adVarWChar, 50
      .Columns("NOMBRE").Properties("Nullable").Value = True
    .Columns.Append "PAIS", adVarWChar, 25
      .Columns("PAIS").Properties("Nullable").Value = True
    .Columns.Append "POBLACION", adVarWChar, 40
      .Columns("POBLACION").Properties("Nullable").Value = True
    .Columns.Append "PROVINCIA", adVarWChar, 25
      .Columns("PROVINCIA").Properties("Nullable").Value = True
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "SUCURSAL", adVarWChar, 4
      .Columns("SUCURSAL").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO1", adVarWChar, 17
      .Columns("TELEFONO1").Properties("Nullable").Value = True
    .Columns.Append "TELEFONO2", adVarWChar, 17
      .Columns("TELEFONO2").Properties("Nullable").Value = True
    .Columns.Append "TIPPERM", adUnsignedTinyInt
      .Columns("TIPPERM").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(20)

  '----------* Table Definition of PTRANS *----------
  Set Tbl(21) = New ADOX.Table
  Tbl(21).ParentCatalog = Cat
  With Tbl(21)
    .Name = "PTRANS"
    .Columns.Append "CODALMDEST", adUnsignedTinyInt
      .Columns("CODALMDEST").Properties("Nullable").Value = False
    .Columns.Append "CODALMORIG", adUnsignedTinyInt
      .Columns("CODALMORIG").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adInteger
      .Columns("CODIGO").Properties("Nullable").Value = False
    
    .Columns.Append "CODUSR", adSmallInt
    .Columns("CODUSR").Properties("Nullable").Value = False
    
    .Columns.Append "NUMPED", adInteger
    .Columns("NUMPED").Properties("Nullable").Value = False
          
    .Columns.Append "DCTO", adUnsignedTinyInt
      .Columns("DCTO").Properties("Nullable").Value = False
    .Columns.Append "ESTADO", adUnsignedTinyInt
      .Columns("ESTADO").Properties("Nullable").Value = False
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "GASTOS", adSingle
      .Columns("GASTOS").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TOTAL", adSingle
      .Columns("TOTAL").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(21)

  '----------* Table Definition of PTRANSMSG *----------
  Set Tbl(22) = New ADOX.Table
  Tbl(22).ParentCatalog = Cat
  With Tbl(22)
    .Name = "PTRANSMSG"
    .Columns.Append "CODALM", adUnsignedTinyInt
      .Columns("CODALM").Properties("Nullable").Value = False
    .Columns.Append "CODALMORIG", adUnsignedTinyInt
      .Columns("CODALMORIG").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adInteger
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODUSR", adSmallInt
      .Columns("CODUSR").Properties("Nullable").Value = False
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "ID", adInteger
      .Columns("ID").Properties("Nullable").Value = False
    .Columns.Append "MSG", adLongVarWChar
      .Columns("MSG").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(22)

  '----------* Table Definition of SECCIONES *----------
  Set Tbl(23) = New ADOX.Table
  Tbl(23).ParentCatalog = Cat
  With Tbl(23)
    .Name = "SECCIONES"
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(23)

  '----------* Table Definition of SUBFAM *----------
  Set Tbl(24) = New ADOX.Table
  Tbl(24).ParentCatalog = Cat
  With Tbl(24)
    .Name = "SUBFAM"
    .Columns.Append "CODFAM", adSmallInt
      .Columns("CODFAM").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(24)

  '----------* Table Definition of TALLAS *----------
  Set Tbl(25) = New ADOX.Table
  Tbl(25).ParentCatalog = Cat
  With Tbl(25)
    .Name = "TALLAS"
    .Columns.Append "CATTALL", adInteger
      .Columns("CATTALL").Properties("Nullable").Value = False
    .Columns.Append "CODIGO", adSmallInt
      .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "DESCRIPCION", adVarWChar, 15
      .Columns("DESCRIPCION").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(25)

  '----------* Table Definition of TEMPOR *----------
  Set Tbl(26) = New ADOX.Table
  Tbl(26).ParentCatalog = Cat
  With Tbl(26)
    .Name = "TEMPOR"
    .Columns.Append "ABREVIA", adVarWChar, 5
      .Columns("ABREVIA").Properties("Nullable").Value = True
    .Columns.Append "ACTUAL", adBoolean
      .Columns("ACTUAL").Properties("Nullable").Value = False
    .Columns.Append "AÑO", adVarWChar, 4
      .Columns("AÑO").Properties("Nullable").Value = False
    .Columns.Append "FALTA", adDate 'adVarWChar, 20
      .Columns("FALTA").Properties("Nullable").Value = False
     .Columns.Append "FBAJA", adDate 'adVarWChar, 20
      .Columns("FBAJA").Properties("Nullable").Value = True
    .Columns.Append "FMODI", adDate 'adDBTimeStamp 'adVarWChar, 20
      .Columns("FMODI").Properties("Nullable").Value = False
    .Columns.Append "HIST", adBoolean
      .Columns("HIST").Properties("Nullable").Value = False
    .Columns.Append "IDTEM", adUnsignedTinyInt
      .Columns("IDTEM").Properties("Nullable").Value = False
    .Columns.Append "MBAJA", adBoolean
      .Columns("MBAJA").Properties("Nullable").Value = False
    .Columns.Append "rowguid", adGUID
      .Columns("rowguid").Properties("Nullable").Value = False
    .Columns.Append "TEMPORADA", adVarWChar, 10
      .Columns("TEMPORADA").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(26)
  
  
  '----------* Table Definition of CONF_TRN *----------
  Set Tbl(27) = New ADOX.Table
  Tbl(27).ParentCatalog = Cat
  With Tbl(27)
    .Name = "CONF_TRN"
    
    .Columns.Append "CODUSR", adSmallInt
    .Columns("CODUSR").Properties("Nullable").Value = False
    
    .Columns.Append "CODALMORIG", adSmallInt
    .Columns("CODALMORIG").Properties("Nullable").Value = False
    
    .Columns.Append "CODALMDEST", adSmallInt
    .Columns("CODALMDEST").Properties("Nullable").Value = False
    
    .Columns.Append "NUMTRANS", adSmallInt
    .Columns("NUMTRANS").Properties("Nullable").Value = False
    
    .Columns.Append "TOTAL", adSingle
    .Columns("TOTAL").Properties("Nullable").Value = False
    
    .Columns.Append "FHORA", adVarWChar, 20
    .Columns("FHORA").Properties("Nullable").Value = False

  End With
  Cat.tables.Append Tbl(27)
 

  Set Cat = Nothing
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


Public Function HayDisket() As Boolean

       On Error GoTo Error
       HayDisket = False
       ChDrive "A"
       HayDisket = True
       ChDrive "C" ' para que proximo intento funcione correctamente
        Exit Function
Error:
        ChDrive "C" ' para que proximo intento funcione correctamente
    Exit Function
    
    End Function
    




'______________________________________________________
' Conforma el WHERE para la búsqueda dependiendo del tipo
' de campo sobre el que se desee buscar (para el find)
' campo:  campo sobre el que se va a realizar la busqueda
' valor: valor que deseamos buscar (iobusqueda.text)
' tipobus: tipo de busqueda: =, likes,  (cbtipobus.text)
'______________________________________________________

Public Function ver_tipo_campo(campo As ADODB.Field) As Byte
Dim tipo As Byte
' tipo = 1  -> numerico
' tipo = 2  -> string
' tipo = 3  -> fecha/hora
' tipo = 4  -> booleano
' tipo = 0  -> no reconocido (tipo no valido para buscar)

'Debug.Print campo.Type

tipo = 0
ver_tipo_campo = 0

'If campo.Name = "MBAJA" Then
'Debug.Print "MBAJA"
'End If


Select Case campo.Type

Case adBigInt  'Un entero con signo de 8 bytes (DBTYPE_I8).
    tipo = 1

Case adBoolean 'Un valor Boolean (DBTYPE_BOOL).
    tipo = 4
Case adBSTR 'Una cadena de caracteres terminada en nulo (Unicode) (DBTYPE_BSTR).
    tipo = 2
Case adChar 'Un valor de tipo String (DBTYPE_STR).
    tipo = 2
Case adCurrency  'Un valor de tipo Currency (DBTYPE_CY). Un valor Currency es un número de coma fija con cuatro dígitos a la derecha del signo decimal. Se almacena en un entero con signo de 8 bytes en escala de 10.000.
    tipo = 1
Case adDate 'Un valor de tipo Date (DBTYPE_DATE). Un valor Date se almacena como un valor de tipo Double; la parte entera es el número de días transcurridos desde el 30 de diciembre de 1899 y la parte fraccionaria es la fracción de un día.
    tipo = 1
Case adDBDate 'Un valor de fecha (aaaammdd) (DBTYPE_DBDATE).
    tipo = 3
Case adDBTime 'Un valor de hora (hhmmss) (DBTYPE_DBTIME).
    tipo = 3
Case adDBTimeStamp 'Una marca de fecha y hora (aaaammddhhmmss más una fracción de miles de millones) (DBTYPE_DBTIMESTAMP).
    tipo = 3
Case adDecimal  'Un valor numérico exacto con una precisión y una escala fijas (DBTYPE_DECIMAL).
    tipo = 1
Case adInteger  'Un entero firmado de 4 bytes (DBTYPE_I4).
    tipo = 1
Case adNumeric 'Un valor numérico exacto con una precisión y una escala exactas (DBTYPE_NUMERIC).
    tipo = 1
Case adSingle  'Un valor de coma flotante de simple precisión (DBTYPE_R4).
    tipo = 1
Case adSmallInt  'Un entero con signo de 2 bytes (DBTYPE_I2).
    tipo = 1
Case adTinyInt  'Un entero con signo de 1 byte (DBTYPE_I1).
    tipo = 1
Case adUnsignedBigInt  'Un entero sin signo de 8 bytes (DBTYPE_UI8).
    tipo = 1
Case adUnsignedInt 'Un entero sin signo de 4 bytes (DBTYPE_UI4).
    tipo = 1
Case adUnsignedSmallInt 'Un entero sin signo de 2 bytes (DBTYPE_UI2).
    tipo = 1
Case adUnsignedTinyInt 'Un entero sin signo de 1 bytes (DBTYPE_UI1).
    tipo = 1
Case 5 'numerico
    tipo = 1
Case adVarChar 'texto
   tipo = 2
Case 200 'texto
    tipo = 2
Case 202  'texto
    tipo = 2
Case adGUID
    tipo = 2
       
End Select

ver_tipo_campo = tipo
'MsgBox tipo

End Function


Public Function FormatDrive(formulario As Form, ByVal DriveLetter As String, _
  Optional PermitNonRemovableFormat As Boolean = False) As _
  Boolean

'**************************************************
'Formats a drive specified by Drive Letter.
'Confirmation box will appear

'Set PermitNonRemovableFormat to true if you want to allow for _
 formating of fixed drive or other non-removable drive (e.g., C:\)


'Returns true if successful, false otherwise

'EXAMPLE 1: FormatDrive "A:\"
'formats drive A:

'EXAMPLE 2: FormatDrive "C:\"
'Will fail because PermitNonRemovableFormat is not set
'to true

'I have not tested formatting fixed drives because there
'are no fixed drives I want to format

'USE WITH CAUTION: IF YOU DON'T FOLLOW INSTRUCTIONS
'YOU CAN WIPE OUT SOMEONE'S HARD DRIVE

'**************************************************
Dim sDrive As String
Dim lDrive As Long
Dim iDriveType As Integer
Dim iAns As Integer
Dim sDriveLetter
Dim lRet As Long

sDrive = UCase(DriveLetter)
sDriveLetter = sDrive
'format as [Letter]:/ if not done already
If Len(sDrive) = 1 Then sDriveLetter = sDriveLetter & ":\"
If Len(sDrive) = 2 And Right$(sDrive, 1) = ":" _
    Then sDriveLetter = sDrive & "\"


lDrive = Asc(Left(sDrive, 1)) - 65
iDriveType = DriveType(sDrive)
Select Case iDriveType

Case 2

lRet = SHFormatDrive(formulario.hwnd, lDrive, &HFFFF, FORMAT_FULL)
FormatDrive = lRet = 0
Case 3, 4, 5, 6
    If Not PermitNonRemovableFormat Then Exit Function
    lRet = SHFormatDrive(formulario.hwnd, lDrive, &HFFFF, FORMAT_FULL)
    FormatDrive = lRet = 0
Case Else 'no such drive
    Exit Function
End Select

End Function

Private Function DriveType(Drive As String) As Integer

Dim sAns As String, lAns As Long

'fix bad parameter values
If Len(Drive) = 1 Then Drive = Drive & ":\"
If Len(Drive) = 2 And Right$(Drive, 1) = ":" _
    Then Drive = Drive & "\"

DriveType = GetDriveType(Drive)

End Function




'Un procedimiento para ejecutar y esperar a que termine
Public Sub ExecCmdNoFocus(ByVal CmdLine As String)
    'Esperar a que un proceso termine,
    'la ventana se mostrará minimizada sin foco
    Dim hProcess As Long
    Dim RetVal As Long

   On Error GoTo ExecCmdNoFocus_Error

    'The next line launches CmdLine as icon,
    'captures process ID
    'MsgBox CmdLine
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, _
            Shell(CmdLine, vbNormalFocus))
    Do
        'Get the status of the process
        GetExitCodeProcess hProcess, RetVal
        'Sleep command recommended as well
        'as DoEvents
        DoEvents
        Sleep 100
    'Loop while the process is active
    Loop While RetVal = STILL_ACTIVE

   On Error GoTo 0
   Exit Sub

ExecCmdNoFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ExecCmdNoFocus de Formulario frmGenera"
End Sub






