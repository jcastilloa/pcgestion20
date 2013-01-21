Attribute VB_Name = "Rutinas_Ventas"
'---------------------------------------------------------------------------------------
' Modulo        : Rutinas_Ventas
' Fecha/Hora  : 18/01/2004 12:37
' Autor          : JCASTILLO
' Propósito    : Rutinas para el proceso de ventas y facturación en general
'---------------------------------------------------------------------------------------
Option Explicit

'Tipo para guardar los datos del cierre de caja
Public Type Datos_Cierre

    t_contado As Double       'total contado
    n_contado As Long         'numero de ventas contado
    n_vales_acep As Long     'numero de vales aceptados
    t_vales_acep As Double   'total vales aceptados
    n_vales_emi As Long      'numero de vales emitidos
    t_vales_emi As Double    'total vales emitidos
    
   ''''''''''''''''''''''''''''''''''''
    n_valdctoe As Long     'vales de descuento emitidos
    t_valdctoe As Double   'y aceptados
    n_valdctoa As Long
    t_valdctoa As Double
    
    n_difcampr As Long     'total diferencias por cambio de precio
    t_difcampr As Double  'importe diferencias por cambio de precio
    
    n_comtar As Long        'nº de comisiones tarjeta
    t_comtar As Double     'total comisiones tarjetas
     
   ''''''''''''''''''''''''''''''''
    
    n_tarjeta As Long         'numero de tarjetas
    t_tarjeta As Double       'importe tarjeta
    n_devol As Long           'numero de devoluciones
    t_devol As Double         'importe de devoluciones
    n_arreglos As Long        'numero de arreglos
    T_Arreglos As Double      'importe de arreglos
    T_ArreCon As Double      'importe de arreglos que se pagaron al contado
    n_movi As Long            'numero de movimientos
    t_movi As Double          'importe de movimientos
    t_caja As Double          'total en caja (REAL)
    t_caja_Teorico As Double  'total en caja (registrado por programa)
    n_pagos As Long           'numero de pagos
    t_pagos As Double         'total en pagos
    n_ingresos As Long        'numero de ingresos
    t_ingresos As Double      'total en ingresos
    n_deudc As Long           'numero de deudas cliente
    t_deudc As Double         'total en deudas cliente
    n_deudc_pag As Long     'numero de deudas cliente pagadas hoy
    t_deudc_pag As Double  'total en deudas cliente pagadas hoy
    Total_A As Double         'importe total en caja A
    Total_B As Double         'importe total en caja B
    Total_Efectivo As Double  'importe total en efectivo
    n_dcto As Long           'numero de prendas con descuentos
    t_dcto As Double          'importe total de los descuentos
    n_iva As Long           'numero de prendas con descuentos
    t_iva As Double            'importe total del iva
    Saldo_Caja_Efectivo As Double
    Ventas_Netas As Double
    Cobros_Realizados As Double
    Ventas_Brutas As Double  'ventas brutas (sin descuento ni iva)
    Se_Ha_Calculado As Boolean  'si se ha calculado el cierre
    
End Type

'---------------------------------------------------------------------------------------
' Subrutina   : añadir_vale
' Fecha/Hora  : 18/01/2004 12:40
' Autor       : JCASTILLO
' Propósito   : Añade un nuevo vale a la base de datos. Devuelve el codigo de vale
'                   Si me devuelve un 0, condición de error
'---------------------------------------------------------------------------------------
Public Function añadir_vale(V_Codigo_Venta, V_Dependiente, V_Cliente, V_CajaCliente, v_importe, v_dcto, V_Estado, V_Tipo, V_Caduca As Variant, conexion As ADODB.Connection) As Long
Dim tmp_Codvale As Variant

   On Error GoTo añadir_vale_Error

      'CODIGO DE VALE Q LE CORRESPONDE
      tmp_Codvale = devuelve_campo("SELECT max(CODIGO) + 1 from VALES WHERE CODCAJA = " & CajaActual, conexion)
      
      If tmp_Codvale = "@" Then tmp_Codvale = 1
      
      'INSERTAR VALE POR SQL
      
      'ESTADO: 0=pendiente, 1=aceptado
      'TIPO: 1=VENTA, 2=DEVOLUCION, 3=SOBRANTE
      
      If IsNull(V_Cliente) Then V_Cliente = 0
      If IsNull(V_CajaCliente) Then V_CajaCliente = 0
          
      
      If IsNull(V_Caduca) Then
         
       conexion.Execute "INSERT INTO VALES (CODIGO, CODCAJA, CODVEN, CODPER, CODCLI, CAJACLI, IMPORTE, DCTO, ESTADO, TIPO) values (" & tmp_Codvale & "," & _
       CajaActual & ", " & V_Codigo_Venta & ", " & V_Dependiente & ", " & V_Cliente & ", " & V_CajaCliente & ", " & Replace(CStr(v_importe), ",", ".") & ", " & v_dcto & ", " & V_Estado & ", " & V_Tipo & ")"
      
      Else
      
       conexion.Execute "INSERT INTO VALES (CODIGO, CODCAJA, CODVEN, CODPER, CODCLI, CAJACLI, IMPORTE, DCTO, ESTADO, TIPO, CADUCA) values (" & tmp_Codvale & "," & _
       CajaActual & ", " & V_Codigo_Venta & ", " & V_Dependiente & ", " & V_Cliente & ", " & V_CajaCliente & ", " & Replace(CStr(v_importe), ",", ".") & ", " & v_dcto & ", " & V_Estado & ", " & V_Tipo & ", " & V_Caduca & ")"
     
      End If
      
      añadir_vale = tmp_Codvale

   On Error GoTo 0
   Exit Function

añadir_vale_Error:

    añadir_vale = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añadir_vale de Módulo Rutinas_Ventas"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : Imprime_Vale
' Fecha/Hora    : 27/01/2004 12:36
' Autor         : JCastillo
' Propósito     :    Imprime el vale especificado
'---------------------------------------------------------------------------------------
Public Sub Imprime_Vale(numvale As Long, codcaja As Byte, conexion As ADODB.Connection)
Dim rcvale As New ADODB.Recordset

Dim v_codcli As Long
Dim v_codcajacli As Long

Dim v_codper As Long

Dim v_importe As Single
Dim v_dcto As Single

Dim v_caducidad As String

Dim tmpstr As String

   On Error GoTo Imprime_Vale_Error
                      
        With rcvale
        
        .Open "SELECT * FROM VALES WHERE CODIGO = " & numvale & " AND CODCAJA = " & codcaja, conexion, adOpenDynamic, adLockReadOnly
        
        If Not IsNull(.fields("codcli")) Then v_codcli = .fields("codcli")
        If Not IsNull(.fields("cajacli")) Then v_codcajacli = .fields("cajacli")
        If Not IsNull(.fields("CODPER")) Then v_codper = .fields("CODPER")
        If Not IsNull(.fields("importe")) Then v_importe = .fields("IMPORTE")
        If Not IsNull(.fields("dcto")) Then v_dcto = .fields("DCTO")
        If Not IsNull(.fields("CADUCA")) Then v_caducidad = Format(.fields("CADUCA"), "dd/mm/yyyy")
        
        If .fields("TIPO") = 2 Then
            tmpstr = " - DEVOLUCION -"
        Else
            tmpstr = ""
        End If
        
        .Close
       
        Printer.Print ""
        Printer.Print ""
        Printer.FontSize = Printer.FontSize + 3
        Printer.Print "FUSION S.L."
        Printer.Print ""
        Printer.FontSize = Printer.FontSize - 3
        Printer.Print numvale & Format(codcaja, "000") & tmpstr
        
        
        Printer.Print ""
        Printer.Print ""
        
        
        If v_importe > 0 Then
            Printer.Print "VALE por un descuento de " & Format(v_importe, "Currency") & " en su"
            Printer.Print "próxima compra."
        Else
            Printer.Print "VALE por un descuento del " & v_dcto & "% en su próxima compra."
        End If
        
        Printer.Print ""
        Printer.Print "Emitido: " & UCase(Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & v_codper, conexion)))
        Printer.Print "Origen: " & UCase(Trim(devuelve_campo("SELECT DESCRIPCION FROM CENTROS WHERE CODIGO = " & CentroActual, conexion)))
        Printer.Print "Fecha/Hora: " & Now
                 
        'si tiene cliente
        If v_codcli > 0 Then
         Printer.Print ""
         Printer.Print "Cliente: " & UCase(Trim(devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & v_codcli & " AND CODCAJA = " & v_codcajacli, conexion)))
        End If
        
        Printer.Print ""
        Printer.Print ""
        
        'si tiene caducidad
        If Trim(v_caducidad) <> "" Then
         Printer.Print ""
         Printer.Print "Caduca en: " & v_caducidad
        End If
       
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print "."
        
       
        Printer.EndDoc
        
        End With
        Set rcvale = Nothing
        
        v_codcli = 0
        v_codcajacli = 0
        v_codper = 0
        v_importe = 0
        v_dcto = 0
        v_caducidad = ""
        tmpstr = ""
        
   On Error GoTo 0
   Exit Sub

Imprime_Vale_Error:

    Printer.EndDoc
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Imprime_Vale de Módulo Rutinas_Ventas"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : aceptar_vale
' Fecha/Hora  : 08/02/2004 19:22
' Autor       : JCASTILLO
' Propósito   : Acepta el vale seleccionado
'---------------------------------------------------------------------------------------
Public Function aceptar_vale(codvale As Long, codcaja As Byte, conexion As ADODB.Connection) As Boolean
    
   On Error GoTo aceptar_vale_Error

        conexion.Execute "UPDATE VALES SET ESTADO = 1, FACEP = '" & Format(Date, "yyyymmdd") & "' WHERE CODIGO = " & codvale & " AND CODCAJA = " & codcaja

   On Error GoTo 0
   Exit Function

aceptar_vale_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento aceptar_vale de Módulo Rutinas_Ventas"
                         
End Function

'Lo convierte en un string del tipo : CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)
Public Function Convierte_CodTicket(codigo_ticket As Long) As String
Dim t_codcaja As Byte
Dim t_codven As Long
Dim s_codven As String * 10
Dim s_codcaja As String * 3
'CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)

t_codcaja = Right(codigo_ticket, 3)
t_codven = Left(codigo_ticket, Len(codigo_ticket) - 3)
'coger los 3 ultimos dígitos (código de caja)

'sacar el código de caja
Mid(s_codven, 1, Len(t_codven)) = t_codven
Mid(s_codcaja, 1, Len(t_codcaja)) = t_codcaja

Convierte_CodTicket = s_codven & s_codcaja

End Function



'---------------------------------------------------------------------------------------
' Procedimiento : calcula_cierre_caja
' Fecha/Hora    : 12/02/2004 08:59
' Autor         : JCastillo
' Propósito     :  Calcula el cierre para la fecha y graba (o actualiza) un registro
'                      en CIERREDIA
'                      Devuelve   estructura Datos_Cierre
'
'---------------------------------------------------------------------------------------
Public Function calcula_cierre_caja(fechaYYYYMMDD As String, codcaja As Byte, conexion As ADODB.Connection) As Datos_Cierre
Dim mifecha As String

Dim C As Datos_Cierre
Dim tmpvar As Variant
Dim ffin As Date

'Dim t_contado As Double     'total contado
'dim n_contado as long         'numero de ventas al contado
'Dim n_vales As Long           'numero de vales
'Dim t_vales As Double         'total vales
'Dim n_vtar As Long             'numero de tarjetas
'Dim t_tarjeta As Double       'importe tarjeta
'Dim n_devol As Long           'numero de devoluciones
'Dim t_devol As Double         'importe de devoluciones
'Dim n_arreglos As Long        'numero de arreglos
'Dim t_arreglos As Double     'importe de arreglos
'Dim n_movi As Long             'numero de movimientos
'Dim t_caja As Double           'total en caja
'Dim n_pagos As Long           'numero de pagos
'Dim t_pagos As Double         'total en pagos
'Dim n_deudc As Long           'numero de deudas cliente
'Dim t_deudc As Double         'total en deudas cliente
Dim m_matriz As Variant
Dim nuefech As String

Dim pA As Byte                  'porcentajes A y B %
Dim pB As Byte


On Error GoTo calcula_cierre_caja_Error

'1 CONTADO
'2 TARJETA
'3 CONTADO Y TARJETA
'4 VALE
'5 CONTADO Y VALE
'6 TARJETA Y VALE
'7 PENDIENTE
'8 CONTADO Y PENDIENTE
'9 PENDIENTE Y TARJETA
'10 PENDIENTE Y VALE

   'pasar a dd/mm/yyyy
   mifecha = Right(fechaYYYYMMDD, 2) & "/" & Mid(fechaYYYYMMDD, 5, 2) & "/" & Left(fechaYYYYMMDD, 4)

   nuefech = DateAdd("d", 1, mifecha)

   'pasar otra vez a yyyymmdd
   nuefech = Format(Year(nuefech), "0000") & Format(Month(nuefech), "00") & Format(Day(nuefech), "00")
   
   'sacar porcentajes A B
   m_matriz = devuelve_matriz("SELECT CAJA_A, CAJA_B from CAJAS WHERE CODIGO = " & codcaja, conexion)

    'obtener porcentajes
    pA = m_matriz(0)
    pB = m_matriz(1)

   ReDim m_matriz(0)
   
   
   'SACAR TOTALES DE LA TABLA CABVENTA:
   
   '******************** CONTADO ***************************************************

   'total contado  (importe primera para formas de pago 1, 3 , 5 y 8) (ventas aceptadas)
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (1, 3, 5, 8, 12) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
   'C.t_contado = devuelve_campo("SELECT sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (1, 3, 5) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)

   If Not IsNull(m_matriz(1)) Then C.t_contado = m_matriz(1)
   C.n_contado = m_matriz(0)
   
   ReDim m_matriz(0)
   
   '******************** VALES ACEPTADOS  ***************************************************
   'no incluir TIPO = 1
    
   'numero de vales
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE (FACEP >= '" & fechaYYYYMMDD & "') AND (FACEP < '" & nuefech & "') AND TIPO NOT IN (1,4) AND (CODCAJA = " & codcaja & ") AND (ESTADO = 1)", locCnn)
   'm_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1 AND TIPO <> 2", locCnn)
   
   
   If Not IsNull(m_matriz(1)) Then C.t_vales_acep = m_matriz(1)
   C.n_vales_acep = m_matriz(0)
   
   ReDim m_matriz(0)
   
   
   '******************** VALES ACEPTADOS POR DESCUENTO  ***************************************************
   'Solo Tipo = 1
   
   'numero de vales
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE (FACEP >= '" & fechaYYYYMMDD & "') AND (FACEP < '" & nuefech & "') AND TIPO = 1 AND (CODCAJA = " & codcaja & ") AND (ESTADO = 1)", locCnn)
   'm_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1 AND TIPO <> 2", locCnn)
   
   
   If Not IsNull(m_matriz(1)) Then C.t_valdctoa = m_matriz(1)
   C.n_valdctoa = m_matriz(0)
   
   ReDim m_matriz(0)
   
   '******************** VALES EMITIDOS  ***************************************************
    
   'SE OMITEN LOS VALES DE TIPO 1 (sacados por el formulario de crear vale)
    
   'numero de vales
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE (FMODI >= '" & fechaYYYYMMDD & "') AND (FMODI < '" & nuefech & "') AND  TIPO NOT IN(1, 4) AND (CODCAJA = " & codcaja & ")", locCnn)
   'm_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1 AND TIPO <> 2", locCnn)
      
   If Not IsNull(m_matriz(1)) Then C.t_vales_emi = m_matriz(1)
   C.n_vales_emi = m_matriz(0)
       
   ReDim m_matriz(0)
   'Tarjeta de Crédito
        
        
   '******************** VALES EMITIDOS  DE DESCUENTO  ***************************************************
    
   'SOLO TIPO 1 (sacados por el formulario de crear vale)
    
   'numero de vales
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE (FMODI >= '" & fechaYYYYMMDD & "') AND (FMODI < '" & nuefech & "') AND  TIPO = 1 AND (CODCAJA = " & codcaja & ")", locCnn)
   'm_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM VALES WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1 AND TIPO <> 2", locCnn)
      
   If Not IsNull(m_matriz(1)) Then C.t_valdctoe = m_matriz(1)
   C.n_valdctoe = m_matriz(0)
       
   ReDim m_matriz(0)
   
   
   'Tarjeta de Crédito

  '******************** TARJETA ***************************************************
   'total tarjeta
   
   'tmpvar = devuelve_campo("SELECT sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (2, 6)  AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
   
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (2, 6, 13)  AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
      
   If Not IsNull(m_matriz(1)) Then C.t_tarjeta = m_matriz(1)
   C.n_tarjeta = m_matriz(0)
   
   ReDim m_matriz(0)
   
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMP_SEGUNDA) FROM CABVENTA WHERE FCOBRO IN (3, 9) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
   
   If Not IsNull(m_matriz(1)) Then C.t_tarjeta = C.t_tarjeta + m_matriz(1)
   C.n_tarjeta = C.n_tarjeta + m_matriz(0)
   
   'If tmpvar <> "@" Then
   ' C.t_tarjeta = tmpvar
   'Else
   ' C.t_tarjeta = 0
   'End If
   
   'tmpvar = devuelve_campo("SELECT sum(IMP_SEGUNDA) FROM CABVENTA WHERE FCOBRO IN (3, 9) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
   
  'If tmpvar <> "@" Then
   ' C.t_tarjeta = C.t_tarjeta + tmpvar
   'Else
   ' C.t_tarjeta = C.t_tarjeta
  'End If
   
  ' C.n_tarjeta = devuelve_campo("SELECT COUNT(CODIGO) FROM CABVENTA WHERE FCOBRO IN (2, 3, 6)  AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
     
     
     
   ReDim m_matriz(0)
     
   '******************** DEVOLUCIONES ***************************************************
   m_matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM DEVOL WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND MBAJA =0", locCnn)

   'M_Matriz = devuelve_matriz("SELECT count(CODIGO), sum(IMPORTE) FROM DEVOL WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJA = " & codcaja & " AND MBAJA =0", locCnn)
   C.n_devol = m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_devol = m_matriz(1)
   
   ReDim m_matriz(0)
   
   '******************** ARREGLOS **************************************************
   m_matriz = devuelve_matriz("SELECT count(ID), sum(PVP) FROM ARREGLOS WHERE FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND CODCAJ = " & codcaja & " AND ESTADO = 2", locCnn)
   If Not IsNull(m_matriz(1)) Then C.T_Arreglos = m_matriz(1)
   C.n_arreglos = m_matriz(0)
   
   ReDim m_matriz(0)
   
   '******************** ARREGLOS PAGADOS AL CONTADO *************************************************
   m_matriz = devuelve_matriz("SELECT sum(PVP) FROM ARREGLOS WHERE (FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "') AND (CODCAJ = " & codcaja & ") AND (CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE FCOBRO IN(1, 3, 5, 8, 12) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1))", locCnn)
   If Not IsNull(m_matriz(0)) Then C.T_ArreCon = m_matriz(0)
   'C.n_arreglos = m_matriz(0)
   
   ReDim m_matriz(0)
   

   '******************** INGRESOS ***************************************************
   
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMPORTE) FROM INGRESOS WHERE  FALTA >= '" & fechaYYYYMMDD & "' AND FALTA < '" & nuefech & "' AND CODCAJA = " & codcaja, locCnn)
   C.n_ingresos = m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_ingresos = m_matriz(1)
   
   ReDim m_matriz(0)
   
   '******************** PAGOS ***************************************************
   'm_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMPORTE) FROM DETPAGOS WHERE  FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND (MBAJA = 0) AND CODCAJA = " & codcaja, locCnn)
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMPORTE) FROM PAGOS WHERE  FMODI >= '" & fechaYYYYMMDD & "' AND FMODI < '" & nuefech & "' AND (MBAJA = 0) AND (TIPOPAGO>1) AND CODCAJA = " & codcaja, locCnn)
 
   C.n_pagos = m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_pagos = m_matriz(1)
   
   ReDim m_matriz(0)
   
    
    '******************** DEUDAS CLIENTES (VTAS PENDIENTES) ***************************************************
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (7, 9, 10) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND (ESTADO = 1)", locCnn)
   C.n_deudc = m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_deudc = m_matriz(1)
   
   ReDim m_matriz(0)
   
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMP_SEGUNDA) FROM CABVENTA WHERE (FCOBRO = 8) AND FHORA = '" & fechaYYYYMMDD & "' AND (CODCAJA = " & codcaja & ") AND (ESTADO = 1)", locCnn)
   
   C.n_deudc = C.n_deudc + m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_deudc = C.t_deudc + m_matriz(1)
   
   ReDim m_matriz(0)
   
   
   '******************** DEUDAS COBRADAS (VTAS PENDIENTES COBRADAS) ***************************************************
   ffin = CDate(Right(fechaYYYYMMDD, 2) & "/" & Mid(fechaYYYYMMDD, 5, 2) & "/" & Left(fechaYYYYMMDD, 4))
   ffin = DateAdd("d", 1, ffin)
   
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMPORTE) FROM DETDEUDCLI WHERE FMODI >= '" & fechaYYYYMMDD & "'  AND FMODI < '" & Format(ffin, "yyyymmdd") & "' AND (MBAJA = 0) AND CODCAJA = " & codcaja, locCnn)
   C.n_deudc_pag = m_matriz(0)
   If Not IsNull(m_matriz(1)) Then C.t_deudc_pag = m_matriz(1)
   
   ReDim m_matriz(0)
   
    
   '******************** MOVIMIENTOS CAJA  ***************************************************
   m_matriz = devuelve_matriz("SELECT COUNT(CODIGO), sum(IMPORTE) FROM MOVCAJA WHERE  FCIERRE >= '" & fechaYYYYMMDD & "' AND FCIERRE < '" & nuefech & "' AND CODCAJA = " & codcaja, locCnn)
  
   C.n_movi = m_matriz(0)
   'C.t_movi = M_Matriz(1)
      
   If Not IsNull(m_matriz(1)) Then
       C.t_movi = m_matriz(1)
   'Else
   '     C.t_caja_Teorico = 0
   End If
   
   '******************** TOTAL DCTO  ***************************************************
   m_matriz = devuelve_matriz("SELECT COUNT(CODVEN), sum(((PREVEN * UNIDADES) * DCTO) / 100) FROM DETVENTA WHERE (CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & ")) AND (CODCAJA = " & codcaja & ") AND (DCTO > 0)", locCnn)
  
   C.n_dcto = m_matriz(0)
      
   If Not IsNull(m_matriz(1)) Then
       C.t_dcto = m_matriz(1)
   Else
        C.t_dcto = 0
   End If
   
      '******************** TOTAL IVA  ***************************************************
   m_matriz = devuelve_matriz("SELECT COUNT(CODVEN), sum(((PREVEN * UNIDADES) * IVA) / 100) FROM DETVENTA WHERE (CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & ")) AND (CODCAJA = " & codcaja & ") AND (IVA > 0)", locCnn)
  
   C.n_iva = m_matriz(0)
      
   If Not IsNull(m_matriz(1)) Then
       C.t_iva = m_matriz(1)
   Else
        C.t_iva = 0
   End If
   
   '*****TOTAL CAJA TEORICO*****:   MOVIMIENTOS CAJA (la diferencia) + CONTADO
   C.t_caja_Teorico = (C.t_movi + C.t_contado)
   
   '******************** TOTAL CAJA A  ***************************************************
  
  'sacar el contado de la caja A
   
  ' tmpvar = devuelve_campo("SELECT sum(IMP_PRIMERA + IMP_SEGUNDA) FROM CABVENTA WHERE (CAJADES = 0)  AND (FHORA = '" & fechaYYYYMMDD & "') AND (CODCAJA = " & codcaja & ") AND (ESTADO = 1)", locCnn)
    
  ' If tmpvar <> "@" Then
  '  C.Total_A = tmpvar
  ' Else
  '  C.Total_A = 0
  ' End If
    
   'quitar arreglos del total A
  ' C.Total_A = (C.Total_A - C.T_Arreglos)
      
   '******************** TOTAL CAJA B  ***************************************************
 '  tmpvar = devuelve_campo("SELECT sum(IMP_PRIMERA + IMP_SEGUNDA) FROM CABVENTA WHERE (CAJADES = 1)  AND (FHORA = '" & fechaYYYYMMDD & "') AND (CODCAJA = " & codcaja & ") AND (ESTADO = 1)", locCnn)

 
  ' If tmpvar <> "@" Then
  '  C.Total_B = tmpvar
  ' Else
  '  C.Total_B = 0
  ' End If
       
      
      
   'SELECT sum(UNIDADES * MAARTIC.PREVEN)
   'FROM DETVENTA INNER JOIN MAARTIC ON (DETVENTA.CODART = MAARTIC.CODIGO) AND (DETVENTA.TEMPOR = MAARTIC.TEMPOR);
      
      
      
   '******************** VENTAS BRUTAS  ***************************************************
   ' y las diferencias por cambio de precios ...
   
   
   'Cálculo con el PVP final de la venta
   '(cambiar por el otro calculo cuando este echo el cambio)
   'tmpvar = devuelve_campo("SELECT sum(UNIDADES * PREVEN) FROM DETVENTA WHERE CODCAJA = " & codcaja & "  AND CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE (FHORA = '" & fechaYYYYMMDD & "') AND (CODCAJA = " & codcaja & "))", locCnn)
   
   
   'If tmpvar <> "@" Then
   ' C.Ventas_Brutas = tmpvar
   'Else
    'C.Ventas_Brutas = 0
   'End If
   
   'Set tmpvar = Nothing
   
  '////////////////////////////////// VENTAS BRUTAS CON EL PVP ORIGINAL ///////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  ' 'CALCULO BUENO, activar esto cuando este echo el cambio
   m_matriz = devuelve_matriz("SELECT sum(UNIDADES * MAARTIC.PREVEN), sum(UNIDADES * DETVENTA.PREVEN) FROM DETVENTA INNER JOIN MAARTIC ON (DETVENTA.CODART = MAARTIC.CODIGO) AND (DETVENTA.TEMPOR = MAARTIC.TEMPOR) WHERE CODCAJA = " & codcaja & "  AND CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE (FHORA = '" & fechaYYYYMMDD & "') AND (CODCAJA = " & codcaja & "))", locCnn)
 
   If Not IsNull(m_matriz(0)) Then
       C.Ventas_Brutas = m_matriz(0)
   Else
        C.Ventas_Brutas = 0
   End If
   ' 'sacar la diferencia por cambio de precio
   If Not IsNull(m_matriz(1)) Then
          C.t_difcampr = m_matriz(1) - m_matriz(0)
          
   End If
   
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  'número de articulos para los que ha cambiado el precio
  '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   ReDim m_matriz(0)
   
   tmpvar = devuelve_campo("SELECT count(*) FROM (DETVENTA INNER JOIN MAARTIC ON (DETVENTA.TEMPOR = MAARTIC.TEMPOR) AND (DETVENTA.CODART = MAARTIC.CODIGO)) INNER JOIN CABVENTA ON (DETVENTA.CODCAJA = CABVENTA.CODCAJA) AND (DETVENTA.CODVEN = CABVENTA.CODIGO) WHERE ((MAARTIC.PREVEN<> DETVENTA.PREVEN)) and (CABVENTA.FHORA = '" & fechaYYYYMMDD & "') and CABVENTA.CODCAJA =" & CajaActual, locCnn)
   
   If tmpvar <> "@" Then
    C.n_difcampr = tmpvar
   End If
      
   'SELECT DETVENTA.CODART, DETVENTA.TEMPOR, DETVENTA.PREVEN, MAARTIC.PREVEN, CABVENTA.FHORA FROM (DETVENTA INNER JOIN MAARTIC ON (DETVENTA.TEMPOR = MAARTIC.TEMPOR) AND (DETVENTA.CODART = MAARTIC.CODIGO)) INNER JOIN CABVENTA ON (DETVENTA.CODCAJA = CABVENTA.CODCAJA) AND (DETVENTA.CODVEN = CABVENTA.CODIGO) Where (((MAARTIC.Preven) <> DETVENTA.Preven))
  
   ReDim m_matriz(0)
   
   '******************** Comisiones Tarjetas ***************************************************

   'total contado  (importe primera para formas de pago 1, 3 , 5 y 8) (ventas aceptadas)
   m_matriz = devuelve_matriz("SELECT count(COMISION), sum(COMISION) FROM CABVENTA WHERE (COMISION > 0) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)
   'C.t_contado = devuelve_campo("SELECT sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (1, 3, 5) AND FHORA = '" & fechaYYYYMMDD & "' AND CODCAJA = " & codcaja & " AND ESTADO = 1", locCnn)

   If Not IsNull(m_matriz(1)) Then C.t_comtar = m_matriz(1)
   C.n_comtar = m_matriz(0)
   
   ReDim m_matriz(0)
        
   C = Calcula_formulas_cierre(C, False)
   
   'sacar porcentajes
   C.Total_A = (C.Ventas_Netas * pA) / 100
   C.Total_B = (C.Ventas_Netas * pB) / 100
   
   'asegurarse de que como minimo va en A el importe de las tarjetas
   If C.Total_A < C.t_tarjeta Then
    'redistribuir
      C.Total_A = C.t_tarjeta
      C.Total_B = C.Ventas_Netas - C.t_tarjeta
   End If
   
   C.Se_Ha_Calculado = True
   
   calcula_cierre_caja = C
      
      
   On Error GoTo 0
   
   nuefech = ""
   
   Exit Function

calcula_cierre_caja_Error:

    C.Se_Ha_Calculado = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento calcula_cierre_caja de Módulo Rutinas_Ventas"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : Calcula_formulas_cierre
' Fecha/Hora     : 18/06/2004 17:47
' Autor             : JCastillo
' Propósito       : Calculo de las formulas del cierre.
'
'                  Si consulta = true no vuelve a calcular la formula de la tarjeta
'                  pues esta leyendo de cierredia y eso ya esta por tarjetas.
'---------------------------------------------------------------------------------------
'
Public Function Calcula_formulas_cierre(C As Datos_Cierre, Consulta As Boolean) As Datos_Cierre

   On Error GoTo Calcula_formulas_cierre_Error

   If Not Consulta Then
    C.t_tarjeta = (C.t_tarjeta - (C.T_Arreglos - C.T_ArreCon))
   End If
   
   C.Total_Efectivo = ((C.t_contado - C.T_ArreCon))
   
   
   'he quitado el sumatorio de arreglos (VER AQUI PUEDE HABER PROBLEMAS
   'si en el efectivo deben figurar los arreglos.
   C.Saldo_Caja_Efectivo = (C.Total_Efectivo - C.t_pagos) '+ C.T_Arreglos
      
   ' Ventas Brutas:
   ' Precio original de la venta.
   ' - Descuentos
   ' - Descuentos por vales (vales de dcto.)
   ' - Devoluciones
   ' - Diferencias de PVP (puede ser + o -)
   ' - Ventas PENDIENTES
   ' + Ventas PENDIENTES Cobradas (en el dia).
   'tiene como resultado las ventas netas
   
   C.Ventas_Netas = (C.Ventas_Brutas - C.t_devol - C.t_dcto - C.t_deudc + C.t_deudc_pag - (-C.t_difcampr))
   
   'Ventas Netas
   ' (quitado ) + Vales emitidos por devolución (no contar los vales emitidos por descuento cuando no se devuelve ninguna prenda)
   '+ Arreglos
   'tiene como resultado los cobros realizados
   'C.Cobros_Realizados = C.Ventas_Netas + C.t_vales_emi + C.T_Arreglos
   
   C.Cobros_Realizados = C.Ventas_Netas + C.t_vales_emi  '+ C.T_Arreglos
   
   
   Calcula_formulas_cierre = C


   On Error GoTo 0
   Exit Function

Calcula_formulas_cierre_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Calcula_formulas_cierre de Módulo Rutinas_Ventas"

End Function
  

'---------------------------------------------------------------------------------------
' Procedimiento : Distribuye_Caja_AB
' Fecha/Hora    : 13/02/2004 11:30
' Autor         : JCastillo
' Propósito     :  Distribuir % entra caja A y caja B de manera automatica
'                      Devuelve = Array con importes (0= CAJA A, 1= CAJA B)
'                               = 0, hubo algun error
'                               = 1, no es necesario distribir (caja A o B al 100%)
'---------------------------------------------------------------------------------------
Public Function Distribuye_Caja_AB(fcierre As String, C_CAJA As Byte, conexion As ADODB.Connection) As Variant
Dim m_matriz As Variant
Dim pA As Byte                  'porcentajes A y B %
Dim iA As Double
Dim pB As Byte
Dim iB As Double
Dim I_Total As Double        'importe total de la caja
Dim tmpvar As Variant
Dim tmpcont As Variant

Dim rc_ventas As New ADODB.Recordset

Dim Acum_Total As Double

m_matriz = devuelve_matriz("SELECT CAJA_A, CAJA_B from CAJAS WHERE CODIGO = " & C_CAJA, conexion)

'obtener porcentajes
pA = m_matriz(0)
pB = m_matriz(1)

'si se establece el 100% para la caja A, no redistribuir
If pA = 100 Then
    conexion.Execute "UPDATE CABVENTA SET CAJADES = 0 WHERE FHORA = '" & Format(fcierre, "yyyymmdd") & "' AND CODCAJA =" & C_CAJA
    Distribuye_Caja_AB = 1
    Exit Function
End If

'si se establece el 100% para la caja B, hacer un update a todas las
'ventas y salir
If pB = 100 Then
    conexion.Execute "UPDATE CABVENTA SET CAJADES = 1 WHERE FHORA = '" & Format(fcierre, "yyyymmdd") & "' AND CODCAJA =" & C_CAJA
    Distribuye_Caja_AB = 1
    Exit Function
End If

'ReDim M_Matriz(2)

   'total contado  (importe primera para formas de pago 1, 3 y 5) (ventas aceptadas)
   tmpcont = devuelve_campo("SELECT sum(IMP_PRIMERA) FROM CABVENTA WHERE FCOBRO IN (1, 3, 5) AND FHORA = '" & Format(fcierre, "yyyymmdd") & "' AND CODCAJA = " & C_CAJA & " AND ESTADO = 1", locCnn)
   If tmpcont = "@" Then tmpcont = 0
  
   'movimientos caja
   tmpvar = devuelve_campo("SELECT sum(IMPORTE) FROM MOVCAJA WHERE FCIERRE >= '" & Format(fcierre, "yyyymmdd") & "' AND FCIERRE < '" & Format(fcierre, "yyyymmdd") & "' AND CODCAJA = " & C_CAJA, locCnn)
   If tmpvar = "@" Then tmpvar = 0
     
   'obtener el total en caja
   I_Total = tmpvar + tmpcont

'pasar todas las ventas a A primero
conexion.Execute "UPDATE CABVENTA SET CAJADES = 0 WHERE FHORA = '" & Format(fcierre, "yyyymmdd") & "' AND CODCAJA =" & C_CAJA
'obtener total en caja
'Debug.Print devuelve_campo("SELECT T_ENCAJA FROM CIERREDIA WHERE FECIERRE = '" & Format(fcierre, "yyyymmdd") & "'", conexion)
'tmpvar = devuelve_campo("SELECT T_ENCAJA FROM CIERREDIA WHERE FECIERRE = '" & Format(fcierre, "yyyymmdd") & "'", conexion)

'If tmpvar <> "@" Then
'I_Total = tmpvar
'Else
'I_Total = 0
'End If

If I_Total = 0 Then
    Set rc_ventas = Nothing
    Distribuye_Caja_AB = 1  'no hay registros de cierre de caja, o el
                            'registro no es válido
    Exit Function
End If

'sacar importes
iA = (I_Total * pA) / 100
iB = (I_Total * pB) / 100


'ahora ajustar B hasta llegar mas o menos al importe
With rc_ventas

    .Open "SELECT * FROM CABVENTA WHERE FHORA = '" & Format(fcierre, "yyyymmdd") & "' AND FCOBRO NOT IN(2,3,6) AND CODCAJA =" & C_CAJA, conexion, adOpenDynamic, adLockOptimistic
       
    
    Do Until .EOF
    
            'ver si el importe sobrepasa al importe maximo para B
            If (.fields("IMP_PRIMERA") + .fields("IMP_SEGUNDA") + Acum_Total) <= iB Then
        
                'acumular el importe
                Acum_Total = (Acum_Total + .fields("IMP_PRIMERA") + .fields("IMP_SEGUNDA"))
        
                'pasar a B
                .fields("CAJADES") = 1
                .Update
                
                DoEvents
        
            End If
        
            If Not .EOF Then .MoveNext
           ' If Not .EOF Then .MoveNext
            
        Loop
    

    .Close
End With

'obtener importe CAJA A
m_matriz(0) = (I_Total - Acum_Total)

'obtener importe CAJA B
m_matriz(1) = Acum_Total

Distribuye_Caja_AB = m_matriz

ReDim m_matriz(0)

Set rc_ventas = Nothing

End Function




'---------------------------------------------------------------------------------------
' Subrutina   : cierra_caja
' Fecha/Hora  : 15/02/2004 01:50
' Autor       : JCASTILLO
' Propósito   : Inserta o actualiza el registro de cierre de caja diario.
'---------------------------------------------------------------------------------------
Public Sub cierra_caja(fcierre_ddmmyyyy As String, mi_Cierre As Datos_Cierre, conexion As ADODB.Connection)
Dim tmpvar As Variant
Dim strCon As String

   On Error GoTo cierra_caja_Error

'cepillarse las ventas que esten como PENDIENTES ...
'de detventa
conexion.Execute "DELETE FROM DETVENTA WHERE (CODCAJA = " & CajaActual & ") AND (CODVEN IN (SELECT CODIGO FROM CABVENTA WHERE ESTADO = 0 AND CODCAJA = " & CajaActual & "))"
'de cabventa
conexion.Execute "DELETE FROM CABVENTA WHERE ESTADO = 0 AND CODCAJA = " & CajaActual

'comprobar si ya existe un registro de cierre
tmpvar = devuelve_campo("SELECT CODIGO FROM CIERREDIA WHERE CODCAJA = " & CajaActual & " AND FECIERRE = '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "'", conexion)

'si no existe el registro, añadir nuevo
If tmpvar = "@" Then

    tmpvar = devuelve_campo("SELECT max(CODIGO) + 1  FROM CIERREDIA WHERE CODCAJA = " & CajaActual, conexion)
    If tmpvar = "@" Then tmpvar = 1
    
    strCon = "INSERT INTO CIERREDIA (CODIGO, CODCAJA, CODUSR, FECIERRE,T_CONTADO, N_CONTADO,  NVALES_ACEP, T_VALE_ACEP, NVALES_EMI, T_VALE_EMI, NVTAR, T_TARJETA, NDEVOL, T_DEVOLU, NARRE, T_ARREGLOS, T_ARRECON, N_MOVI, T_MOVI, T_ENCAJA, N_PAGOS, T_PAGOS, N_INGRESOS, T_INGRESOS, N_DEUDC, T_DEUDC, N_DEUDC_PAG, T_DEUDC_PAG, N_DESC, T_DESC, N_IVA, T_IVA, T_VEN_BRU, T_CAJAA, T_CAJAB, N_VALDCTOA, T_VALDCTOA, N_VALDCTOE, T_VALDCTOE, N_DIFCAMPR, T_DIFCAMPR) VALUES" & _
             " (" & tmpvar & ", " & CajaActual & ", " & UsuarioActual & ", '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "', " & _
             Replace(mi_Cierre.t_contado, ",", ".") & ", " & mi_Cierre.n_contado & ", " & mi_Cierre.n_vales_acep & ", " & _
             Replace(mi_Cierre.t_vales_acep, ",", ".") & ", " & _
             mi_Cierre.n_vales_emi & ", " & Replace(mi_Cierre.t_vales_emi, ",", ".") & ", " & _
             mi_Cierre.n_tarjeta & ", " & Replace(mi_Cierre.t_tarjeta, ",", ".") & ", " & mi_Cierre.n_devol & ", " & _
             Replace(mi_Cierre.t_devol, ",", ".") & ", " & mi_Cierre.n_arreglos & ", " & _
             Replace(mi_Cierre.T_Arreglos, ",", ".") & ", " & Replace(mi_Cierre.T_ArreCon, ",", ".") & ", " & _
             mi_Cierre.n_movi & ", " & Replace(mi_Cierre.t_movi, ",", ".") & ", " & Replace(mi_Cierre.t_caja, ",", ".") & ", " & _
             mi_Cierre.n_pagos & ", " & Replace(mi_Cierre.t_pagos, ",", ".") & ", " & _
             mi_Cierre.n_ingresos & ", " & Replace(mi_Cierre.t_ingresos, ",", ".") & ", " & mi_Cierre.n_deudc & ", " & _
             Replace(mi_Cierre.t_deudc, ",", ".") & ", " & mi_Cierre.n_deudc_pag & ", " & Replace(mi_Cierre.t_deudc_pag, ",", ".") & ", " & _
             mi_Cierre.n_dcto & ", " & _
             Replace(mi_Cierre.t_dcto, ",", ".") & ", " & _
             mi_Cierre.n_iva & ", " & _
             Replace(mi_Cierre.t_iva, ",", ".") & ", " & _
             Replace(mi_Cierre.Ventas_Brutas, ",", ".") & ", " & _
             Replace(mi_Cierre.Total_A, ",", ".") & ", " & Replace(mi_Cierre.Total_B, ",", ".") & ", " & _
             mi_Cierre.n_valdctoa & ", " & _
             Replace(mi_Cierre.t_valdctoa, ",", ".") & ", " & _
             mi_Cierre.n_valdctoe & ", " & _
             Replace(mi_Cierre.t_valdctoe, ",", ".") & ", " & _
             mi_Cierre.n_difcampr & ", " & _
             Replace(mi_Cierre.t_difcampr, ",", ".") & ") "

'si existe, actualizar el registro
Else
              ' "SET CODIGO = " & tmpvar & ", " & _
            ' "CODCAJA = " & CajaActual & ", " & _
   '(CODIGO, CODCAJA, CODUSR, FECIERRE, T_CONTADO, NVALES, T_VALE, NVTAR, T_TARJETA, NDEVOL, T_DEVOLU, NARRE, T_ARREGLOS, N_MOVI, T_MOVI, T_ENCAJA, N_PAGOS, T_PAGOS, N_DEUDC, T_DEUDCOBRA) VALUES"
    Debug.Print "TOTAL A:" & mi_Cierre.Total_A & "-TOTAL B:" & mi_Cierre.Total_B
   
    strCon = "UPDATE CIERREDIA SET CODUSR = " & UsuarioActual & ", FECIERRE = '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "', " & _
             "T_CONTADO = " & Replace(mi_Cierre.t_contado, ",", ".") & ", N_CONTADO = " & mi_Cierre.n_contado & ", " & _
             "NVALES_ACEP = " & mi_Cierre.n_vales_acep & ", T_VALE_ACEP = " & Replace(mi_Cierre.t_vales_acep, ",", ".") & ", " & _
             "NVALES_EMI = " & mi_Cierre.n_vales_emi & ", " & "T_VALE_EMI = " & Replace(mi_Cierre.t_vales_emi, ",", ".") & ", " & "NVTAR = " & mi_Cierre.n_tarjeta & ", " & "T_TARJETA = " & Replace(mi_Cierre.t_tarjeta, ",", ".") & ", " & _
             "NDEVOL = " & mi_Cierre.n_devol & ", " & "T_DEVOLU = " & Replace(mi_Cierre.t_devol, ",", ".") & ", " & _
             "NARRE = " & mi_Cierre.n_arreglos & ", " & _
             "T_ARREGLOS = " & Replace(mi_Cierre.T_Arreglos, ",", ".") & ", " & "T_ARRECON = " & Replace(mi_Cierre.T_ArreCon, ",", ".") & ", " & _
             "N_MOVI = " & mi_Cierre.n_movi & ", " & "T_MOVI = " & Replace(mi_Cierre.t_movi, ",", ".") & ", " & _
             "T_ENCAJA = " & Replace(mi_Cierre.t_caja, ",", ".") & ", " & "N_PAGOS = " & mi_Cierre.n_pagos & ", " & _
             "T_PAGOS = " & Replace(mi_Cierre.t_pagos, ",", ".") & ", " & "N_INGRESOS = " & mi_Cierre.n_ingresos & ", " & "T_INGRESOS = " & Replace(mi_Cierre.t_ingresos, ",", ".") & ", " & _
             "N_DEUDC = " & mi_Cierre.n_deudc & ", " & "T_DEUDC = " & Replace(mi_Cierre.t_deudc, ",", ".") & ", " & _
             "N_DEUDC_PAG = " & mi_Cierre.n_deudc_pag & ", " & "T_DEUDC_PAG = " & Replace(mi_Cierre.t_deudc_pag, ",", ".") & ", " & _
             "N_DESC = " & mi_Cierre.n_dcto & ", " & _
             "T_DESC = " & Replace(mi_Cierre.t_dcto, ",", ".") & ", " & _
             "N_IVA = " & mi_Cierre.n_iva & ", " & _
             "T_IVA = " & Replace(mi_Cierre.t_iva, ",", ".") & ", " & _
             "T_VEN_BRU = " & Replace(mi_Cierre.Ventas_Brutas, ",", ".") & ", " & _
             "T_CAJAA = " & Replace(mi_Cierre.Total_A, ",", ".") & ", " & "T_CAJAB = " & Replace(mi_Cierre.Total_B, ",", ".") & ", " & _
             "N_VALDCTOE = " & mi_Cierre.n_valdctoe & ", " & _
             "T_VALDCTOE = " & Replace(mi_Cierre.t_valdctoe, ",", ".") & ", " & _
             "N_VALDCTOA = " & mi_Cierre.n_valdctoa & ", " & _
             "T_VALDCTOA = " & Replace(mi_Cierre.t_valdctoa, ",", ".") & ", " & _
             "N_DIFCAMPR = " & mi_Cierre.n_difcampr & ", " & _
             "T_DIFCAMPR = " & Replace(mi_Cierre.t_difcampr, ",", ".") & " " & _
             "WHERE (CODIGO = " & tmpvar & ") AND (CODCAJA = " & CajaActual & ") AND (FECIERRE = '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "')"

End If

Debug.Print strCon

conexion.Execute strCon

DoEvents

'limpiar el .log temporal de sqlsqrver
conexion.Execute "BACKUP LOG LOCAL WITH TRUNCATE_ONLY"

DoEvents

   On Error GoTo 0
   Exit Sub

cierra_caja_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cierra_caja de Formulario frmCierreDia"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : ver_cierre_caja
' Fecha/Hora     : 18/06/2004 16:32
' Autor             : JCastillo
' Propósito       : Consulta la base de datos y obtiene un dia, o la suma de varios dias, devuelve un
'                       tipo datos_cierre con el resultado
'---------------------------------------------------------------------------------------
'
Public Function ver_cierre_caja(finicial As String, ffinal As String, Caja As Byte, conexion As ADODB.Connection) As Datos_Cierre
Dim rc As New ADODB.Recordset
Dim tmpw As String

   On Error GoTo ver_cierre_caja_Error
        
    If Trim(ffinal) <> "" Then
        tmpw = "(FECIERRE >= '" & Format(finicial, "yyyymmdd") & "') AND (FECIERRE <= '" & Format(ffinal, "yyyymmdd") & "')"
    Else
        tmpw = "(FECIERRE = '" & Format(finicial, "yyyymmdd") & "')"
    End If
        
    If Caja > 0 Then
        tmpw = tmpw & " AND (CODCAJA = " & Caja & ")"
    End If

    With ver_cierre_caja
    
        rc.Open "SELECT sum(t_contado) as t_contado, sum(n_contado) as n_contado, sum(NVALES_ACEP) as NVALES_ACEP," & _
        "sum(T_VALE_ACEP) as T_VALE_ACEP, sum(NVALES_EMI) as NVALES_EMI, sum(T_VALE_EMI) as T_VALE_EMI, sum(N_VALDCTOA) as N_VALDCTOA, sum(T_VALDCTOA) as T_VALDCTOA, sum(N_VALDCTOE) as N_VALDCTOE, sum(T_VALDCTOE) as T_VALDCTOE, sum(N_DIFCAMPR) as N_DIFCAMPR, sum(T_DIFCAMPR) as T_DIFCAMPR,  sum(NVTAR) as NVTAR, sum(T_TARJETA) as T_TARJETA," & _
        "sum(NDEVOL) as NDEVOL , sum(T_DEVOLU) as T_DEVOLU, sum(NARRE) as NARRE, sum(T_ARREGLOS) as T_ARREGLOS, sum(T_ARRECON) as T_ARRECON," & _
        "sum(N_MOVI) as N_MOVI, sum(T_MOVI ) as T_MOVI, sum(T_ENCAJA) as T_ENCAJA, sum(N_PAGOS) as N_PAGOS, sum(T_PAGOS) as T_PAGOS,sum(N_INGRESOS) as N_INGRESOS," & _
        "sum(T_INGRESOS) as T_INGRESOS, sum(N_DEUDC) as N_DEUDC, sum(T_DEUDC) as T_DEUDC, sum(N_DEUDC_PAG) as N_DEUDC_PAG, sum(T_DEUDC_PAG) as T_DEUDC_PAG," & _
        "sum(N_DESC) as N_DESC, sum(T_DESC) as T_DESC, sum(N_IVA) as N_IVA, sum(T_IVA) as T_IVA," & _
        "sum(T_VEN_BRU) as T_VEN_BRU, sum(T_CAJAA) as T_CAJAA, sum(T_CAJAB) as T_CAJAB FROM CIERREDIA where " & tmpw, conexion, adOpenStatic, adLockReadOnly
        
        
        If rc.RecordCount <= 0 Then
        
            'salida por error
            .Se_Ha_Calculado = False
            Exit Function
        
        End If
        
        If Not IsNull(rc.fields("T_CONTADO")) Then .t_contado = rc.fields("T_CONTADO")
        If Not IsNull(rc.fields("N_CONTADO")) Then .n_contado = rc.fields("N_CONTADO")
        If Not IsNull(rc.fields("NVALES_ACEP")) Then .n_vales_acep = rc.fields("NVALES_ACEP")
        If Not IsNull(rc.fields("T_VALE_ACEP")) Then .t_vales_acep = rc.fields("T_VALE_ACEP")
        If Not IsNull(rc.fields("NVALES_EMI")) Then .n_vales_emi = rc.fields("NVALES_EMI")
        If Not IsNull(rc.fields("T_VALE_EMI")) Then .t_vales_emi = rc.fields("T_VALE_EMI")
        
        If Not IsNull(rc.fields("N_VALDCTOE")) Then .n_valdctoe = rc.fields("N_VALDCTOE")
        If Not IsNull(rc.fields("T_VALDCTOE")) Then .t_valdctoe = rc.fields("T_VALDCTOE")
        
        If Not IsNull(rc.fields("N_VALDCTOA")) Then .n_valdctoa = rc.fields("N_VALDCTOA")
        If Not IsNull(rc.fields("T_VALDCTOA")) Then .t_valdctoa = rc.fields("T_VALDCTOA")
 
        If Not IsNull(rc.fields("N_VALDCTOA")) Then .n_valdctoa = rc.fields("N_VALDCTOA")
        If Not IsNull(rc.fields("T_VALDCTOA")) Then .t_valdctoa = rc.fields("T_VALDCTOA")
        
        If Not IsNull(rc.fields("NVTAR")) Then .n_tarjeta = rc.fields("NVTAR")
        If Not IsNull(rc.fields("T_TARJETA")) Then .t_tarjeta = rc.fields("T_TARJETA")
        If Not IsNull(rc.fields("NDEVOL")) Then .n_devol = rc.fields("NDEVOL")
        If Not IsNull(rc.fields("T_DEVOLU")) Then .t_devol = rc.fields("T_DEVOLU")
        If Not IsNull(rc.fields("NARRE")) Then .n_arreglos = rc.fields("NARRE")
        If Not IsNull(rc.fields("T_ARREGLOS")) Then .T_Arreglos = rc.fields("T_ARREGLOS")
        If Not IsNull(rc.fields("NARRE")) Then .n_arreglos = rc.fields("NARRE")
        If Not IsNull(rc.fields("T_ARRECON")) Then .T_ArreCon = rc.fields("T_ARRECON")
        If Not IsNull(rc.fields("N_MOVI")) Then .n_movi = rc.fields("N_MOVI")
        If Not IsNull(rc.fields("T_MOVI")) Then .t_movi = rc.fields("T_MOVI")
        If Not IsNull(rc.fields("T_ENCAJA")) Then .t_caja = rc.fields("T_ENCAJA")
        If Not IsNull(rc.fields("n_pagos")) Then .n_pagos = rc.fields("n_pagos")
        If Not IsNull(rc.fields("t_pagos")) Then .t_pagos = rc.fields("t_pagos")
        If Not IsNull(rc.fields("n_ingresos")) Then .n_ingresos = rc.fields("n_ingresos")
        If Not IsNull(rc.fields("t_ingresos")) Then .t_ingresos = rc.fields("t_ingresos")
        If Not IsNull(rc.fields("n_deudc")) Then .n_deudc = rc.fields("n_deudc")
        If Not IsNull(rc.fields("t_deudc")) Then .t_deudc = rc.fields("t_deudc")
        If Not IsNull(rc.fields("n_deudc_pag")) Then .n_deudc_pag = rc.fields("n_deudc_pag")
        If Not IsNull(rc.fields("t_deudc_pag")) Then .t_deudc_pag = rc.fields("t_deudc_pag")
        If Not IsNull(rc.fields("N_DESC")) Then .n_dcto = rc.fields("N_DESC")
        If Not IsNull(rc.fields("T_DESC")) Then .t_dcto = rc.fields("T_DESC")
        
        If Not IsNull(rc.fields("N_difcampr")) Then .n_difcampr = rc.fields("n_difcampr")
        If Not IsNull(rc.fields("T_difcampr")) Then .t_difcampr = rc.fields("t_difcampr")

        If Not IsNull(rc.fields("n_iva")) Then .n_iva = rc.fields("n_iva")
        If Not IsNull(rc.fields("t_iva")) Then .n_iva = rc.fields("t_iva")
        If Not IsNull(rc.fields("T_VEN_BRU")) Then .Ventas_Brutas = rc.fields("T_VEN_BRU")
        If Not IsNull(rc.fields("T_CAJAA")) Then .Total_A = rc.fields("T_CAJAA")
        If Not IsNull(rc.fields("T_CAJAB")) Then .Total_B = rc.fields("T_CAJAB")
        
        .Se_Ha_Calculado = True
                
    End With
    
    ver_cierre_caja = Calcula_formulas_cierre(ver_cierre_caja, True)
    
    
    rc.Close
    Set rc = Nothing
    
    
    'strCon = "UPDATE CIERREDIA SET CODUSR = " & UsuarioActual & ", FECIERRE = '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "', " & _
             "T_CONTADO = " & Replace(mi_Cierre.t_contado, ",", ".") & ", N_CONTADO = " & mi_Cierre.n_contado & ", " & _
             "NVALES_ACEP = " & mi_Cierre.n_vales_acep & ", T_VALE_ACEP = " & Replace(mi_Cierre.t_vales_acep, ",", ".") & ", " & _
             "NVALES_EMI = " & mi_Cierre.n_vales_emi & ", " & "T_VALE_EMI = " & Replace(mi_Cierre.t_vales_emi, ",", ".") & ", " & "NVTAR = " & mi_Cierre.n_tarjeta & ", " & "T_TARJETA = " & Replace(mi_Cierre.t_tarjeta, ",", ".") & ", " & _
             "NDEVOL = " & mi_Cierre.n_devol & ", " & "T_DEVOLU = " & Replace(mi_Cierre.t_devol, ",", ".") & ", " & _
             "NARRE = " & mi_Cierre.n_arreglos & ", " & _
             "T_ARREGLOS = " & Replace(mi_Cierre.T_Arreglos, ",", ".") & ", " & "T_ARRECON = " & Replace(mi_Cierre.T_ArreCon, ",", ".") & ", " & _
             "N_MOVI = " & mi_Cierre.n_movi & ", " & _
             "T_MOVI = " & Replace(mi_Cierre.t_movi, ",", ".") & ", " & _
             "T_ENCAJA = " & Replace(mi_Cierre.t_caja, ",", ".") & ", " & _
             "N_PAGOS = " & mi_Cierre.n_pagos & ", " & _
             "T_PAGOS = " & Replace(mi_Cierre.t_pagos, ",", ".") & ", " & _
             "N_INGRESOS = " & mi_Cierre.n_ingresos & ", " & _
             "T_INGRESOS = " & Replace(mi_Cierre.t_ingresos, ",", ".") & ", " & _
             "N_DEUDC = " & mi_Cierre.n_deudc & ", " & _
             "T_DEUDC = " & Replace(mi_Cierre.t_deudc, ",", ".") & ", " & _
             "N_DEUDC_PAG = " & mi_Cierre.n_deudc_pag & ", " & _
             "T_DEUDC_PAG = " & Replace(mi_Cierre.t_deudc_pag, ",", ".") & ", " & _
             "N_DESC = " & mi_Cierre.n_dcto & ", " & _
             "T_DESC = " & Replace(mi_Cierre.t_dcto, ",", ".") & ", " & _
             "N_IVA = " & mi_Cierre.n_iva & ", " & _
             "T_IVA = " & Replace(mi_Cierre.t_iva, ",", ".") & ", " & _
             "T_VEN_BRU = " & Replace(mi_Cierre.Ventas_Brutas, ",", ".") & ", " & _
             "T_CAJAA = " & Replace(mi_Cierre.Total_A, ",", ".") & ", " & "T_CAJAB = " & Replace(mi_Cierre.Total_B, ",", ".") & " " & _
             "WHERE (CODIGO = " & tmpvar & ") AND (CODCAJA =" & CajaActual & ") AND (FECIERRE = '" & Format(fcierre_ddmmyyyy, "yyyymmdd") & "')"
          

   On Error GoTo 0
   Exit Function

ver_cierre_caja_Error:

    ver_cierre_caja.Se_Ha_Calculado = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ver_cierre_caja de Módulo Rutinas_Ventas"
    
End Function


'---------------------------------------------------------------------------------------
' Subrutina   : Imprime_Ticket
' Fecha/Hora  : 15/02/2004 21:13
' Autor       : JCASTILLO
' Propósito   : Imprime el Ticket de Venta especificado, por 3 recordsets
'---------------------------------------------------------------------------------------
Public Sub Imprime_Ticket(cambio As Double, rcCabVen As ADODB.Recordset, rcDetVen As ADODB.Recordset, rcarreglos As ADODB.Recordset)
Const LenLinea = 39
Dim tmplinea As String
Dim tmpmodel As String * 12
Dim tmpprov As String * 3
Dim t_articulo As Variant
'Dim tmptalla As String * 3
'Dim tmpcolor As String * 4
Dim tmppreven As String * 7
Dim tmpuds As String * 3
Dim tmptotal As String * 7
Dim linea As String * LenLinea

'---------------------------------------------------
'
'ESTRUCTURA Total: 34 Posiciones
'
'    MODELO   Prov   PVP      UDS    TOTAL
'         12    3     7        3         7
'XXXXXXXXXXX        0000,00   000   0000,00
'
'3 blancos
'---------------------------------------------------

Dim tmppvp As Double
Dim totcompra As Double

   'On Error GoTo Imprime_Ticket_Error

Printer.Font.Name = "Courier New"

'imprimir cabecera
Printer.Print " "
Printer.Print " "
If Trim(LineasTicket(0)) <> "" Then Printer.Print LineasTicket(0)
If Trim(LineasTicket(1)) <> "" Then Printer.Print LineasTicket(1)
If Trim(LineasTicket(2)) <> "" Then Printer.Print LineasTicket(2)
Printer.Print " "

'Imprimir código y fecha/hora
Printer.Print "Ticket: " & rcCabVen.fields("CODIGO") & Format(rcCabVen.fields("CODCAJA"), "000")
Printer.Print " "

'solo si hay ventas
If rcDetVen.RecordCount > 0 Then

'Imprimir cabeceras para el detalle
Printer.Print "   Articulo           PVP   UDS  TOTAL"
Printer.Print " "

rcDetVen.MoveFirst

'IMPRIMIR ARTÍCULOS DE LA VENTA
Do Until rcDetVen.EOF


t_articulo = devuelve_matriz("SELECT MODELO, CODPROV FROM MAARTIC WHERE CODIGO = " & rcDetVen.fields("CODART") & " AND TEMPOR = " & rcDetVen.fields("TEMPOR"), locCnn)
tmpprov = UCase(Left(Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(1), locCnn)), 3))
'modelo de 10 + 1 espacio
tmpmodel = UCase(Left(Trim(t_articulo(0)), 12))
tmplinea = tmplinea & tmpmodel & "-"
tmplinea = tmplinea & tmpprov & "     "
'talla 3 + 1 espacio
'tmptalla = UCase(Left(Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcDetVen.fields("CODTALLA"), locCnn)), 3))
'tmplinea = tmplinea & tmptalla & "-"
'color 6 + 1 espacio
'tmpcolor = UCase(Left(Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcDetVen.Fields("CODCOL"), locCnn)), 6))
'tmplinea = tmplinea & tmpcolor & " "

'preven 6 + 1 espacio
tmppreven = Left(Format(rcDetVen.fields("PREVEN"), "00.00"), 7)
tmplinea = tmplinea & tmppreven & " "

'UDS 3 + 1 espacio
tmpuds = Left(CStr(rcDetVen.fields("UNIDADES")), 3) & " "
tmplinea = tmplinea & tmpuds & " "

'TOTAL 8 + 1 espacio

'pvp menos dcto
tmppvp = rcDetVen.fields("PREVEN") - ((rcDetVen.fields("PREVEN") * rcDetVen.fields("DCTO")) / 100)
'precio * unidades
tmppvp = tmppvp * rcDetVen.fields("UNIDADES")
'total + IVA
tmppvp = tmppvp + ((tmppvp * rcDetVen.fields("IVA")) / 100)

totcompra = totcompra + tmppvp

tmptotal = Left(Format(tmppvp, "00.00"), 7)

tmplinea = tmplinea & tmptotal

'imprimir linea
Printer.Print tmplinea

tmplinea = ""
tmpmodel = ""
'tmptalla = ""
tmpprov = ""
'tmpcolor = ""
tmppreven = ""
tmpuds = ""
tmptotal = ""

'conformar código de barras del artículo (13 digitos)
tmplinea = Format(rcDetVen.fields("CODART"), "00000") & Format(rcDetVen.fields("TEMPOR"), "000") & Format(rcDetVen.fields("CODTALLA"), "00") & Format(rcDetVen.fields("CODCOL"), "000")


'si esta en oferta, mirar primero si es una oferta con Dcto, en ese caso imprimir los 2
'mensajes (oferta y dcto)
If rcDetVen.fields("OFERTA") = True Then
    
    'imprimir 2 mensajes
    If rcDetVen.fields("DCTO") > 0 Then
        tmplinea = tmplinea & " *Oferta* (dcto del " & rcDetVen.fields("DCTO") & "%)"
    'imprimir solo el mensaje de oferta
    Else
       tmplinea = tmplinea & " *Oferta*"
    End If
    
    
ElseIf rcDetVen.fields("DCTO") > 0 Then
'si tenia descuento imprimir tb información de descuento
    tmplinea = tmplinea & "   (dcto del " & rcDetVen.fields("DCTO") & " %)"
End If




Printer.Print tmplinea
tmplinea = ""

    If Not rcDetVen.EOF Then rcDetVen.MoveNext

Loop

'si hay ventas (fin)
End If


'si hay arreglos ...
If rcarreglos.RecordCount > 0 Then

    Printer.Print " "
    Printer.Print "ARREGLOS:"
    Printer.Print " "
    
    'IMPRIMIR ARREGLOS DE LA VENTA
    rcarreglos.MoveFirst
    Do Until rcarreglos.EOF
                
        'modelo de 10 + 1 espacio
        
        If Not IsNull(rcarreglos.fields("CODART")) Then
            tmpmodel = UCase(Left(Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcarreglos.fields("CODART") & " AND TEMPOR = " & rcarreglos.fields("TEMPOR"), locCnn)), 12))
            Else
            tmpmodel = "Arreglos Varios"
        End If
        
        'tmpmodel = UCase(Left(Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcArreglos.fields("CODART") & " AND TEMPOR = " & rcArreglos.fields("TEMPOR"), locCnn)), 12))
        tmplinea = tmplinea & tmpmodel & "     "

        'talla 3 + 1 espacio
       ' tmptalla = UCase(Left(Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcArreglos.fields("CODTALLA"), locCnn)), 3))
       ' tmplinea = tmplinea & tmptalla & "-"
        'color 6 + 1 espacio
        'tmpcolor = UCase(Left(Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcArreglos.Fields("CODCOL"), locCnn)), 6))
        'tmplinea = tmplinea & tmpcolor & " "
        
        'PREVEN (blancos)
        tmplinea = tmplinea & "      " & " "
        
        'UDS (blancos)
        tmpuds = "   "
        tmplinea = tmplinea & tmpuds & "      "
         
        totcompra = totcompra + rcarreglos.fields("PVP")
        
        tmptotal = Left(CStr(Format(rcarreglos.fields("PVP"), "00.00")), 8)
        tmplinea = tmplinea & tmptotal
        
        Printer.Print tmplinea
        tmplinea = ""

        If Not rcarreglos.EOF Then rcarreglos.MoveNext
    Loop
    
End If

Printer.Print " "
linea = ""

'insertar al final

Printer.Print Justifica_Cadena("Total: " & Left(Format(totcompra, "00.00"), 10), LenLinea - 1, False)
'Mid(linea, LenLinea - Len("Total: " & Left(Format(totcompra, "00.00"), 10)), Len("Total: " & Left(Format(totcompra, "Currency"), 10))) = "Total: " & Left(Format(totcompra, "00.00"), 10)

DoEvents

Printer.Print linea

linea = ""

Printer.Print "Forma Pago: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM FCOBRO WHERE CODIGO = " & rcCabVen.fields("FCOBRO"), locCnn))

Select Case rcCabVen.fields("FCOBRO")

'contado
Case 1
    tmplinea = "CONTADO: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00")
'tarjeta
Case 2
    tmplinea = "TARJETA: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00")
'contado y tarjeta
Case 3
    tmplinea = "CONTADO: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00") & " - TARJETA: " & Format(rcCabVen.fields("IMP_SEGUNDA"), "00.00")
'vale
Case 4
    tmplinea = "VALE: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00")
'contado y vale
Case 5
    tmplinea = "CONTADO: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00") & " - VALE: " & Format(rcCabVen.fields("IMP_SEGUNDA"), "00.00")
'tarjeta y vale
Case 6
    tmplinea = "TARJETA: " & Format(rcCabVen.fields("IMP_PRIMERA"), "00.00") & " - VALE: " & Format(rcCabVen.fields("IMP_SEGUNDA"), "00.00")

End Select

'Mid(linea, LenLinea - Len(tmplinea), Len(tmplinea)) = tmplinea

Printer.Print Justifica_Cadena(tmplinea, LenLinea - 1, False)

'si tiene cambio, imprimir la linea del cambio
If cambio > 0 Then
    tmplinea = ""
    tmplinea = "CAMBIO: " & Format(cambio, "00.00")
    Printer.Print Justifica_Cadena(tmplinea, LenLinea - 1, False)
End If

Printer.Print "Importes expresados en euros"
'imprimir importes ...
'Printer.Print linea

Printer.Print ""
Printer.Print "Le atendio: " & UCase(Left(Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rcCabVen.fields("CODPER"), locCnn)), 28))
Printer.Print "Fecha/Hora: " & Now
Printer.Print " "
Printer.Print " "
If Trim(LineasTicket(3)) <> "" Then Printer.Print LineasTicket(3)
If Trim(LineasTicket(4)) <> "" Then Printer.Print LineasTicket(4)
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "

Printer.EndDoc


   On Error GoTo 0
   Exit Sub

Imprime_Ticket_Error:

    Printer.EndDoc
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Imprime_Ticket de Módulo Rutinas_Ventas"

End Sub


'justifica una cadena de caracteres a la izquierda o derecha hasta llegar al la Longitud Total. Para
'justificar lineas con Printer.Print
Public Function Justifica_Cadena(sOriginal As String, LongitudTotal As Long, a_la_Izquierda As Boolean) As String

'si es mayor no hacer nada y salir
If Len(sOriginal) > LongitudTotal Then
    Justifica_Cadena = ""
    Exit Function
End If

'justificar a la IZQUIERDA
If a_la_Izquierda Then
Justifica_Cadena = sOriginal & String(LongitudTotal - Len(sOriginal), " ")
'justificar a la DERECHA
Else
Justifica_Cadena = String(LongitudTotal - Len(sOriginal), " ") & sOriginal
End If

End Function


'---------------------------------------------------------------------------------------
' Procedimiento : Deshacer_Venta
' Fecha/Hora     : 17/02/2004 12:36
' Autor             : JCastillo
' Propósito       :  Deshace la venta especificada (devuelve las unidades al almacén y pone la
'                        venta como estado 0.
'---------------------------------------------------------------------------------------
Public Function Deshacer_Venta(codven As Long, codcaja As Byte, conexion As ADODB.Connection) As Boolean
Dim rc As ADODB.Recordset
Dim tmpestado As Variant

   On Error GoTo Deshacer_Venta_Error

tmpestado = devuelve_campo("SELECT ESTADO FROM CABVENTA WHERE CODIGO = " & codven & " AND CODCAJA = " & codcaja, locCnn)
Set rc = New ADODB.Recordset

'solo si ya ha sido pasada
If tmpestado > 0 Then

    'obtener registros de la venta ...
    rc.Open "SELECT * FROM DETVENTA WHERE CODVEN = " & codven & " AND CODCAJA =" & codcaja, conexion, adOpenDynamic, adLockReadOnly

    Do Until rc.EOF

        'introducir unidades otra vez en stock
        Call stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, rc.fields("UNIDADES"), True, conexion)
        rc.MoveNext
    
    Loop

    rc.Close

End If

'pasar la venta a pendiente ...
conexion.Execute "UPDATE CABVENTA SET ESTADO = 0 WHERE CODIGO = " & codven & " AND CODCAJA = " & codcaja

'poner los arreglos como estado = 1 si hay de esa venta ...
conexion.Execute "UPDATE ARREGLOS SET ESTADO = 1 WHERE CODVEN = " & codven & " AND CODCAJ = " & codcaja


If rc.State = 1 Then rc.Close
'abrir todas las devoluciones para la venta
rc.Open "SELECT * FROM DEVOL WHERE CODVEN = " & codven & " AND CODCAJA = " & codcaja & " ORDER BY CODIGO", conexion, adOpenDynamic, adLockOptimistic
        
If Not (rc.EOF And rc.BOF) Then
    
    Do Until rc.EOF
        'sacar la unidad devuelta del stock
        stock rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, rc.fields("UNIDADES"), False, conexion
        rc.MoveNext
        '
    Loop
    
End If
        
If rc.State = 1 Then rc.Close
Set rc = Nothing

   On Error GoTo 0
   Exit Function

Deshacer_Venta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Deshacer_Venta de Módulo Rutinas_Ventas"

End Function


Private Function recodifica_devoluciones(codigo_venta As Long, nuevo_codigo_venta, codigo_caja As Byte, conexion As ADODB.Connection) As Byte
Dim rc As ADODB.Recordset
Dim tmpvar As Variant
Dim tmpvales As Variant

   On Error GoTo recodifica_devoluciones_Error
   
   

        'abrir todos los arreglos para esa venta ...
        rc.Open "SELECT * FROM DEVOL WHERE CODVEN = " & codigo_venta & " AND CODCAJA = " & codigo_caja & " ORDER BY CODIGO", conexion, adOpenDynamic, adLockOptimistic
        
        If rc.EOF And rc.BOF Then
            recodifica_devoluciones = 0
            Exit Function
        End If
        
        'sacar el ultimo id de arreglos para la caja actual
        tmpvar = devuelve_campo("SELECT MAX(CODIGO) FROM DEVOL WHERE CODCAJA = " & CajaActual, conexion)
    
        'incrementar código
        If tmpvar = "@" Then
            tmpvar = 1
        Else
            If IsNumeric(tmpvar) Then
                tmpvar = tmpvar + 1
            Else
                tmpvar = 1
            End If
        End If
        
        'sacar el ultimo id de arreglos para la caja actual
        tmpvales = devuelve_campo("SELECT MAX(CODIGO) FROM VALES WHERE CODCAJA = " & CajaActual, conexion)
    
        Do Until rc.EOF
            
            'ACTUALIZAR EL VALE SI TIENE LA DEVOLUCION
            '(cambiar codigo, codcaja y poner el nuevo codigo de venta)
            conexion.Execute "UPDATE VALES SET CODVEN = " & nuevo_codigo_venta & ", CODIGO = " & tmpvales & ", CODCAJA = " & CajaActual & " WHERE CODVEN =" & codigo_venta & " AND CODCAJA =" & codigo_caja
            
            rc.fields("ID") = tmpvar
            rc.fields("CODCAJ") = CajaActual
            rc.fields("CODVEN") = nuevo_codigo_venta
            rc.Update
            DoEvents
            
            rc.MoveNext
            tmpvar = tmpvar + 1
            tmpvales = tmpvales + 1
            
                      
        Loop
        
        
   rc.Close
   Set rc = Nothing

   On Error GoTo 0
   Exit Function

recodifica_devoluciones_Error:

    recodifica_devoluciones = 1
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento recodifica_devoluciones de Formulario frmTrabaDatos"
End Function


'---------------------------------------------------------------------------------------
' Subrutina   : inserta_saldo_inicial
' Fecha/Hora  : 26/02/2004 21:17
' Autor       : JCASTILLO
' Propósito   : Inserta el saldo inicial en la caja. Recibe YYYYMMDD sin las //
'               Devuelve FALSE -> no hubo ningun error
'                        TRUE -> salida por error
'---------------------------------------------------------------------------------------
Public Function inserta_saldo_inicial(fcierre As Date, conexion As ADODB.Connection) As Boolean
Dim fcierreYYYYMMDD As String
Dim tmpcount As Variant
Dim tmpmax As Variant
Dim tmpsal As String
Dim saldo As Double

   On Error GoTo inserta_saldo_inicial_Error

    fcierreYYYYMMDD = Format(fcierre, "yyyymmdd")
    
    tmpcount = devuelve_campo("SELECT COUNT(CODIGO) FROM MOVCAJA WHERE FCIERRE = '" & fcierreYYYYMMDD & "' AND CODCAJA =" & CajaActual, conexion)
    tmpmax = devuelve_campo("SELECT MAX(CODIGO) + 1 FROM MOVCAJA WHERE CODCAJA =" & CajaActual, conexion)
    
    If tmpcount = "@" Then tmpcount = 0
    If tmpmax = "@" Then tmpmax = 1
    
    'si hay registros salir
    If tmpcount > 0 Then
        inserta_saldo_inicial = True
        Exit Function
    End If
    
    saldo = devuelve_campo("SELECT SALDOINI FROM CAJAS WHERE CODIGO = " & CajaActual, conexion)
    tmpsal = InputBox("Introduzca el saldo inicial de la caja:", "Abrir Caja", saldo)
    
    'validaciones ...
    If Trim(tmpsal) = "" Or Not IsNumeric(tmpsal) Then
        inserta_saldo_inicial = True
        Exit Function
    End If
    
    saldo = CDbl(Replace(tmpsal, ".", ","))
       
    'insertar el registro en movcaja
    conexion.Execute "INSERT INTO MOVCAJA (CODIGO, LINEA, CODCAJA, CODUSR, IMPORTE, COMENTARIO, FCIERRE) " & _
                     "VALUES (" & tmpmax & ",0 , " & CajaActual & ", " & UsuarioActual & ", " & Replace(CStr(saldo), ",", ".") & ",'SALDO INICIAL DE CAJA', '" & fcierreYYYYMMDD & "')"
    
    ReDim tmpvar(0)

   On Error GoTo 0
   Exit Function

inserta_saldo_inicial_Error:

    inserta_saldo_inicial = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento inserta_saldo_inicial de Módulo Rutinas_Ventas"
End Function



'---------------------------------------------------------------------------------------
' Subrutina   : inserta_deuda_cliente
' Fecha/Hora  : 21/04/2004 20:52
' Autor       : JCASTILLO
' Propósito   : Inserta una deuda de un cliente en la tabla de cabecera de deudas.
'               devuelve el código de la deuda creada.
'---------------------------------------------------------------------------------------
Public Function inserta_deuda_cliente(CodigoCliente As Long, CajaCliente As Byte, CodigoVenta As Long, ImporteDeuda As Double, Comentario As String, conexion As ADODB.Connection) As Long

   On Error GoTo inserta_deuda_cliente_Error
   
   Dim tmpvar As Variant
   
   'obtener el ultimo código
   tmpvar = devuelve_campo("SELECT MAX(CODIGO) + 1 FROM CABDEUDCLI WHERE CODCAJA = " & CajaActual, conexion)
            
   If tmpvar = "@" Then tmpvar = 1
            
    conexion.Execute "INSERT INTO CABDEUDCLI (CODIGO, CODCAJA, CODCLI, CAJACLI, CODVEN, COMEN, IMPORTE) VALUES" & _
                     " (" & tmpvar & ", " & CajaActual & ", " & CodigoCliente & ", " & CajaCliente & ", " & CodigoVenta & ", '" & Comentario & "', " & Replace(CStr(ImporteDeuda), ",", ".") & ")"

   inserta_deuda_cliente = tmpvar

   On Error GoTo 0
   Exit Function

inserta_deuda_cliente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento inserta_deuda_cliente de Módulo Rutinas_Ventas"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : inserta_pago_cliente
' Fecha/Hora     : 14/05/2004 09:57
' Autor             : JCastillo
' Propósito       :  Inserta una pago de un cliente, sobre una deuda, y recalcula el resto de la deuda q
'                        y actualiza el registro de cabecera con ese dato y el estado del pago.
'                        Devuelve TRUE si se ha creado correctamente, y FALSE si no se ha creado
'---------------------------------------------------------------------------------------'
Public Function inserta_pago_cliente(codigo As Long, codcaja As Byte, Usuario As Integer, Importe As Double, factura As Long, Comentario As String, conexion As ADODB.Connection) As Boolean
Dim tmpimp As Variant
Dim totimp As Double
Dim tmpestado As Long

   On Error GoTo inserta_pago_cliente_Error

        conexion.Execute "INSERT INTO DETDEUDCLI (CODIGO, CODCAJA, CODPER, IMPORTE, FACTURA, DESCRIPCION) VALUES (" & codigo & "," & codcaja & "," & Usuario & "," & Replace(Importe, ",", ".") & "," & factura & ",'" & Comentario & "')"
        
        DoEvents
        
        tmpimp = devuelve_campo("SELECT SUM(IMPORTE) FROM DETDEUDCLI WHERE CODIGO = " & codigo & " AND CODCAJA = " & codcaja, conexion)
        If tmpimp = "@" Then tmpimp = 0
        
        DoEvents
        
        totimp = devuelve_campo("SELECT IMPORTE FROM CABDEUDCLI WHERE CODIGO = " & codigo & " AND CODCAJA = " & codcaja, conexion)
        
        DoEvents
        
        'si es mayor o igual al importe total de la deuda, poner como pagada
        If (tmpimp >= totimp) Then
            tmpestado = 2
        'si es menor, poner como parcial
        ElseIf (tmpimp < totimp) And (tmpimp > 0) Then
            tmpestado = 1
        End If
        
        conexion.Execute "UPDATE CABDEUDCLI SET PAGADO = " & Replace(tmpimp, ",", ".") & ", ESTADO = " & tmpestado & " WHERE CODIGO = " & codigo & " AND CODCAJA = " & codcaja
        
        inserta_pago_cliente = True
    
   On Error GoTo 0
   Exit Function

inserta_pago_cliente_Error:

    inserta_pago_cliente = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento inserta_pago_cliente de Módulo Rutinas_Ventas"
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : añade_arreglo_rapido
' Fecha/Hora     : 28/04/2004 16:25
' Autor             : JCastillo
' Propósito       : Añade un arreglo rápido en la base de datos de arreglos. Devuelve el codigo del arreglo
'                       creado, 0 si se produce algun error
'---------------------------------------------------------------------------------------
Public Function añade_arreglo_rapido(codven As Long, codcaja As Byte, coste As Currency, pvp As Currency, descripcion As String, codusr As Long, conexion As ADODB.Connection)

    Dim tmpvar As Variant
    On Error GoTo añade_arreglo_rapido_Error

    tmpvar = devuelve_campo("SELECT MAX(ID) + 1 FROM ARREGLOS WHERE CODCAJ = " & codcaja, conexion)
    
    If tmpvar = "@" Then tmpvar = 1
    
    conexion.Execute "INSERT INTO ARREGLOS (ID, CODVEN, CODCAJ, COSTE, PVP, DESCRIPCION, CODUSR, ESTADO) VALUES " & _
    "(" & tmpvar & ", " & codven & ", " & codcaja & ", " & Replace(coste, ",", ".") & ", " & Replace(pvp, ",", ".") & ", '" & descripcion & "', " & codusr & ", 1)"
    
    añade_arreglo_rapido = tmpvar

   On Error GoTo 0
   Exit Function

añade_arreglo_rapido_Error:
    
    añade_arreglo_rapido = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añade_arreglo_rapido de Módulo Rutinas_Ventas"
End Function


'---------------------------------------------------------------------------------------
' Subrutina   : Conforma_CB
' Fecha/Hora  : 16/05/2004 20:40
' Autor       : JCASTILLO
' Propósito   : Conforma un string con el codigo de barras recibiendo los parametros
'
'---------------------------------------------------------------------------------------
Public Function Conforma_CB(codart As Long, tempor As Integer, codtalla As Byte, codcol As Integer) As String

    Conforma_CB = Format(codart, "00000") & Format(tempor, "000") & Format(codtalla, "00") & Format(codcol, "000")

End Function





