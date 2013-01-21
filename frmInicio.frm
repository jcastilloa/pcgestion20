VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A71B311E-9427-431E-AC22-204298AB866C}#1.0#0"; "miBwordOCX.ocx"
Begin VB.MDIForm frmInicio 
   BackColor       =   &H00EEA78E&
   Caption         =   "MDIForm1"
   ClientHeight    =   7890
   ClientLeft      =   1425
   ClientTop       =   1485
   ClientWidth     =   11505
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   1515
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin miBwordOCX.miBword Editor 
      Left            =   15
      Top             =   15
      _ExtentX        =   2646
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mantenimientos 
      Caption         =   "Mantenimientos"
      Begin VB.Menu MNT_Compras 
         Caption         =   "Compras"
         Begin VB.Menu Mnt_Articulos 
            Caption         =   "Artículos"
         End
         Begin VB.Menu Mnt_Proveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu Mnt_Tallas 
            Caption         =   "Tallas"
         End
         Begin VB.Menu Mnt_Colores 
            Caption         =   "Colores"
         End
         Begin VB.Menu Mnt_Secciones 
            Caption         =   "Secciones"
         End
         Begin VB.Menu Mnt_Familias 
            Caption         =   "Familias"
         End
         Begin VB.Menu Mnt_SubFamilias 
            Caption         =   "SubFamilias"
         End
         Begin VB.Menu Mnt_FormasPago 
            Caption         =   "Formas de Pago"
         End
         Begin VB.Menu Mnt_Sectores 
            Caption         =   "Sectores (Prov)"
         End
         Begin VB.Menu Mnt_CatTallas 
            Caption         =   "Categorías de Tallas"
         End
         Begin VB.Menu Mnt_Temporadas 
            Caption         =   "Temporadas"
         End
      End
      Begin VB.Menu Mnt_Ventas 
         Caption         =   "Ventas"
         Begin VB.Menu Mnt_Clientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu Mnt_Costureras 
            Caption         =   "Costureras"
         End
         Begin VB.Menu Mnt_FormasCobro 
            Caption         =   "Formas de Cobro"
         End
         Begin VB.Menu IVA 
            Caption         =   "IVA"
         End
         Begin VB.Menu Mnt_Ofertas 
            Caption         =   "Ofertas"
         End
      End
      Begin VB.Menu Mnt_Varios 
         Caption         =   "Varios"
         Begin VB.Menu Mnt_Personal 
            Caption         =   "Personal"
         End
         Begin VB.Menu Mnt_Centros 
            Caption         =   "Centros"
         End
         Begin VB.Menu Mnt_Cajas 
            Caption         =   "Cajas"
         End
         Begin VB.Menu Mnt_Almacenes 
            Caption         =   "Almacenes"
         End
         Begin VB.Menu Mnt_Bancos 
            Caption         =   "Bancos"
         End
         Begin VB.Menu Mnt_Tarjetas 
            Caption         =   "Tarjetas de Crédito"
         End
      End
   End
   Begin VB.Menu dep_ConsultaStock 
      Caption         =   "Consulta de Stock"
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Begin VB.Menu Compras_RecibirMercancia 
         Caption         =   "Recibir Mercancía"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Ver_Existencias 
         Caption         =   "Ver Existencias"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Consulta_Pedidos 
         Caption         =   "Consulta de Pedidos"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Consulta_Pedidos_Totales 
         Caption         =   "Consulta de Pedidos Totales"
      End
      Begin VB.Menu Pagos_Proveedores 
         Caption         =   "Pagos a Proveedores"
         Begin VB.Menu Introducir_Pagos 
            Caption         =   "Introducir Pagos"
         End
         Begin VB.Menu Ver_Pagos 
            Caption         =   "Ver Pagos"
         End
      End
      Begin VB.Menu ComVarios 
         Caption         =   "Varios"
         Begin VB.Menu Crear_Etiquetas 
            Caption         =   "Crear Etiquetas"
         End
         Begin VB.Menu Actualizar_Precios_Compra 
            Caption         =   "Actualizar Precios Compra"
         End
         Begin VB.Menu Actualizar_Iva 
            Caption         =   "Actualizar Iva Compra"
         End
         Begin VB.Menu Ajusta_Numero_Pedido 
            Caption         =   "Ajusta Número Pedido"
         End
      End
   End
   Begin VB.Menu Transferencias 
      Caption         =   "Transferencias"
      Begin VB.Menu PTrans_Crear 
         Caption         =   "Crear Transferencia"
         Shortcut        =   {F6}
      End
      Begin VB.Menu PTrans_Ver 
         Caption         =   "Ver Transferencias"
         Begin VB.Menu Ver_Trn_Pendientes 
            Caption         =   "Ver Pendientes"
         End
         Begin VB.Menu Consulta_Detallada_Trn 
            Caption         =   "Consulta Detallada"
         End
         Begin VB.Menu Consulta_Trn_Totales 
            Caption         =   "Consulta de Totales"
         End
      End
      Begin VB.Menu Recibir_Transferencia_Fichero 
         Caption         =   "Incorporar Transferencia desde fichero"
      End
   End
   Begin VB.Menu Ventas 
      Caption         =   "Ventas"
      Begin VB.Menu Nueva_Venta 
         Caption         =   "Nueva Venta"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Mnt_Ventas_Ver 
         Caption         =   "Entrar en Caja"
      End
      Begin VB.Menu Ver_Ventas 
         Caption         =   "Ver Ventas"
         Shortcut        =   {F9}
      End
      Begin VB.Menu Mnt_Devoluciones 
         Caption         =   "Devoluciones"
         Begin VB.Menu Nueva_Devolucion 
            Caption         =   "Nueva Devolución"
            Shortcut        =   {F11}
         End
         Begin VB.Menu Ver_Devoluciones 
            Caption         =   "Ver Devoluciones"
         End
      End
      Begin VB.Menu Mnt_Vales 
         Caption         =   "Vales"
         Begin VB.Menu Nuevo_Vales 
            Caption         =   "Nuevo Vale"
            Shortcut        =   {F12}
         End
         Begin VB.Menu Ver_Vales 
            Caption         =   "Ver Vales"
         End
      End
      Begin VB.Menu Mnu_Arreglos 
         Caption         =   "Arreglos"
         Begin VB.Menu Nuevo_Arreglo 
            Caption         =   "Nuevo Arreglo"
         End
         Begin VB.Menu Ver_Arreglos 
            Caption         =   "Ver Arreglos"
         End
      End
      Begin VB.Menu Pruebas_Mercancia 
         Caption         =   "Pruebas Mercancía"
         Begin VB.Menu Nueva_Prueba 
            Caption         =   "Nueva Prueba"
         End
         Begin VB.Menu Ver_Pruebas 
            Caption         =   "Ver Pruebas"
         End
      End
      Begin VB.Menu Deudas_Clientes 
         Caption         =   "Deudas Clientes"
         Begin VB.Menu Cobrar_Deudas 
            Caption         =   "Cobrar Deudas"
         End
      End
      Begin VB.Menu Informes_Ventas 
         Caption         =   "Informes"
         Begin VB.Menu Ticket_Resumen_Ventas 
            Caption         =   "Ticket Resumen de Ventas"
         End
      End
   End
   Begin VB.Menu Cierre_Caja 
      Caption         =   "Caja"
      Begin VB.Menu Caja_Pagos 
         Caption         =   "Pagos"
         Begin VB.Menu Caja_IntroducirPagos 
            Caption         =   "Introducir Pagos"
         End
         Begin VB.Menu Caja_VerPagos 
            Caption         =   "Ver Pagos"
         End
      End
      Begin VB.Menu Movimientos_Caja 
         Caption         =   "Movimientos Caja"
      End
      Begin VB.Menu Calcular_Cierre 
         Caption         =   "Calcular Cierre"
      End
      Begin VB.Menu Ver_Cierres 
         Caption         =   "Ver Cierres"
         Begin VB.Menu Vista_Formulario 
            Caption         =   "Vista Formulario"
         End
         Begin VB.Menu Vista_Rejilla 
            Caption         =   "Vista Rejilla"
         End
      End
      Begin VB.Menu Caj_Dev_Pro 
         Caption         =   "Devoluciones a proveedores"
      End
   End
   Begin VB.Menu Utilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu Inicio_Sesion 
         Caption         =   "Inicio de Sesión"
      End
      Begin VB.Menu Transferencias_Magnetico 
         Caption         =   "Transferencias en soporte magnético"
      End
      Begin VB.Menu Trabajar_Datos 
         Caption         =   "Trabajar con Datos"
         Begin VB.Menu Trabajar_Datos_Pedidos 
            Caption         =   "Pedidos"
         End
         Begin VB.Menu Trabajar_Datos_Ventas 
            Caption         =   "Ventas"
         End
         Begin VB.Menu Trabajar_Datos_Transferencias 
            Caption         =   "Transferencias"
         End
      End
      Begin VB.Menu Crea_Inventario 
         Caption         =   "Inventario"
      End
      Begin VB.Menu Opt_Varios 
         Caption         =   "Varios"
         Begin VB.Menu Comprimir_Fichero 
            Caption         =   "Comprimir Fichero"
         End
         Begin VB.Menu Descomprimir_Fichero 
            Caption         =   "Descomprimir Fichero"
         End
      End
      Begin VB.Menu Copia_de_seguridad 
         Caption         =   "Copia de seguridad"
      End
      Begin VB.Menu Configuracion 
         Caption         =   "Configuración"
         Begin VB.Menu Mnt_Empresas 
            Caption         =   "Empresas"
         End
         Begin VB.Menu Puesto_Actual 
            Caption         =   "Puesto Actual"
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : FrmInicio
' Fecha/Hora  : 30/12/2003 13:48
' Autor       : JCASTILLO
' Propósito   : Inicio de la aplicación.
'---------------------------------------------------------------------------------------
' para no cerrar la aplicación al cambiar de usuario
Dim es_inicio As Boolean


'---------------------------------------------------------------------------------------
' Procedimiento : Actualizar_Iva_Click
' Fecha/Hora    : 03/03/2004 11:23
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
Private Sub Actualizar_Iva_Click()

Dim rc As ADODB.Recordset
Dim entrans As Boolean

 
   On Error GoTo Actualizar_Iva_Click_Error

If MsgBox("Se va a proceder a la actualización del IVA de compra desde el pedido. ¿Continuar?", vbQuestion + vbYesNo, titulo) = vbYes Then

    With locCnn
        If .State = 1 Then .Close
            .CursorLocation = adUseServer
            .Open strLocCnn
            .BeginTrans
            entrans = True
    End With
    
    
    Set rc = New ADODB.Recordset
    
    rc.Open "SELECT CODIGO, TEMPOR, IVACOM FROM MAARTIC WHERE TEMPOR = " & TemporadaActual, locCnn, adOpenDynamic, adLockOptimistic
    
    Do Until rc.EOF
    
        rc.fields("IVACOM") = Obtiene_IVA_Pedido(rc.fields("CODIGO"), rc.fields("TEMPOR"), 0, 0, locCnn)
        rc.MoveNext
    
    Loop
    
    rc.Close
    Set rc = Nothing
    
    With locCnn
            .CommitTrans
            .Close
           ' .CursorLocation = adUseClient
           ' .Open strLocCnn
    End With
    
    entrans = False
    MsgBox "El IVA de compra se han actualizado correctamente", vbInformation + vbYesNo

Exit Sub
End If


   On Error GoTo 0
   Exit Sub

Actualizar_Iva_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Actualizar_Iva_Click de Formulario FrmInicio"

End Sub

Private Sub Actualizar_Precios_Compra_Click()
Dim rc As ADODB.Recordset
Dim entrans As Boolean

   On Error GoTo Actualizar_Precios_Compra_Click_Error

If MsgBox("Se va a proceder a la actualización del precio de compra desde el pedido. ¿Continuar?", vbQuestion + vbYesNo, titulo) = vbYes Then

    With locCnn
        If .State = 1 Then .Close
            .CursorLocation = adUseServer
            .Open strLocCnn
            .BeginTrans
            entrans = True
    End With
    
    
    Set rc = New ADODB.Recordset
    
    rc.Open "SELECT CODIGO, TEMPOR, PRECOM FROM MAARTIC WHERE TEMPOR = " & TemporadaActual, locCnn, adOpenDynamic, adLockOptimistic
    
    Do Until rc.EOF
    
        rc.fields("PRECOM") = Obtiene_Precom_Pedido(rc.fields("CODIGO"), rc.fields("TEMPOR"), 0, 0, locCnn)
        rc.MoveNext
    
    Loop
    
    rc.Close
    Set rc = Nothing
    
    With locCnn
            .CommitTrans
            .Close
           ' .CursorLocation = adUseClient
           ' .Open strLocCnn
    End With
    
    entrans = False
    MsgBox "Los precios de compra se han actualizado correctamente", vbInformation + vbYesNo

Exit Sub
End If

   On Error GoTo 0
   Exit Sub

Actualizar_Precios_Compra_Click_Error:

    If entrans Then locCnn.RollbackTrans
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Actualizar_Precios_Compra_Click de Formulario FrmInicio"

End Sub

Private Sub Ajusta_Numero_Pedido_Click()

    Call ajusta_pedidos(InputBox("Ultimo", "Ultimo", 1000))

End Sub

Private Sub Caj_Dev_Pro_Click()

Call ExecCmdNoFocus("uwdevprov.exe")

End Sub

Private Sub Caja_IntroducirPagos_Click()


Call ExecCmdNoFocus("uwpagcaj.exe")
    
   ' With frmCabPagPro
   '     .Pago_A_Proveedor = False
    '    .Show
   ' End With

End Sub

Private Sub Caja_VerPagos_Click()

With frmFlexPagPro
    .Pago_A_Proveedor = False
    .Show
End With

End Sub

Private Sub Calcular_Cierre_Click()
    frmCierreDia.Show
    DoEvents
    'Unload frmMenuTactil
End Sub

Private Sub Cobrar_Deudas_Click()

'Call ExecCmdNoFocus("uwprestaropa.exe")
'frmCabPag.Show

End Sub

Private Sub Compras_RecibirMercancia_Click()
    frmPedProv.Show
    DoEvents
    'Unload frmMenuTactil
End Sub

Private Sub Comprimir_Fichero_Click()
Dim forigen As String
Dim fdestino As String

   On Error GoTo Comprimir_Fichero_Click_Error

With FrmInicio.Dialogo
     .DialogTitle = "Abrir fichero ..."
     .Filter = "Fichero (*.*)|*.*|"
    .ShowOpen
     .Filter = ""
End With
    
     If (FrmInicio.Dialogo.CancelError = True) Or (Trim(FrmInicio.Dialogo.filename) = "") Then Exit Sub
     forigen = FrmInicio.Dialogo.filename
     
With FrmInicio.Dialogo
     .DialogTitle = "Guardar como ..."
     .Filter = "Fichero comprimido zlib (*.*)|*.*|"
    .ShowSave
     .Filter = ""
End With

  If (FrmInicio.Dialogo.CancelError = True) Or (Trim(FrmInicio.Dialogo.filename) = "") Then Exit Sub
     fdestino = FrmInicio.Dialogo.filename
    
    Call CompressFile(forigen, fdestino)
    
    MsgBox "Se ha comprimido satisfactoriamente", vbInformation, titulo

   On Error GoTo 0
   Exit Sub

Comprimir_Fichero_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Comprimir_Fichero_Click de Formulario FrmInicio"

End Sub

Private Sub Consulta_Detallada_Trn_Click()

 With frmFlexDtrans
    .Show
 End With

End Sub

Private Sub Consulta_Pedidos_Click()
    frmFlexPed.Show
    DoEvents
    'Unload frmMenuTactil
End Sub

Private Sub Consulta_Pedidos_Totales_Click()

frmTotPed.Show

End Sub

Private Sub Consulta_Trn_Totales_Click()
frmTotTransf.Show
End Sub

Private Sub Copia_de_seguridad_Click()
frmBackupSQL.Show
End Sub

Private Sub Crea_Inventario_Click()
    frmInventario.Show
End Sub

Private Sub Crear_Etiquetas_Click()
    frmEtiqLibre.Show
    DoEvents
    'Unload frmMenuTactil
End Sub

Private Sub dep_ConsultaStock_Click()
frmFlexStockTallCol.Show
End Sub

Private Sub Descomprimir_Fichero_Click()
Dim forigen As String
Dim fdestino As String

   On Error GoTo Descomprimir_Fichero_Click_Error

With FrmInicio.Dialogo
     .DialogTitle = "Abrir fichero comprimido ..."
     .Filter = "Fichero comprimido zlib (*.*)|*.*|"
    .ShowOpen
     .Filter = ""
End With
    
     If (FrmInicio.Dialogo.CancelError = True) Or (Trim(FrmInicio.Dialogo.filename) = "") Then Exit Sub
     forigen = FrmInicio.Dialogo.filename
     
With FrmInicio.Dialogo
     .DialogTitle = "Guardar como ..."
     .Filter = "Fichero (*.*)|*.*|"
    .ShowSave
     .Filter = ""
End With

  If (FrmInicio.Dialogo.CancelError = True) Or (Trim(FrmInicio.Dialogo.filename) = "") Then Exit Sub
     fdestino = FrmInicio.Dialogo.filename
    
    Call DecompressFile(forigen, fdestino)
    
    MsgBox "Se ha descomprimido satisfactoriamente", vbInformation, titulo

   On Error GoTo 0
   Exit Sub

Descomprimir_Fichero_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Descomprimir_Fichero_Click de Formulario FrmInicio"

End Sub

Private Sub Inicio_Sesion_Click()

es_inicio = True
Unload Me
frmInicSesion.Show

End Sub

Private Sub Introducir_Pagos_Click()

With frmCabPagPro
    .Pago_A_Proveedor = True
    .Show
End With

End Sub

Private Sub IVA_Click()

frmMntIva.Show
DoEvents
'Unload frmMenuTactil

End Sub


Private Sub MDIForm_Load()

Me.Caption = titulo & "          Usuario: " & devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual)
'If frmMenuTactil.Visible = False Then frmMenuTactil.Show

'ver tipo de permiso
Select Case TipoPermiso

'usuario restringido
Case 0

    Mantenimientos.Visible = False
    Compras.Visible = False
    dep_ConsultaStock.Visible = True
    Configuracion.Visible = False

'supervisor (todo)
Case 1

    'Show
    'frmMntAvisos.Show
    dep_ConsultaStock.Visible = False

End Select

Show
frmVerFecha.Show 1

End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not es_inicio Then
    If UsarTimerIP > 0 Then KillTimer &H0, UsarTimerIP
    Call cerrar_conexiones
    Set FrmInicio = Nothing
    End
End If

es_inicio = False

End Sub

Private Sub MDIForm_Resize()

'Me.WindowState = vbMaximized

End Sub

Private Sub Mnt_Almacenes_Click()
frmMntAlm.Show
DoEvents
'Unload frmMenuTactil

End Sub

Private Sub Mnt_Articulos_Click()

FrmMntArt.Show
DoEvents
'Unload frmMenuTactil

End Sub

Private Sub Mnt_Bancos_Click()
frmMntBan.Show
End Sub

Private Sub Mnt_Cajas_Click()
frmMntCaj.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_CatTallas_Click()

frmMntCatTall.Show
DoEvents
'Unload frmMenuTactil

End Sub

Private Sub Mnt_Centros_Click()
frmMntCen.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Clientes_Click()
frmMntCli.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Colores_Click()
frmMntCol.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Costureras_Click()
frmMntCost.Show
DoEvents
'Unload frmMenuTactil
End Sub



Private Sub Mnt_Empresas_Click()
frmMntEmp.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Familias_Click()
FrmMntFam.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_FormasCobro_Click()
frmMntFco.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_FormasPago_Click()
FrmMntFpa.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Ofertas_Click()
    frmMntOfe.Show
End Sub

Private Sub Mnt_Personal_Click()
frmMntPer.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Proveedores_Click()
FrmMntProv.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Secciones_Click()
FrmMntSec.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Sectores_Click()
frmMntSct.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_SubFamilias_Click()
FrmMntSfm.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Tallas_Click()
frmMntTall.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Tarjetas_Click()
frmMntTCre.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Mnt_Temporadas_Click()
frmMntTem.Show
DoEvents
'Unload frmMenuTactil
End Sub



Private Sub Mnt_Ventas_Ver_Click()

  
With frmCabVen
    .Entrar_Agregando = False
    .Show 1
End With

End Sub

Private Sub Movimientos_Caja_Click()
FrmMovCaja.Show
DoEvents
'Unload frmMenuTactil
End Sub

Private Sub Nueva_Devolucion_Click()

frmMntDev.Show
DoEvents
'Unload frmMenuTactil

End Sub

Private Sub Nueva_Prueba_Click()

Call ExecCmdNoFocus("uwprestaropa.exe")
'frmCabPrue.Show

End Sub

Private Sub Nueva_Venta_Click()

 With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
      
   Call inserta_saldo_inicial(Date, locCnn)

With frmCabVen
    .Entrar_Agregando = True
    .Show 1
End With

End Sub

Private Sub Nuevo_Arreglo_Click()

With frmNuArr
    .Sel_ID = 0
    .mi_Codart = 0
    .mi_Tempor = 0
    .mi_talla = 0
    .mi_Color = 0
    .Solo_Actualizar = False
    .Show 1
End With

End Sub

Private Sub Nuevo_Vales_Click()
frmMntVal.Show

DoEvents
'Unload frmMenuTactil

End Sub



Private Sub PTrans_Crear_Click()

With frmPtrans
    .añade_nueva = True
    .Show
End With

DoEvents
'Unload frmMenuTactil
End Sub



Private Sub Puesto_Actual_Click()
    frmConfPue.Show
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Recibir_Transferencia_Fichero_Click
' Fecha/Hora    : 20/02/2004 18:01
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
Private Sub Recibir_Transferencia_Fichero_Click()

  On Error GoTo Recibir_Transferencia_Fichero_Click_Error

  With FrmInicio.Dialogo
     .DialogTitle = "Incorporar transferencia desde fichero ..."
     .Filter = "Transferencias (*.trnz)|*.trnz|"
    .InitDir = "A:\"
    .ShowOpen
     .Filter = ""
    End With
    
     If (FrmInicio.Dialogo.CancelError = False) And (Trim(FrmInicio.Dialogo.filename <> "")) Then
     
     With locCnn
     If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
     End If
     End With

     'comprobar si existe en el directorio local, y borrar
     If Dir("C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle) <> "" Then
     
        'quitar atributos que pudiera tener de solo lectura
        SetAttr "C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle, vbNormal
        
        'borrar
        Kill "C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle
        
     End If
     
     'copiar el fichero al disco
     FileCopy FrmInicio.Dialogo.filename, "C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle
        
     'quitar atributos que pudiera tener de solo lectura
     SetAttr "C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle, vbNormal
     
     
     If Leer_TRN_Datos_Clave("C:\TRANSFERENCIAS\" & FrmInicio.Dialogo.FileTitle, locCnn, True) = 0 Then
          DoEvents
          MsgBox "La transferencia se ha incorporado correctamente", vbInformation
     End If

     End If

   On Error GoTo 0
   Exit Sub

Recibir_Transferencia_Fichero_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Recibir_Transferencia_Fichero_Click de Formulario FrmInicio"
  
End Sub

Private Sub Salir_Click()
If MsgBox("¿Desea salir de la aplicación?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
Unload Me
End Sub

Private Sub Ticket_Resumen_Ventas_Click()

frmImpRes.Show 1

End Sub

Private Sub Trabajar_Datos_Pedidos_Click()

With frmTrabaDatos
    .Show
    .Tab1 = 0
End With

End Sub

Private Sub Trabajar_Datos_Transferencias_Click()

With frmTrabaDatos
    .Show
    .Tab1 = 2
End With

End Sub

Private Sub Trabajar_Datos_Ventas_Click()

With frmTrabaDatos
    .Show
    .Tab1 = 1
End With

End Sub

Private Sub Transferencias_Magnetico_Click()
frmTransSop.Show 1
End Sub

Private Sub Ver_Arreglos_Click()
frmFlexArre.Show
End Sub



Private Sub Ver_Devoluciones_Click()
frmFlexDev.Show
End Sub

Private Sub Ver_Existencias_Click()

frmFlexStockTallCol.Show

End Sub

Private Sub Ver_Pagos_Click()

With frmFlexPagPro
    .Pago_A_Proveedor = True
    .Show
End With

End Sub

Private Sub Ver_Trn_Pendientes_Click()

With frmPtrans
    .añade_nueva = False
    .Show
End With

DoEvents
'Unload frmMenuTactil

End Sub


Private Sub Ver_Vales_Click()

frmFlexVal.Show

End Sub

Private Sub Ver_Ventas_Click()

frmFlexVen.Show

End Sub

Private Sub Vista_Formulario_Click()

With frmCierreDia
    .consultar = True
    .Show
End With

End Sub

Private Sub Vista_Rejilla_Click()

frmFlexCie.Show

End Sub
