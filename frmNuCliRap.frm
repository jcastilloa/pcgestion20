VERSION 5.00
Begin VB.Form frmNuCliRap 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Añadir Cliente Rápido"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miText ioNOMBRE 
      Height          =   525
      Left            =   1065
      TabIndex        =   0
      Top             =   75
      Width           =   3945
      _extentx        =   6959
      _extenty        =   926
      font            =   "frmNuCliRap.frx":0000
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   1575
      TabIndex        =   3
      Top             =   1530
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmNuCliRap.frx":002C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmNuCliRap.frx":0058
      picn            =   "frmNuCliRap.frx":0076
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   2550
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1530
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmNuCliRap.frx":0D52
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmNuCliRap.frx":0D7E
      picn            =   "frmNuCliRap.frx":0D9C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioNIF 
      Height          =   525
      Left            =   1065
      TabIndex        =   1
      Top             =   585
      Width           =   1545
      _extentx        =   2725
      _extenty        =   926
      font            =   "frmNuCliRap.frx":1678
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   15
      Top             =   1110
      Width           =   5040
      _extentx        =   8890
      _extenty        =   661
      caption         =   ""
      fount           =   "frmNuCliRap.frx":16A4
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.miText ioTELEFONO 
      Height          =   525
      Left            =   3705
      TabIndex        =   2
      Top             =   585
      Width           =   1305
      _extentx        =   2302
      _extenty        =   926
      font            =   "frmNuCliRap.frx":16D2
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      Height          =   300
      Left            =   2565
      TabIndex        =   7
      Top             =   675
      Width           =   1125
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIF"
      Height          =   300
      Left            =   555
      TabIndex        =   6
      Top             =   690
      Width           =   450
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RAZO /  NOMBRE"
      Height          =   555
      Left            =   105
      TabIndex        =   4
      Top             =   15
      Width           =   930
   End
End
Attribute VB_Name = "frmNuCliRap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmNuCliRap
' Fecha/Hora  : 18/01/2004 15:10
' Autor       : JCASTILLO
' Propósito   : Añadir cliente Rápido. Añade un cliente solo con los campos NOMBRE y NIF
'               (solo usar para los formularios de ventas)
'---------------------------------------------------------------------------------------

Option Explicit

Public ID_Cliente_Creado As Long
Public Caja_Cliente As Long  'para cuando es un cliente que ya existe
                             'y ha sido creado desde otra caja
Public RAZO_Creado As String

Dim creado As Boolean

Private Sub cbAceptar_Click()

If creado And cbAceptar.Caption = "Terminar" Then
    DoEvents
    Unload Me
End If

End Sub

Private Sub cbAceptar_GotFocus()

DoEvents
Call cbAceptar_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DoEvents
End Sub





Private Sub iotelefono_Validate(Cancel As Boolean)
Dim nif As New clsNIF

'si esta a blancos salir
   'On Error GoTo ioNIF_Validate_Error

'If Trim(ioNIF.Text) = "" Then
'    ioNIF.CancelarValidacion
'    Cancel = True
'    Exit Sub
'End If

If Trim(ioNIF.Text) <> "" And Trim(ioNIF.Text) <> "0" Then

nif.DarFormato = True
nif.nif = ioNIF.Text

If nif.Err Then
    ioNIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioNIF.Text = nif.nif
End If

Set nif = Nothing

End If

'intentar aceptar el cliente ...
Call crea_cliente

cbAceptar.Caption = "Terminar"
DoEvents
cbAceptar.SetFocus


   On Error GoTo 0
   Exit Sub

ioNIF_Validate_Error:
    Set nif = Nothing
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioNIF_Validate of Formulario frmMntCli"
 
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : crea_cliente
' Fecha/Hora  : 18/01/2004 19:56
' Autor       : JCASTILLO
' Propósito   : Crea un nuevo cliente rápidamente con los datos introducidos
'---------------------------------------------------------------------------------------
Private Sub crea_cliente()

Dim tmpcodigo As Variant
'Dim cm As New ADODB.Command
Dim numreg As Long

Dim tmpcom As Variant

creado = False

'Dim rc As New ADODB.Recordset

'validaciones ...
   'On Error GoTo crea_cliente_Error

If Trim(ioNOMBRE.Text = "") Then
    lblstatus.Caption = "RAZO/Nombre Incorrecto"
    ioNOMBRE.SetFocus
    ioNOMBRE.CancelarValidacion
    Exit Sub
End If

If Trim(ioNIF.Text = "") And Trim(ioTELEFONO.Text = "") Then
    lblstatus.Caption = "Rellene NIF o TELEFONO"
    ioNIF.SetFocus
    ioNIF.CancelarValidacion
    Exit Sub
End If

If Trim(ioNIF.Text = "") And Trim(ioTELEFONO.Text <> "") Then
    ioNIF.Text = "0"
ElseIf Trim(ioNIF.Text <> "") And Trim(ioTELEFONO.Text = "") Then
    ioTELEFONO.Text = "0"
End If

If Not IsNumeric(Trim(ioTELEFONO.Text)) Then
    lblstatus.Caption = "TELEFONO no válido"
    ioTELEFONO.SetFocus
    ioTELEFONO.CancelarValidacion
    Exit Sub
End If

'si YA existe el nif en la base e datos
If ioNIF.Text <> "0" Then
    
    tmpcom = devuelve_matriz("SELECT RAZO, CODIGO, CODCAJA FROM CLIENTES WHERE NIF = '" & ioNIF.Text & "'", locCnn)

    If IsArray(tmpcom) Then
        
        lblstatus.Caption = "¡NIF ya Existe!"
        ioNIF.SetFocus
        ioNIF.CancelarValidacion
        
        'asignar el cliente si ya existe
        If MsgBox("El NIF ya existe. Pertenece al cliente:" & Chr(13) & Trim(tmpcom(0)) & Chr(13) & "¿Desea asignar el cliente?", vbQuestion + vbYesNo) = vbYes Then
            ID_Cliente_Creado = tmpcom(1)
            Caja_Cliente = tmpcom(2)
            RAZO_Creado = Trim(tmpcom(0))
            creado = True
        End If
        
        Exit Sub
    End If
    
    'tmpcom = ""

End If

'si ya existe el telefono en la base de datos
If ioTELEFONO.Text <> "0" Then
    
    ReDim tmpcom(0)
    tmpcom = devuelve_matriz("SELECT RAZO, CODIGO, CODCAJA FROM CLIENTES WHERE TELEFONO1 = '" & ioTELEFONO.Text & "'", locCnn)

    If IsArray(tmpcom) Then
        lblstatus.Caption = "¡Teléfono ya Existe!"
        ioNIF.SetFocus
        ioNIF.CancelarValidacion
        
        'asignar el cliente si ya existe
        If MsgBox("El Teléfono ya existe. Pertenece al cliente:" & Chr(13) & Trim(tmpcom(0)) & Chr(13) & "¿Desea asignar el cliente?", vbQuestion + vbYesNo) = vbYes Then
            ID_Cliente_Creado = tmpcom(1)
            Caja_Cliente = tmpcom(2)
            RAZO_Creado = Trim(tmpcom(0))
            creado = True
        End If
        
        Exit Sub
    End If
    
    'tmpcom = ""

End If

'obtener el siguiente ID de cliente
tmpcodigo = devuelve_campo("select max(codigo) + 1 from clientes where codcaja = " & CajaActual, locCnn)

If tmpcodigo = "@" Then tmpcodigo = 1

'insertar cliente ...
locCnn.Execute "INSERT INTO CLIENTES (CODIGO, CODCAJA, RAZO, NIF, TELEFONO1) VALUES(" & tmpcodigo & ", " & CajaActual & ", '" & ioNOMBRE.Text & "', '" & ioNIF.Text & "','" & ioTELEFONO.Text & "')"

'With rc
'    .ActiveConnection = locCnn
'    .Open "Select top 1 * from clientes", , adOpenDynamic, adLockOptimistic
'    .AddNew
'    .Fields("CODIGO") = tmpcodigo
'    .Fields("CODCAJA") = CajaActual
'    .Fields("RAZO") = ioNOMBRE.Text
'    .Fields("NIF") = ioNIF.Text
'    .Update
'    DoEvents
'    .Close
'End With
'Set rc = Nothing
    
DoEvents

'Set cm = Nothing

'DoEvents

'devolver valores ...
ID_Cliente_Creado = tmpcodigo
Caja_Cliente = CajaActual
RAZO_Creado = ioNOMBRE.Text

creado = True

lblstatus.Caption = "Cliente creado corretamente (Pulse Enter)"

'End If

DoEvents

   On Error GoTo 0
   Exit Sub

crea_cliente_Error:

    lblstatus.Caption = "Error al crear el cliente"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento crea_cliente de Formulario frmNuCliRap"

End Sub

Private Sub cbCancelar_Click()

ID_Cliente_Creado = 0
Caja_Cliente = 0
RAZO_Creado = ""

Unload Me

End Sub

Private Sub Form_Load()

'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

With ioNOMBRE
    .LongMaxima = 40
    .PermitirBlanco = False
End With

With ioNIF
    .LongMaxima = 15
    '.PermitirBlanco = False
    .Alineacion = 1
End With

With ioTELEFONO
    .LongMaxima = 9
    '.PermitirBlanco = False
    .Alineacion = 1
End With

End Sub

