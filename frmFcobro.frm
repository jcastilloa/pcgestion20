VERSION 5.00
Begin VB.Form frmFcobro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Forma de Cobro ..."
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11205
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11205
   Begin PCGestion.chameleonButton cmRecibirDevolLista 
      Height          =   1620
      Left            =   7320
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   2858
      BTYPE           =   3
      TX              =   "Buscar Devolución en Lista"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFcobro.frx":0000
      PICN            =   "frmFcobro.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioPRIMERA 
      Height          =   525
      Left            =   2094
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1320
      _extentx        =   2328
      _extenty        =   926
      font            =   "frmFcobro.frx":0576
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   9480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   45
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFcobro.frx":05A2
      PICN            =   "frmFcobro.frx":05BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   9480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   855
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFcobro.frx":1298
      PICN            =   "frmFcobro.frx":12B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioSEGUNDA 
      Height          =   525
      Left            =   2094
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1320
      _extentx        =   2328
      _extenty        =   926
      font            =   "frmFcobro.frx":1B8E
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.miCombo cbFCOBRO 
      Height          =   495
      Left            =   2079
      TabIndex        =   0
      Top             =   585
      Width           =   4005
      _extentx        =   7064
      _extenty        =   873
      font            =   "frmFcobro.frx":1BBA
   End
   Begin PCGestion.miCombo cbTARJETAS 
      Height          =   495
      Left            =   2079
      TabIndex        =   1
      Top             =   1095
      Width           =   4005
      _extentx        =   7064
      _extenty        =   873
      font            =   "frmFcobro.frx":1BE6
   End
   Begin PCGestion.bsGradientLabel lblImporte 
      Height          =   465
      Left            =   2109
      Top             =   60
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   820
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   192
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblImporteMax 
      Height          =   3990
      Left            =   15
      Top             =   3660
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   7038
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   6298138
      Colour1         =   16558731
      Colour2         =   12632256
      CaptionAlignment=   1
      TextShadowColour=   12632064
      TextShadow      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblCambio 
      Height          =   465
      Left            =   2102
      Top             =   2745
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   820
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   16711680
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cmRecibirVale 
      Height          =   1620
      Left            =   7326
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   2858
      BTYPE           =   3
      TX              =   "Recibir &Vale"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFcobro.frx":1C12
      PICN            =   "frmFcobro.frx":1C2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   15
      Top             =   7665
      Width           =   11160
      _ExtentX        =   15610
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   375
      Left            =   6225
      Top             =   1710
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "- V - Recibir Vale"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblValesCount 
      Height          =   375
      Left            =   6225
      Top             =   2835
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Vales (0)"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel3 
      Height          =   375
      Left            =   6225
      Top             =   2085
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "- A - Aceptar"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel5 
      Height          =   375
      Left            =   6225
      Top             =   2460
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "- C - Cancelar"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cmRecibirDevol 
      Height          =   1620
      Left            =   6210
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2858
      BTYPE           =   3
      TX              =   "Recibir ultima Devolución"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFcobro.frx":2188
      PICN            =   "frmFcobro.frx":21A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   375
      Left            =   8475
      Top             =   1695
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      Caption         =   "- D - Devolución"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.miText ioINTROP 
      Height          =   525
      Left            =   4806
      TabIndex        =   3
      Top             =   1620
      Width           =   1320
      _extentx        =   2328
      _extenty        =   926
      font            =   "frmFcobro.frx":26FE
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.miText ioINTROS 
      Height          =   525
      Left            =   4806
      TabIndex        =   5
      Top             =   2175
      Width           =   1320
      _extentx        =   2328
      _extenty        =   926
      font            =   "frmFcobro.frx":272A
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.bsGradientLabel lblDevolCount 
      Height          =   375
      Left            =   6225
      Top             =   3210
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Devoluciones (0)"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin VB.Label lblINTROS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INTRODUCE"
      Height          =   270
      Left            =   3396
      TabIndex        =   17
      Top             =   2235
      Width           =   1290
   End
   Begin VB.Label lblINTROP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INTRODUCE"
      Height          =   270
      Left            =   3396
      TabIndex        =   16
      Top             =   1695
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAMBIO"
      Height          =   330
      Left            =   1074
      TabIndex        =   15
      Top             =   2820
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   330
      Left            =   1089
      TabIndex        =   14
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TARJETA"
      Height          =   330
      Left            =   969
      TabIndex        =   13
      Top             =   1185
      Width           =   1080
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F. COBRO"
      Height          =   330
      Left            =   984
      TabIndex        =   12
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label lblSegunda 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SEGUNDA"
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   2265
      Width           =   1065
   End
   Begin VB.Label lblPrimera 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRIMERA"
      Height          =   270
      Left            =   960
      TabIndex        =   10
      Top             =   1710
      Width           =   1050
   End
End
Attribute VB_Name = "frmFcobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     :  frmFcobro
' Fecha/Hora :28/01/2004 17:55
' Autor      :   JCastillo
' Propósito  :  Seleccionar forma de cobro y calcular cuanto va para la primera
'                   o segunda y como se distribuye
'---------------------------------------------------------------------------------------
Option Explicit

' FORMAS DE COBRO

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

'11 DEVOLUCION                    -> vale
'12 CONTADO Y DEVOLUCION  -> contado y vale
'13 DEVOLUCION Y TARJETA       -> vale y tarjeta
'14 DEVOLUCION Y VALE            -> contado y vale


Public I_Primera As Currency
Public I_Segunda As Currency
Public Importe As Currency
Public F_Cobro As Integer
Public I_Cambio As Currency
Public I_Comision As Currency

'ID del usuario de la venta actual
Public ID_Dependiente As Long

'para comprobar si ha sido asignado un cliente antes
'de aceptar la forma de pago pendiente ...
'''''''''''''''''''''''''''''''''''''''''''''''''
Public ID_Cliente As Long
Public Caja_Cliente As Long
'''''''''''''''''''''''''''''''''''''''''''''''''

'si sale por cancelar
Public Cancelado As Boolean

Private Type LstVales
    Codigo_Vale() As Long
    Caja_Vale() As Byte
    cuantos As Long
    impdcto() As Currency
End Type

'Private Type LstDevol
    'Codigo_Devol() As Long
    'Caja_Devol() As Byte
    'cuantos As Long
'End Type
       
Dim mis_Vales As LstVales
Dim mis_Devol As Long

Dim impor_vales As Double
Dim impor_devol As Double

Public impor_deuda As Double

Dim vale_en_segunda As Boolean
Dim devol_en_segunda As Boolean

Dim suma As Double

Dim falta_cliente As Boolean

'_______________________________________________________________
'para imprimir un nuevo vale con el importe en sobra_val
Dim sobra_val As Double
Dim falta_imp As Double
'_______________________________________________________________

Dim tmpfcobro As Integer

Private Sub cbAceptar_Click()
Dim tmpcodval As Long

If (ioINTROP.Visible Or ioINTROS.Visible) And (suma = 0) Then
    lblstatus.Caption = "Debe introducir forma de cobro y cantidades"
    cbFCOBRO.SetFocus
    Exit Sub
End If

If cbFCOBRO.Text = "" Then
    lblstatus.Caption = "No se permite Forma de Cobro en blanco"
    cbFCOBRO.SetFocus
    Exit Sub
End If

If ioINTROP.Visible Then
    If ioINTROP.Text = "" Then
        ioINTROP.SetFocus
        lblstatus.Caption = "No se permite INTRODUCE en BLANCO"
        Exit Sub
    End If
End If

If ioINTROS.Visible Then
    If ioINTROS.Text = "" Then
        ioINTROS.SetFocus
        lblstatus.Caption = "No se permite INTRODUCE en BLANCO"
        Exit Sub
    End If
End If

If falta_cliente Then
        lblstatus.Caption = "Debe volver a la pantalla anterior y ASIGNAR UN CLIENTE"
        cbFCOBRO.SetFocus
        Beep
        Exit Sub
End If

If falta_imp > 0 Then

    lblstatus.Caption = "Faltan " & Format(falta_imp, "Currency") & " para aceptar la venta"
    
    If ioPRIMERA.Visible = True Then
        ioPRIMERA.SetFocus
    ElseIf ioSEGUNDA.Visible Then
        ioSEGUNDA.SetFocus
    End If
    
    Exit Sub
End If

'si sale por error, no continuar
If procesa_forma_cobro(cbFCOBRO.Text) = True Then
    Exit Sub
End If

If sobra_val > 0 Then
    'si se desea imprimir un nuevo vale ...
    If MsgBox("Se debe imprimir un nuevo vale por un importe de " & Format(sobra_val, "Currency") & " ¿Desea hacerlo ahora?", vbQuestion + vbYesNo, titulo) = vbYes Then
    
      tmpcodval = añadir_vale(CLng(frmCabVen.ioCODIGO.Caption), ID_Dependiente, 0, 0, sobra_val, 0, 0, 3, Null, locCnn)
      DoEvents
      Call Imprime_Vale(tmpcodval, CajaActual, locCnn)
    
    End If
End If

Cancelado = False
Unload Me

End Sub

Private Sub cbCancelar_Click()

Cancelado = True
Unload Me

End Sub


Private Sub cbFCOBRO_GotFocus()

If cbFCOBRO.Text <> "" Then
    tmpfcobro = cbFCOBRO.Text
Else
    tmpfcobro = 0
End If

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbFCOBRO_lostfocus
' Fecha/Hora  : 05/02/2004 15:37
' Autor       : JCASTILLO
' Propósito   : Preparar los controles del formulario dependiendo de la
'               forma de pago
'---------------------------------------------------------------------------------------
Private Sub cbFCOBRO_lostfocus()
   
   On Error GoTo cbFCOBRO_lostfocus_Error
   
'If Me.Visible = False Then Exit Sub
            
If cbFCOBRO.Text = "" Then
    cbFCOBRO.SetFocus
    lblstatus.Caption = "No se permite Forma de Cobro en BLANCO"
    Exit Sub
Else
    If tmpfcobro = cbFCOBRO.Text Then Exit Sub
    lblstatus.Caption = ""
End If

tmpfcobro = cbFCOBRO.Text

'limpiar datos ...

'limpiar devoluciones ...
mis_Devol = 0
lblDevolCount.Caption = "Devoluciones (0)"
impor_devol = 0


'limpiar vales
ReDim mis_Vales.Caja_Vale(0)
ReDim mis_Vales.Codigo_Vale(0)

mis_Vales.cuantos = 0
lblValesCount.Caption = "Vales (0)"

impor_vales = 0

impor_deuda = 0
falta_cliente = False

devol_en_segunda = False
vale_en_segunda = False

'limpiar variables
I_Segunda = 0
I_Primera = 0
sobra_val = 0
falta_imp = 0
suma = 0
I_Cambio = 0
I_Comision = 0

'limpiar textboxs
ioPRIMERA.Text = "0"
ioSEGUNDA.Text = "0"
ioINTROP.Text = "0"
ioINTROS.Text = "0"

DoEvents

lblCambio.Caption = ""

DoEvents

cbTARJETAS.Visible = False

Select Case cbFCOBRO.Text

'CONTADO
Case 1
    
    lblPrimera.Caption = "Contado"
    
    lblPrimera.Visible = True
    ioPRIMERA.Locked = True
    
    ioPRIMERA.Visible = True
    ioINTROP.Visible = True
    ioINTROP.Locked = False
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False
    ioSEGUNDA.Visible = False
    lblINTROS.Visible = False
    
    cbTARJETAS.Visible = False
    Label1.Visible = False
    
    cmRecibirVale.Visible = False
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    vale_en_segunda = False
    devol_en_segunda = False
        
    'introducir el total en ioprimera
    ioPRIMERA.Text = Importe
    
    DoEvents
    
    If ioINTROP.Visible Then ioINTROP.SetFocus
    DoEvents

'TARJETA
Case 2

    lblPrimera.Caption = "Tarjeta"
    ioPRIMERA.Locked = True
        
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Visible = False
    lblINTROP.Visible = False
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False
    ioSEGUNDA.Visible = False
    lblINTROP.Visible = False
    
    cbTARJETAS.Visible = True
    cbTARJETAS.Enabled = True
    cbTARJETAS.Locked = False
    
    Label1.Visible = True
    
    'introducir el total en ioprimera
    ioPRIMERA.Text = Importe

    cmRecibirVale.Visible = False
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    vale_en_segunda = False
    devol_en_segunda = False
    
    DoEvents
    
    cbTARJETAS.Text = "1"
    SendKeys "{TAB}"

'CONTADO y TARJETA
Case 3

    lblPrimera.Caption = "Contado"
    lblSegunda.Caption = "Tarjeta"
    ioPRIMERA.Locked = True
    ioSEGUNDA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioSEGUNDA.Visible = True
    ioINTROS.Visible = False
    ioINTROS.Locked = True
    lblINTROS.Visible = False
        
    cbTARJETAS.Visible = True
    Label1.Visible = True
    
    cmRecibirVale.Visible = False
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    cbTARJETAS.SetFocus
    
    vale_en_segunda = False
    devol_en_segunda = False
        
    ioPRIMERA.Text = Importe
    
    cbTARJETAS.Text = "1"
    SendKeys "{TAB}"
    
'VALE
Case 4

    lblPrimera.Caption = "Vale"
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    
    ioINTROP.Visible = True
    ioINTROP.Locked = True
    lblINTROP.Visible = True
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False
    ioSEGUNDA.Visible = False
    lblINTROS.Visible = False
    
    'lblSegunda.Visible = False
    'ioSEGUNDA.Visible = False
    cbTARJETAS.Visible = False
    Label1.Visible = False
    cmRecibirVale.Visible = True
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    vale_en_segunda = False
    devol_en_segunda = False
    
    ioPRIMERA.Text = Importe

'CONTADO y VALE
Case 5

    lblPrimera.Caption = "Contado"
    lblSegunda.Caption = "Vale"
    
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioINTROS.Visible = True
    lblINTROS.Visible = True
    ioINTROS.Locked = True
    ioSEGUNDA.Visible = True
    
    ioSEGUNDA.Locked = True
   ' lblSegunda.Visible = True
  '  ioSEGUNDA.Visible = True
    cbTARJETAS.Visible = False
    Label1.Visible = False
    
    cmRecibirVale.Visible = True
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    vale_en_segunda = True
    devol_en_segunda = False
    
    ioPRIMERA.Locked = False
    
    ioPRIMERA.Text = Importe
    
    If ioINTROP.Visible Then ioINTROP.SetFocus

'TARJETA y VALE
Case 6
    
    lblPrimera.Caption = "Tarjeta"
    lblSegunda.Caption = "Vale"
          
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Visible = True
    ioINTROP.Locked = True
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioINTROS.Visible = True
    ioSEGUNDA.Visible = True
    lblINTROS.Visible = True
    ioINTROS.Locked = True
        
    ioPRIMERA.Locked = True
    ioSEGUNDA.Locked = True
   ' lblSegunda.Visible = True
   ' ioSEGUNDA.Visible = True
    cbTARJETAS.Visible = True
    Label1.Visible = True
    cbTARJETAS.SetFocus
    
    cmRecibirVale.Visible = True
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    
    vale_en_segunda = True
    devol_en_segunda = False
    
    ioPRIMERA.Text = Importe
    
'PENDIENTE
Case 7

    If (ID_Cliente = 0 Or Caja_Cliente = 0) Then
        lblstatus.Caption = "Debe volver a la pantalla anterior y ASIGNAR UN CLIENTE"
        cbFCOBRO.SetFocus
        falta_cliente = True
        Beep
        Exit Sub
    End If
    
    lblPrimera.Caption = "Pendiente"
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    
    ioINTROP.Visible = False
    ioINTROP.Locked = False
    lblINTROP.Visible = False
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False
    ioSEGUNDA.Visible = False
    lblINTROS.Visible = False
    
    'lblSegunda.Visible = False
    'ioSEGUNDA.Visible = False
    cbTARJETAS.Visible = False
    Label1.Visible = False
    
    vale_en_segunda = False
    devol_en_segunda = False
    
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    cmRecibirVale.Visible = False
    
    ioPRIMERA.Text = Importe
    
    
    
'CONTADO y PENDIENTE
Case 8

    
    If (ID_Cliente = 0 Or Caja_Cliente = 0) Then
        lblstatus.Caption = "Debe volver a la pantalla anterior y ASIGNAR UN CLIENTE"
        cbFCOBRO.SetFocus
        Beep
        falta_cliente = True
        Exit Sub
    End If
       
    lblPrimera.Caption = "Contado"
    lblSegunda.Caption = "Pendiente"
    
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False

    ioSEGUNDA.Visible = True
    ioSEGUNDA.Locked = True
    
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    cmRecibirVale.Visible = False

    cbTARJETAS.Visible = False
    Label1.Visible = False
    
    vale_en_segunda = False
    devol_en_segunda = False
    
    ioPRIMERA.Locked = False
    
    ioPRIMERA.Text = Importe
    
    If ioINTROP.Visible Then ioINTROP.SetFocus

'Pendiente y TARJETA
Case 9

    
    If (ID_Cliente = 0 Or Caja_Cliente = 0) Then
        lblstatus.Caption = "Debe volver a la pantalla anterior y ASIGNAR UN CLIENTE"
        cbFCOBRO.SetFocus
        Beep
        falta_cliente = True
        Exit Sub
    End If
    
    lblPrimera.Caption = "Pendiente"
    lblSegunda.Caption = "Tarjeta"
    ioPRIMERA.Locked = True
    ioSEGUNDA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioSEGUNDA.Visible = True
    ioINTROS.Visible = False
    ioINTROS.Locked = True
    lblINTROS.Visible = False
    
    'ioINTROS.Visible = True
    'ioSEGUNDA.Visible = True
    
    cbTARJETAS.Visible = True
    Label1.Visible = True
    
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    cmRecibirVale.Visible = False
    
    cbTARJETAS.SetFocus
    vale_en_segunda = False
    devol_en_segunda = False
        
    ioPRIMERA.Text = Importe
    
    cbTARJETAS.Text = "1"
    SendKeys "{TAB}"


'PENDIENTE y VALE
Case 10

    If (ID_Cliente = 0 Or Caja_Cliente = 0) Then
        lblstatus.Caption = "Debe volver a la pantalla anterior y ASIGNAR UN CLIENTE"
        cbFCOBRO.SetFocus
        Beep
        falta_cliente = True
        Exit Sub
    End If
        
    lblPrimera.Caption = "Pendiente"
    lblSegunda.Caption = "Vale"
    
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioINTROS.Visible = True
    lblINTROS.Visible = True
    ioINTROS.Locked = True
    ioSEGUNDA.Visible = True
    
    ioSEGUNDA.Locked = True
   ' lblSegunda.Visible = True
  '  ioSEGUNDA.Visible = True
    cbTARJETAS.Visible = False
    Label1.Visible = False
    cmRecibirVale.Visible = True
    cmRecibirDevol.Visible = False
    cmRecibirDevolLista.Visible = False
    vale_en_segunda = True
    devol_en_segunda = False
    ioPRIMERA.Locked = False
    
    ioPRIMERA.Text = Importe
    
    If ioINTROP.Visible Then ioINTROP.SetFocus


'DEVOLUCION
Case 11
   
    lblPrimera.Caption = "Devolución"
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    
    ioINTROP.Visible = False
    ioINTROP.Locked = False
    lblINTROP.Visible = False
    
    lblSegunda.Visible = False
    ioINTROS.Visible = False
    ioSEGUNDA.Visible = False
    lblINTROS.Visible = False
    
    'lblSegunda.Visible = False
    'ioSEGUNDA.Visible = False
    cbTARJETAS.Visible = False
    Label1.Visible = False
    devol_en_segunda = False
    vale_en_segunda = False
    
    
    cmRecibirDevol.Visible = True
    cmRecibirDevolLista.Visible = True
    cmRecibirVale.Visible = False
    
    ioPRIMERA.Text = Importe
    
'CONTADO Y DEVOLUCION
Case 12

    lblPrimera.Caption = "Contado"
    lblSegunda.Caption = "Devolución"
    
    ioPRIMERA.Locked = True
    
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Locked = False
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioINTROS.Visible = True
    lblINTROS.Visible = True
    ioINTROS.Locked = True
    ioSEGUNDA.Visible = True
    
    ioSEGUNDA.Locked = True
   ' lblSegunda.Visible = True
  '  ioSEGUNDA.Visible = True
    cbTARJETAS.Visible = False
    Label1.Visible = False
    
    cmRecibirVale.Visible = False
    cmRecibirDevol.Visible = True
    cmRecibirDevolLista.Visible = True
    
    vale_en_segunda = False
    devol_en_segunda = True
    
    ioPRIMERA.Locked = False
    
    ioPRIMERA.Text = Importe
    
    If ioINTROP.Visible Then ioINTROP.SetFocus


'TARJETA Y DEVOLUCION
Case 13
    
    lblPrimera.Caption = "Tarjeta"
    lblSegunda.Caption = "Devolución"
          
    lblPrimera.Visible = True
    ioPRIMERA.Visible = True
    ioINTROP.Visible = True
    ioINTROP.Locked = True
    lblINTROP.Visible = True
    
    lblSegunda.Visible = True
    ioINTROS.Visible = True
    ioSEGUNDA.Visible = True
    lblINTROS.Visible = True
    ioINTROS.Locked = True
        
    ioPRIMERA.Locked = True
    ioSEGUNDA.Locked = True
   ' lblSegunda.Visible = True
   ' ioSEGUNDA.Visible = True
    cbTARJETAS.Visible = True
    Label1.Visible = True
    cbTARJETAS.SetFocus
    
    cmRecibirVale.Visible = False
    cmRecibirDevol.Visible = True
    cmRecibirDevolLista.Visible = True
    
    vale_en_segunda = False
    devol_en_segunda = True
    
    ioPRIMERA.Text = Importe

    
End Select

'1 CONTADO
'2 TARJETA
'3 CONTADO Y TARJETA
'4 VALE
'5 CONTADO Y VALE
'6 TARJETA Y VALE
'7 PENDIENTE            -> tarjeta
'8 CONTADO Y PENDIENTE  -> contado y vale
'9 PENDIENTE Y TARJETA  -> contado y tarjeta
'10 PENDIENTE Y VALE    -> contado y vale

'11 DEVOLUCION                    -> vale
'12 CONTADO Y DEVOLUCION  -> contado y vale
'13 TARJETA Y DEVOLUCION         ->   tarjeta y vale
' ******** 14 DEVOLUCION Y VALE            -> contado y vale


   On Error GoTo 0
   Exit Sub

cbFCOBRO_lostfocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbFCOBRO_lostfocus de Formulario frmFcobro"

End Sub

Private Sub cbFCOBRO_Validate(Cancel As Boolean)

If cbFCOBRO.Text <> "" Then
    Call cbFCOBRO_lostfocus
Else
    Cancel = True
    cbFCOBRO.SetFocus
End If

End Sub

Private Sub cbTARJETAS_LostFocus()

If cbTARJETAS.Visible = False Then Exit Sub

If cbTARJETAS.Text = "" Then
    cbTARJETAS.SetFocus
    lblstatus.Caption = "No se permite Tipo de Tarjeta en BLANCO"
    Exit Sub
Else
    lblstatus.Caption = ""
    If ioINTROP.Visible Then ioINTROP.SetFocus
End If

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : cmRecibirDevol_Click
' Fecha/Hora  : 05/02/2004 15:56
' Autor       : JCASTILLO
' Propósito   : Recibir devolución
'---------------------------------------------------------------------------------------
Private Sub cmRecibirDevol_Click()

    
   On Error GoTo cmRecibirDevol_Click_Error

    Call recibir_devoluciones(True, 0, 0)
    
   On Error GoTo 0
   Exit Sub

cmRecibirDevol_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmRecibirDevol_Click de Formulario frmFcobro"
   
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : recibir_devoluciones
' Fecha/Hora     : 05/07/2004 11:14
' Autor             : JCastillo
' Propósito     :   Subrutina para recibir devoluciones, si entra por nultimas = true
'                       entonces, pregunta el nº d devoluciones a coger (cogiendo las n ultimas
'                       que se hayan echo. si es = false, coge codigo_devol y codigo_caja
'                       y devuelve esa especificamente.
'---------------------------------------------------------------------------------------
'
Private Sub recibir_devoluciones(nultimas As Boolean, codigo_devol As Long, codigo_caja As Byte)

Dim T_Vale As Variant
Dim T_modelo As String
Dim Num_Dev As String

Dim rc_dev As New ADODB.Recordset

 
   On Error GoTo recibir_devoluciones_Error

    If nultimas Then

        Num_Dev = InputBox("¿Número de prendas a devolver?", "Insertar Devolución", "1")
    
        If Num_Dev = "" Then
            Set rc_dev = Nothing
            Exit Sub
        End If
    
        If Not IsNumeric(Num_Dev) Then
            Set rc_dev = Nothing
            Exit Sub
        End If

        'si es mayor de 10, posiblemente este mal y intenta introducir un
        'codigo de vale
        If CLng(Num_Dev) > 10 Then
    
            'si quiere cancelar la operación
            If MsgBox("Ha seleccionado " & Num_Dev & " prendas a devolver. ¿Esta ud. seguro de que no es un error?. Pulse SI para cancelar o NO para continuar.", vbQuestion + vbYesNo, titulo) = vbYes Then
                Set rc_dev = Nothing
                Exit Sub
            End If
    
        End If
       
        'seleccionar datos
        rc_dev.Open "SELECT TOP " & Num_Dev & " CODART, TEMPOR, MOTIVO, CODVAL, IMPORTE, CODVEN from DEVOL WHERE CODCAJA = " & CajaActual & " ORDER BY CODIGO DESC", locCnn, adOpenDynamic, adLockOptimistic
    
    
    'si selecciona por codigo y caja directamente ...
    Else
        
        'seleccionar datos
        rc_dev.Open "SELECT CODART, TEMPOR, MOTIVO, CODVAL, IMPORTE, CODVEN from DEVOL WHERE CODCAJA = " & codigo_caja & " AND CODIGO = " & codigo_devol & " ORDER BY CODIGO DESC", locCnn, adOpenDynamic, adLockOptimistic
        
        
    End If
    'obtener el motivo y articulo de la ultima devolución
    'T_Devol = devuelve_matriz("SELECT TOP 1 CODART, TEMPOR, MOTIVO from DEVOL WHERE CODCAJA = " & CajaActual & " ORDER BY CODIGO DESC", locCnn)

    'DoEvents
    'obtener el importe del último vale
    'T_Vale = devuelve_matriz("SELECT TOP 1 CODIGO, IMPORTE from VALES WHERE CODCAJA = " & CajaActual & " ORDER BY CODIGO DESC", locCnn)
    
    'DoEvents
    
    Do Until rc_dev.EOF
    
    'seleccionar solo los vales de devolucion
    T_Vale = devuelve_matriz("SELECT CODIGO, IMPORTE from VALES WHERE CODCAJA = " & CajaActual & " AND CODIGO = " & rc_dev.fields("CODVAL") & " ORDER BY CODIGO DESC", locCnn)
     
    'primero buscar en la matriz de vales de esta sesion, para comprobar
    'que no haya sido ya introducido en esta misma venta ...
    'TRUE si ya existe el vale ...
    
    'si no hay vale, sacar el importe de la propia devolucion, y codigo de vale = 0
    If Not IsArray(T_Vale) Then
        ReDim T_Vale(2)
        T_Vale(0) = 0
        T_Vale(1) = rc_dev.fields("IMPORTE")
        'insertar el código de venta
        If Trim(frmCabVen.ioCODIGO.Caption) <> "" Then
            rc_dev.fields("CODVEN") = frmCabVen.ioCODIGO.Caption
            rc_dev.Update
        End If
    End If
   

    If (buscar_vale_matriz(CLng(T_Vale(0)), CajaActual) = True) Then
            
            MsgBox "¡La devolucion seleccionada ya ha sido introducida para esta misma venta!" & Chr(13) & _
                   "Motivo: " & Trim(rc_dev.fields("MOTIVO")) & Chr(13) & _
                   "Importe: " & Format(T_Vale(1), "Currency"), vbExclamation, titulo
            
    'si no se encuentra en la venta actual, añadir ...
    Else
    
        'obtener el modelo del artículo
        T_modelo = devuelve_campo("SELECT MODELO from MAARTIC WHERE CODIGO = " & rc_dev.fields("CODART") & " AND TEMPOR = " & rc_dev.fields("TEMPOR"), locCnn)
    
        'preguntar al usuario
        If MsgBox("Desea añadir una Devolución por un importe de: " & Format(T_Vale(1), "Currency") & Chr(13) & _
            "Modelo: " & Trim(T_modelo) & Chr(13) & _
            "Motivo: " & Trim(rc_dev.fields("MOTIVO")), vbQuestion + vbYesNo) = vbNo Then Exit Sub

        'si existe un vale para esta devolución
        If T_Vale(0) > 0 Then
        
            'incrementar el contador de vales
            mis_Vales.cuantos = mis_Vales.cuantos + 1

            'redimensionar
            ReDim Preserve mis_Vales.Codigo_Vale(mis_Vales.cuantos)
            ReDim Preserve mis_Vales.Caja_Vale(mis_Vales.cuantos)

            'añadir un nuevo vale a la matriz de vales para esta venta
            mis_Vales.Codigo_Vale(mis_Vales.cuantos) = T_Vale(0)
            mis_Vales.Caja_Vale(mis_Vales.cuantos) = CajaActual
            
            'incrementar el importe de los vales
            impor_vales = impor_vales + T_Vale(1)
           
        End If
        
            'descomponer el codigo de devolución
            mis_Devol = mis_Devol + 1

            'redimensionar
          '  ReDim Preserve mis_Devol.Codigo_Vale(mis_Devol.cuantos)
          '  ReDim Preserve mis_Devol.Caja_Vale(mis_Devol.cuantos)

          '  mis_Devol.Codigo_Devol(mis_Devol.cuantos) = Left(tmpcodigo, 9)
          '  mis_Devol.Caja_Devol(mis_Devol.cuantos) = Right(tmpcodigo, 3)

            impor_devol = impor_devol + CDbl(T_Vale(1))
        
      End If
        
    'pasar al siguiente registro
    If Not rc_dev.EOF Then rc_dev.MoveNext
    
    Loop
    
    'actualizar el campo que corresponda en pantalla
    If vale_en_segunda Then
        ioINTROS.Text = impor_vales
    Else
        ioINTROP.Text = impor_vales
    End If
    
    If devol_en_segunda Then
        ioINTROS.Text = impor_devol
    Else
        ioINTROP.Text = impor_devol
    End If



    Call ioINTROS_lostfocus
    Call ioINTROP_lostfocus

T_modelo = ""
Num_Dev = ""

rc_dev.Close
Set rc_dev = Nothing

lblValesCount.Caption = "Vales(" & mis_Vales.cuantos & ")"
lblDevolCount.Caption = "Devoluciones(" & mis_Devol & ")"

If ioPRIMERA.Visible Then
    ioPRIMERA.SetFocus
ElseIf ioSEGUNDA.Visible Then
    ioSEGUNDA.SetFocus
End If


   On Error GoTo 0
   Exit Sub

recibir_devoluciones_Error:

    If rc_dev.State = 1 Then rc_dev.Close
    Set rc_dev = Nothing
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento recibir_devoluciones de Formulario frmFcobro"

End Sub

Private Sub cmRecibirDevolLista_Click()

   On Error GoTo cmRecibirDevolLista_Click_Error

With frmFlexDev
        .desde_ventas = True
        .Show 1
        
        If (.ID_Devolucion > 0 And .Caja_Devolucion > 0) Then
            Call recibir_devoluciones(False, .ID_Devolucion, .Caja_Devolucion)
            DoEvents
        End If
                
        Set frmFlexDev = Nothing
End With

   On Error GoTo 0
   Exit Sub

cmRecibirDevolLista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmRecibirDevolLista_Click de Formulario frmFcobro"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmRecibirVale_Click
' Fecha/Hora  : 08/02/2004 12:55
' Autor       : JCASTILLO
' Propósito   : Recibir un vale desde el código de vale, y hacer las comprobaciones
'               de que no haya caducado, o haya sido aceptado previamente
'---------------------------------------------------------------------------------------
Private Sub cmRecibirVale_Click()
Dim tmpcodigo As String
Dim tmpsalida As Variant
Dim tmpimporte As Double
Dim tmpsdcto As String

'9 digitos + 3 de la caja
   On Error GoTo cmRecibirVale_Click_Error

tmpcodigo = InputBox("Introduzca Código de Vale", "Introducir Vale")

If tmpcodigo = "" Then
    lblstatus.Caption = "Código de Vale no válido"
    Exit Sub
End If

If Not IsNumeric(tmpcodigo) Then
    lblstatus.Caption = "Código de Vale no válido"
    Exit Sub
End If

tmpcodigo = Format(tmpcodigo, "000000000000")

'primero buscar en la matriz de vales de esta sesion, para comprobar
'que no haya sido ya introducido en esta misma venta ...
'TRUE si ya existe el vale ...
If buscar_vale_matriz(CLng(Left(tmpcodigo, 9)), CByte(Right(tmpcodigo, 3))) = True Then
        MsgBox "¡El vale " & tmpcodigo & " ya ha sido introducido para esta misma venta!", vbExclamation, titulo
        Exit Sub
End If
    
'estados : 0=pendiente, 1=aceptado
tmpsalida = devuelve_matriz("SELECT IMPORTE, ESTADO, CADUCA, FMODI, DCTO FROM VALES WHERE CODIGO = " & Left(tmpcodigo, 9) & " AND CODCAJA =" & Right(tmpcodigo, 3), locCnn)

'si existe ...
If IsArray(tmpsalida) Then

        If tmpsalida(0) = 0 Then
            tmpimporte = (CDbl(Replace(ioPRIMERA.Text, ".", ",")) * tmpsalida(4)) / 100
            tmpsdcto = "(dcto  " & tmpsalida(4) & "%)"
        Else
            tmpimporte = tmpsalida(0)
        End If
        
        
        
    'si el vale ya esta aceptado, preguntar si se desea continuar
    If tmpsalida(1) = True Then
            
        If MsgBox("¿Desea aceptar el vale " & tmpcodigo & " por un importe de: " & Format(tmpimporte, "Currency") & " ? " & tmpsdcto & Chr(13) & "Nota: EL VALE YA HA SIDO ACEPTADO con fecha: " & tmpsalida(3), vbQuestion + vbYesNo, titulo) = vbNo Then
            
            tmpcodigo = ""
            Exit Sub
        End If
    'preguntar si desea aceptar el vale ...
    ElseIf MsgBox("¿Desea aceptar el vale " & tmpcodigo & " por un importe de: " & Format(tmpimporte, "Currency") & " ? " & tmpsdcto, vbQuestion + vbYesNo, titulo) = vbNo Then
        
        tmpcodigo = ""
        Exit Sub
    
    End If
    
    'si esta caducado, preguntar otra vez ...
    If tmpsalida(2) < Date Then
    
        If MsgBox("EL VALE YA HA CADUCADO, ¿desea aceptarlo de todas formas ? " & tmpsdcto & Chr(13) & "Caducó el: " & tmpsalida(2), vbQuestion + vbYesNo, titulo) = vbNo Then
            
            tmpcodigo = ""
            Exit Sub
        End If
    
    End If
    
    'descomponer el codigo de vale
    mis_Vales.cuantos = mis_Vales.cuantos + 1

    'redimensionar
    ReDim Preserve mis_Vales.Codigo_Vale(mis_Vales.cuantos)
    ReDim Preserve mis_Vales.Caja_Vale(mis_Vales.cuantos)
    ReDim Preserve mis_Vales.impdcto(mis_Vales.cuantos)
    

    mis_Vales.Codigo_Vale(mis_Vales.cuantos) = Left(tmpcodigo, 9)
    mis_Vales.Caja_Vale(mis_Vales.cuantos) = Right(tmpcodigo, 3)
    
     'si se ha introducido solo el % para el dcto,
     'hallar el importe de ese descuento, y pasarselo a la matriz
     'para que actualize el vale metiendo en el campo importe
     'el importe de ese descuento (para que salga reflejado en el cierre de caja
   If tmpsalida(0) = 0 Then
        mis_Vales.impdcto(mis_Vales.cuantos) = tmpimporte
    End If

    impor_vales = impor_vales + CDbl(tmpimporte)
    
    If vale_en_segunda Then
        
        ioINTROS.Text = impor_vales
        ioINTROP.Text = "0"
        Call ioINTROS_lostfocus
        Call ioINTROP_lostfocus
        
    Else
        
        ioINTROP.Text = impor_vales
        Call ioINTROP_lostfocus
        Call ioINTROS_lostfocus
        
    End If
    
    
    lblValesCount.Caption = "Vales(" & mis_Vales.cuantos & ")"
        
Else

    lblstatus.Caption = "El vale no se encuentra en la base de datos"
    Exit Sub

End If

If ioPRIMERA.Visible Then
    ioPRIMERA.SetFocus
ElseIf ioSEGUNDA.Visible Then
    ioSEGUNDA.SetFocus
End If

tmpcodigo = ""
lblstatus.Caption = ""

   On Error GoTo 0
   Exit Sub

cmRecibirVale_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmRecibirVale_Click de Formulario frmFcobro"

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : buscar_vale_matriz
' Fecha/Hora  : 08/02/2004 13:02
' Autor       : JCASTILLO
' Propósito   : Comprueba que el vale no exista ya en la matriz de vales de la sesión
'               actual, para no aceptar el mismo vale 2 veces seguidas...
'               devuelve TRUE si ha encontrado el vale.
'---------------------------------------------------------------------------------------
Private Function buscar_vale_matriz(codval As Long, codcaja As Byte) As Boolean
Dim var As Long

   On Error GoTo buscar_vale_matriz_Error
  
    With mis_Vales
        For var = 1 To .cuantos
    
            'si coincide ...
            If (.Caja_Vale(var) = codcaja) And (.Codigo_Vale(var) = codval) Then
            
                buscar_vale_matriz = True
                Exit Function
                
            End If
    
        Next
    End With
    
   On Error GoTo 0
   Exit Function

buscar_vale_matriz_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento buscar_vale_matriz de Formulario frmFcobro"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'recibir vale
Case vbKeyV

 If cmRecibirVale.Visible Then
    Call cmRecibirVale_Click
 End If
 KeyCode = 0
 
'recibir devolución
Case vbKeyD

 If cmRecibirDevol.Visible Then
    Call cmRecibirDevol_Click
 End If
 KeyCode = 0

'cancelar
Case vbKeyC
 Call cbCancelar_Click
 KeyCode = 0

'aceptar
Case vbKeyA
 Call cbAceptar_Click
 KeyCode = 0
 
End Select

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

With ioPRIMERA
   ' .PermitirBlanco = False
    .SoloNumeros = True
    .dspFormat = "Currency"
    .Alineacion = 1
End With

With ioSEGUNDA
   ' .PermitirBlanco = False
    .SoloNumeros = True
    .dspFormat = "Currency"
    .Alineacion = 1
End With

With ioINTROP
   ' .PermitirBlanco = False
    .SoloNumeros = True
    .dspFormat = "Currency"
    .Alineacion = 1
End With

With ioINTROS
   ' .PermitirBlanco = False
    .SoloNumeros = True
    .dspFormat = "Currency"
    .Alineacion = 1
End With

With cbFCOBRO
        .LenCodigo = 2
        .CodigoWidth = 500
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM FCOBRO WHERE MBAJA = 0 ORDER BY CODIGO"
        .carga
        .Text = "1"
End With

With cbTARJETAS
        .LenCodigo = 2
        .CodigoWidth = 500
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM TARJETAS WHERE MBAJA = 0 ORDER BY CODIGO"
        .carga
End With

cbTARJETAS.Visible = False
lblImporte.Caption = Format(Importe, "Currency")
lblImporteMax.Caption = lblImporte.Caption

DoEvents

Call cbFCOBRO_lostfocus
DoEvents

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : procesa_forma_cobro
' Fecha/Hora     : 05/02/2004 09:58
' Autor             : JCastillo
' Propósito       :  Procesa la forma de cobro, devolviendo los importes y realizando
'                   Devuelve TRUE si no es correcto el proceso (salidap por error)
'---------------------------------------------------------------------------------------
Private Function procesa_forma_cobro(codigo As Integer) As Boolean
Dim T_tasa As Integer
Dim tmp_prim As Double

   On Error GoTo procesa_forma_cobro_Error

        Select Case codigo
        
        'CONTADO, insertar todo el importe en primera, segunda a 0
        Case 1
        
           If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
            ioPRIMERA.SetFocus
            lblstatus.Caption = "No se permite importe en blanco o cero"
            procesa_forma_cobro = True
            Exit Function
          End If
          
          ioSEGUNDA.Text = "0"
                  
          I_Primera = Importe
          I_Segunda = 0
        
        'TARJETA
        Case 2
        
        
          If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
            ioPRIMERA.SetFocus
            lblstatus.Caption = "No se permite importe en blanco o cero"
            procesa_forma_cobro = True
            Exit Function
          End If
          
          If cbTARJETAS.Text = "" Then
            cbTARJETAS.SetFocus
            lblstatus.Caption = "No se permite Tarjeta en blanco"
            procesa_forma_cobro = True
            Exit Function
          End If
          
          ioSEGUNDA.Text = "0"
                  
          T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
                   
          tmp_prim = Importe
          
          I_Comision = ((tmp_prim * T_tasa) / 100)
          'obtener la tasa de la tarjeta, y descontarsela al importe
          I_Primera = tmp_prim '- I_Comision
          I_Segunda = 0
        
        'contado y tarjeta
        Case 3
        
            'Contado
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Tarjeta
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
           If cbTARJETAS.Text = "" Then
              cbTARJETAS.SetFocus
              lblstatus.Caption = "No se permite Tarjeta en blanco"
              procesa_forma_cobro = True
              Exit Function
           End If
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
            
            tmp_prim = ioSEGUNDA.Text
            
            I_Comision = ((tmp_prim * T_tasa) / 100)
        
            I_Primera = ioINTROP.Text
            'obtener la tasa de la tarjeta, y descontarsela al importe
            I_Segunda = tmp_prim '- I_Comision
        
        
        Case 4  'Vale
        
           'Contado
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
            'asignar la cantidad
            'ioPRIMERA.Text = Importe
            
            I_Primera = Importe
            I_Segunda = 0
               
       'contado y vale
       Case 5
       
            'Contado
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Vale
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'asignar la cantidad
                     
            'ioSEGUNDA.Text = impor_vales
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            I_Primera = ioINTROP.Text
            I_Segunda = ioSEGUNDA.Text
        
        'tarjeta y vale
        Case 6
        
        
            'Tarjeta
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Vale
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
                   
            tmp_prim = ioINTROP.Text
            
            I_Comision = ((tmp_prim * T_tasa) / 100)
          
          'obtener la tasa de la tarjeta, y descontarsela al importe
            I_Primera = tmp_prim '- I_Comision
            I_Segunda = ioSEGUNDA.Text
            
            
        'Pendiente
        Case 7
        
         
          If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
            ioPRIMERA.SetFocus
            lblstatus.Caption = "No se permite importe en blanco o cero"
            procesa_forma_cobro = True
            Exit Function
          End If
          
          'If cbTARJETAS.Text = "" Then
          '  cbTARJETAS.SetFocus
          '  lblstatus.Caption = "No se permite Tarjeta en blanco"
          '  procesa_forma_cobro = True
          '  Exit Function
          'End If
          
          ioSEGUNDA.Text = "0"
                  
          'T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
                   
          tmp_prim = Importe
          impor_deuda = Importe
                    
          I_Primera = tmp_prim '- ((tmp_prim * T_tasa) / 100)
          I_Segunda = 0
          
       'CONTADO Y PENDIENTE
       Case 8
       
            'Contado
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'PENDIENTE
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'asignar la cantidad
                     
            'ioSEGUNDA.Text = impor_vales
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            I_Primera = ioINTROP.Text
            I_Segunda = ioSEGUNDA.Text
        
            impor_deuda = I_Segunda
                       
                       
        'PENDIENTE Y TARJETA
        Case 9
        
            'PENDIENTE
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Tarjeta
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
           If cbTARJETAS.Text = "" Then
              cbTARJETAS.SetFocus
              lblstatus.Caption = "No se permite Tarjeta en blanco"
              procesa_forma_cobro = True
              Exit Function
           End If
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
            
            tmp_prim = ioSEGUNDA.Text
            
            I_Comision = ((tmp_prim * T_tasa) / 100)
        
            I_Primera = ioINTROP.Text
            'obtener la tasa de la tarjeta, y descontarsela al importe
            I_Segunda = tmp_prim ' I_Comision
        
            impor_deuda = I_Primera
            
            
                        
       'PENDIENTE Y VALE
       Case 10
       
            'PENDIENTE
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Vale
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'asignar la cantidad
                     
            'ioSEGUNDA.Text = impor_vales
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            I_Primera = ioINTROP.Text
            I_Segunda = ioSEGUNDA.Text
        
            impor_deuda = I_Primera
      
      
      Case 11  'Devolución
        
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
            'asignar la cantidad
            'ioPRIMERA.Text = Importe
            
            I_Primera = Importe
            I_Segunda = 0
       
       
      'contado y devolucion
       Case 12
       
            'Contado
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Devolución
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'asignar la cantidad
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            I_Primera = ioINTROP.Text
            I_Segunda = ioSEGUNDA.Text
    
    
        'tarjeta y devolucion
        Case 13
                
            'Tarjeta
            If ioPRIMERA.Text = "" Or ioPRIMERA.Text = "0" Then
                ioPRIMERA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
        
            'Devol
            If ioSEGUNDA.Text = "" Or ioSEGUNDA.Text = "0" Then
                ioSEGUNDA.SetFocus
                lblstatus.Caption = "No se permite importe en blanco o cero"
                procesa_forma_cobro = True
                Exit Function
            End If
            
            If ioINTROP.Text = "" Then
              ioINTROP.SetFocus
              lblstatus.Caption = "No se permite INTRODUCE en blanco"
              procesa_forma_cobro = True
              Exit Function
            End If
            
            T_tasa = devuelve_campo("SELECT TASA FROM TARJETAS WHERE CODIGO = " & cbTARJETAS.Text, locCnn)
                   
            tmp_prim = ioINTROP.Text
            
            I_Comision = ((tmp_prim * T_tasa) / 100)
          
          'obtener la tasa de la tarjeta, y descontarsela al importe
            I_Primera = tmp_prim '- I_Comision
            I_Segunda = ioSEGUNDA.Text
                    
                       
        End Select
        
        
'1 CONTADO
'2 TARJETA
'3 CONTADO Y TARJETA
'4 VALE
'5 CONTADO Y VALE
'6 TARJETA Y VALE

'7 PENDIENTE            -> tarjeta
'8 CONTADO Y PENDIENTE  -> contado y vale
'9 PENDIENTE Y TARJETA  -> contado y tarjeta
'10 PENDIENTE Y VALE    -> contado y vale


'11 DEVOLUCION                    -> vale
'12 CONTADO Y DEVOLUCION  -> contado y vale
'13 DEVOLUCION Y TARJETA       -> vale y tarjeta
' ******** fuera 14 DEVOLUCION Y VALE            -> contado y vale

   On Error GoTo 0
   
   F_Cobro = cbFCOBRO.Text
   procesa_forma_cobro = False
   
   Exit Function
   

procesa_forma_cobro_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento procesa_forma_cobro de Formulario frmFcobro"
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Set frmFcobro = Nothing

End Sub

Private Sub ioINTROP_lostfocus()
Dim tmppri As Double
Dim tmpseg As Double


'si es contado
If ioINTROP.Text = "" Then Exit Sub
If Not IsNumeric(ioINTROP.Text) Then Exit Sub

lblCambio.Caption = ""

tmppri = ioINTROP.Text
If ioINTROS.Text <> "" Then tmpseg = ioINTROS.Text

suma = tmppri + tmpseg

lblCambio.Caption = ""

Select Case lblPrimera.Caption

Case "Contado"

    'If ioINTROP.Text <> "" Then

       falta_imp = 0
       'si es mayor del importe total ...
        If (tmppri > Importe) Then
            If ioSEGUNDA.Visible Then ioSEGUNDA.Text = "0"
            lblCambio.Caption = "Cambio: " & Format(tmppri - Importe, "Currency")
            I_Cambio = tmppri - Importe
        Else
            
            If ioSEGUNDA.Visible = False Then
                If (Importe - tmppri) > 0 Then lblCambio.Caption = "Falta: " & Format(Importe - tmppri, "Currency")
                falta_imp = Importe - tmppri
            Else
                ioSEGUNDA.Text = Importe - tmppri
            End If
                        
        End If
    
    'End If

Case "Vale"

    'If ioINTROP.Text <> "" Then

       sobra_val = 0
       falta_imp = 0
       
       If tmppri > Importe Then
            If ioSEGUNDA.Visible Then ioSEGUNDA.Text = "0"
            lblCambio.Caption = "Nuevo Vale: " & Format(tmppri - Importe, "Currency")
            sobra_val = tmppri - Importe
        Else
                     
            If ioSEGUNDA.Visible = False Then
               If (Importe - tmppri) > 0 Then lblCambio.Caption = "Falta: " & Format(Importe - tmppri, "Currency")
                falta_imp = Importe - tmppri
            Else
                'poner el importe restante en SEGUNDA
                If ioSEGUNDA.Visible Then ioSEGUNDA.Text = Importe - tmppri
            End If
            
            
        End If
    
    'End If
    

Case "Tarjeta"

        falta_imp = 0
        
        If tmppri > Importe Then
            If ioSEGUNDA.Visible Then ioSEGUNDA.Text = "0"
            lblCambio.Caption = "Sobra: " & Format(tmppri - Importe, "Currency")
        Else
            
            If ioSEGUNDA.Visible = False Then
               If (Importe - tmppri) > 0 Then lblCambio.Caption = "Falta: " & Format(Importe - tmppri, "Currency")
               falta_imp = Importe - tmppri
            Else
                'poner el importe restante en SEGUNDA
                If ioSEGUNDA.Visible Then ioSEGUNDA.Text = Importe - tmppri
            End If

        End If
        
Case "Pendiente"

        falta_imp = 0
        
        If tmppri > Importe Then
            If ioSEGUNDA.Visible Then ioSEGUNDA.Text = "0"
            lblCambio.Caption = "Sobra: " & Format(tmppri - Importe, "Currency")
        Else
            
            If ioSEGUNDA.Visible = False Then
               'If (Importe - tmppri) > 0 Then lblCambio.Caption = "Falta: " & Format(Importe - tmppri, "Currency")
               'falta_imp = Importe - tmppri
               falta_imp = 0
            Else
                'poner el importe restante en SEGUNDA
                If ioSEGUNDA.Visible Then ioSEGUNDA.Text = Importe - tmppri
            End If

        End If
        

Case "Devolución"
    
       sobra_val = 0
       falta_imp = 0
       
       If tmppri > Importe Then
            If ioSEGUNDA.Visible Then ioSEGUNDA.Text = "0"
            lblCambio.Caption = "Nuevo Vale: " & Format(tmppri - Importe, "Currency")
            sobra_val = tmppri - Importe
        Else
                     
            If ioSEGUNDA.Visible = False Then
               If (Importe - tmppri) > 0 Then lblCambio.Caption = "Falta: " & Format(Importe - tmppri, "Currency")
                falta_imp = Importe - tmppri
            Else
                'poner el importe restante en SEGUNDA
                If ioSEGUNDA.Visible Then ioSEGUNDA.Text = Importe - tmppri
            End If
            
            
        End If

End Select

DoEvents
End Sub


Private Sub ioINTROS_lostfocus()
Dim tmppri As Double
Dim tmpresto As Double
Dim tmpseg As Double
Dim difvale As Double
'Dim suma As Double

'si es contado
'If ioINTROP.Text <> "" Then tmppri = ioINTROP.Text

If ioINTROS.Text = "" Then Exit Sub
If Not IsNumeric(ioINTROS.Text) Then Exit Sub

If ioSEGUNDA.Text = "" Then Exit Sub
If Not IsNumeric(ioSEGUNDA.Text) Then Exit Sub

tmpseg = ioINTROS.Text
tmpresto = ioSEGUNDA.Text

suma = tmppri + tmpseg

Select Case lblSegunda.Caption

Case "Contado"

   ' If ioINTROS.Text <> "" Then

        falta_imp = 0
        'si es mayor que la cantidad a completar
        If tmpseg > tmpresto Then
            lblCambio.Caption = "Cambio: " & Format(tmpseg - tmpresto, "Currency")
            I_Cambio = tmpseg - tmpresto
        'si es menor ...
        Else
            If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")
            falta_imp = tmpresto - tmpseg
        End If
    
   ' End If

Case "Vale"

   ' If (ioINTROS.Text <> "") Then

       sobra_val = 0
       falta_imp = 0

        'si es mayor q la cantida a completar
        If (tmpseg > tmpresto) Then
          
          
            'si el cambio a devolver es menor que la cantidad
            'devolver dinero, de lo contrario, devolver todo el dinero
            'y un vale con la diferencia
          '  If tmppri > (suma - Importe) Then
            
                'devolver el cambio en metalico, es mayor la cantidad en metalico
                'que el cambio a devolver
            '    lblCambio.Caption = "Cambio: " & Format(suma - Importe, "Currency")
            
            'es menor la cantidad en metalico que el cambio a devolver.
            'devolver todo el metalico, y un vale hasta llegar al cambio
            'Else
                        
            '    lblCambio.Caption = "Cambio: " & Format(tmppri, "Currency") & " Vale: " & Format(suma - Importe, "Currency")
                        
            'End If
            
             'meter la cantidad faltante en introp si esta a 0

            
        
            lblCambio.Caption = "Nuevo Vale: " & Format(tmpseg - tmpresto, "Currency")
            sobra_val = tmpseg - tmpresto
            
        Else
              
              If ioINTROP.Visible Then
             
              If ioINTROP.Text <> "" Then
                 
                 'si es igual a 0, introducir ahi la cantidad q falta
                 If CDbl(ioINTROP.Text) = 0 Then
                 
                    ioINTROP.Text = tmpresto - tmpseg
                    Call ioINTROP_lostfocus
                 'sino, decir por pantalla
                 Else
                 
                    falta_imp = tmpresto - tmpseg
                    If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")

                 End If
                 
              End If
              
              End If
        
         '   lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")
            
        End If
    
   ' End If

Case "Tarjeta"

   falta_imp = 0
   sobra_val = 0

   If (tmpseg > tmpresto) Then
            sobra_val = tmpseg - tmpresto
            lblCambio.Caption = "Sobra: " & Format(tmpseg - tmpresto, "Currency")
   Else
            
            'meter la cantidad faltante en introp si esta a 0
            If ioINTROP.Visible Then
             
              If ioINTROP.Text <> "" Then
                 
                 'si es igual a 0, introducir ahi la cantidad q falta
                 If CDbl(ioINTROP.Text) = 0 Then
                    
                    ioINTROP.Text = tmpresto - tmpseg
                    Call ioINTROP_lostfocus
                    
                 'sino, decir por pantalla
                 Else
                                  
                 falta_imp = tmpresto - tmpseg
                 If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")

                 End If
                 
              End If
                          
            'sino, decir por pantalla
            Else
            
            falta_imp = tmpresto - tmpseg
            If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")
            
            End If
            
            
            
   End If


    Case "Pendiente"

    falta_imp = 0
    sobra_val = 0

   If (tmpseg > tmpresto) Then
            sobra_val = tmpseg - tmpresto
            lblCambio.Caption = "Sobra: " & Format(tmpseg - tmpresto, "Currency")
   Else
            
            'meter la cantidad faltante en introp si esta a 0
            If ioINTROP.Visible Then
             
              If ioINTROP.Text <> "" Then
                 
                 'si es igual a 0, introducir ahi la cantidad q falta
                 If CDbl(ioINTROP.Text) = 0 Then
                    
                    ioINTROP.Text = tmpresto - tmpseg
                    Call ioINTROP_lostfocus
                    
                 'sino, decir por pantalla
                 Else
                                  
                 falta_imp = tmpresto - tmpseg
                 If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")

                 End If
                 
              End If
                          
            'sino, decir por pantalla
            Else
            
            falta_imp = tmpresto - tmpseg
            If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")
            
            End If
            
            
            
   End If

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

Case "Devolución"

   ' If (ioINTROS.Text <> "") Then

       sobra_val = 0
       falta_imp = 0

        'si es mayor q la cantida a completar
        If (tmpseg > tmpresto) Then
          
          
            'si el cambio a devolver es menor que la cantidad
            'devolver dinero, de lo contrario, devolver todo el dinero
            'y un vale con la diferencia
          '  If tmppri > (suma - Importe) Then
            
                'devolver el cambio en metalico, es mayor la cantidad en metalico
                'que el cambio a devolver
            '    lblCambio.Caption = "Cambio: " & Format(suma - Importe, "Currency")
            
            'es menor la cantidad en metalico que el cambio a devolver.
            'devolver todo el metalico, y un vale hasta llegar al cambio
            'Else
                        
            '    lblCambio.Caption = "Cambio: " & Format(tmppri, "Currency") & " Vale: " & Format(suma - Importe, "Currency")
                        
            'End If
            
             'meter la cantidad faltante en introp si esta a 0

            
        
            lblCambio.Caption = "Nuevo Vale: " & Format(tmpseg - tmpresto, "Currency")
            sobra_val = tmpseg - tmpresto
            
        Else
              
              If ioINTROP.Visible Then
             
              If ioINTROP.Text <> "" Then
                 
                 'si es igual a 0, introducir ahi la cantidad q falta
                 If CDbl(ioINTROP.Text) = 0 Then
                 
                    ioINTROP.Text = tmpresto - tmpseg
                    Call ioINTROP_lostfocus
                 'sino, decir por pantalla
                 Else
                 
                    falta_imp = tmpresto - tmpseg
                    If (tmpresto - tmpseg) > 0 Then lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")

                 End If
                 
              End If
              
              End If
        
         '   lblCambio.Caption = "Falta: " & Format(tmpresto - tmpseg, "Currency")
            
        End If
    
   ' End If


End Select

DoEvents
End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : Marcar_Vales_Como_Aceptados
' Fecha/Hora  : 08/02/2004 19:00
' Autor       : JCASTILLO
' Propósito   : Marca como aceptados los vales que se van a aceptar en esta venta. Para
'               llamar de manera externa, una vez aceptada la venta y antes de descargar
'               el formulario
'---------------------------------------------------------------------------------------
Public Sub Marcar_Vales_Como_Aceptados(conexion As ADODB.Connection)
Dim var As Long

   On Error GoTo Marcar_Vales_Como_Aceptados_Error

    For var = 1 To mis_Vales.cuantos
    
        'poner el importe del dcto en el campo importe
        If mis_Vales.impdcto(var) > 0 Then
            conexion.Execute "UPDATE VALES SET IMPORTE = " & Replace(CDbl(mis_Vales.impdcto(var)), ",", ".") & " WHERE CODIGO = " & mis_Vales.Codigo_Vale(var) & " AND CODCAJA =" & mis_Vales.Caja_Vale(var)
        End If
    
        Call aceptar_vale(mis_Vales.Codigo_Vale(var), mis_Vales.Caja_Vale(var), conexion)
            
    Next var

   On Error GoTo 0
   Exit Sub

Marcar_Vales_Como_Aceptados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Marcar_Vales_Como_Aceptados de Formulario frmFcobro"
End Sub

