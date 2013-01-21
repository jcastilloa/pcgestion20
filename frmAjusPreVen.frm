VERSION 5.00
Begin VB.Form frmAjusPreVen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ajuste Precio de Venta"
   ClientHeight    =   3540
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5295
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
   ScaleHeight     =   3540
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   4320
      TabIndex        =   1
      Top             =   1845
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmAjusPreVen.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmAjusPreVen.frx":002C
      picn            =   "frmAjusPreVen.frx":004A
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
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2700
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmAjusPreVen.frx":0D26
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmAjusPreVen.frx":0D52
      picn            =   "frmAjusPreVen.frx":0D70
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblPreVenEx 
      Height          =   495
      Left            =   1695
      Top             =   735
      Width           =   1560
      _extentx        =   2514
      _extenty        =   873
      caption         =   ""
      fount           =   "frmAjusPreVen.frx":164C
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11311500
      captionalignment=   2
   End
   Begin PCGestion.miText ioPREVEN 
      Height          =   480
      Left            =   1665
      TabIndex        =   0
      Top             =   2460
      Width           =   1635
      _extentx        =   2884
      _extenty        =   847
      font            =   "frmAjusPreVen.frx":167A
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblPreVenSug 
      Height          =   495
      Left            =   1695
      Top             =   1860
      Width           =   1560
      _extentx        =   2514
      _extenty        =   873
      caption         =   ""
      fount           =   "frmAjusPreVen.frx":16A6
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11311500
      captionalignment=   2
   End
   Begin PCGestion.bsGradientLabel lblPVPIVA 
      Height          =   495
      Left            =   1695
      Top             =   3015
      Width           =   1560
      _extentx        =   2752
      _extenty        =   873
      caption         =   ""
      fount           =   "frmAjusPreVen.frx":16D4
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11311500
      captionalignment=   2
   End
   Begin PCGestion.bsGradientLabel lblPRECOM 
      Height          =   495
      Left            =   1695
      Top             =   120
      Width           =   1560
      _extentx        =   2514
      _extenty        =   873
      caption         =   ""
      fount           =   "frmAjusPreVen.frx":1702
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11311500
      captionalignment=   2
   End
   Begin PCGestion.bsGradientLabel lblMargen 
      Height          =   495
      Left            =   4380
      Top             =   105
      Width           =   900
      _extentx        =   979
      _extenty        =   873
      caption         =   ""
      fount           =   "frmAjusPreVen.frx":1730
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11311500
      captionalignment=   2
   End
   Begin PCGestion.miCombo cbTIPOIVA 
      Height          =   510
      Left            =   1650
      TabIndex        =   9
      Top             =   1305
      Width           =   3615
      _extentx        =   6376
      _extenty        =   900
      font            =   "frmAjusPreVen.frx":175E
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      Height          =   330
      Left            =   1155
      TabIndex        =   10
      Top             =   1395
      Width           =   420
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO COMPRA"
      Height          =   630
      Left            =   600
      TabIndex        =   8
      Top             =   45
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MARGEN %"
      Height          =   300
      Left            =   3135
      TabIndex        =   7
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PVP + IVA"
      Height          =   300
      Left            =   480
      TabIndex        =   6
      Top             =   3135
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO VENTA FINAL"
      Height          =   630
      Left            =   90
      TabIndex        =   5
      Top             =   2415
      Width           =   1515
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO VENTA EXACTO"
      Height          =   630
      Left            =   75
      TabIndex        =   4
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SUGERENCIA"
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Top             =   1950
      Width           =   1335
   End
End
Attribute VB_Name = "frmAjusPreVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmAjusPreVen
' Fecha/Hora  : 30/11/2003 13:23
' Autor       : JCASTILLO
' Propósito   : Ajustar el precio de venta conforme a la tarifa seleccionada
'               teniendo en cuenta los precios psicologicos.
'---------------------------------------------------------------------------------------

Option Explicit

'precio de Compra (original)
Public AJ_Precio_Compra As Double

'% de margen de la tarifa seleccionada
Public AJ_Margen_Tarifa As Single

'precio de venta final (despues del ajuste)
Public AJ_PreVenta_Final As Double

'si hemos salido del formulario por cancelar (no hacer cambios)
Public AJ_Cancelado As Boolean

'tipo de IVA (16, 7, 4 o 0)
Public AJ_TipoIVA As Byte

Private Sub cbAceptar_Click()
Unload Me
End Sub

Private Sub cbCancelar_Click()

'salimos por cancelar
AJ_Cancelado = True
Unload Me

End Sub



Private Sub cbTIPOIVA_Validate(Cancel As Boolean)

If cbTIPOIVA.Text = "" Then

Cancel = True
cbTIPOIVA.SetFocus

Else


AJ_TipoIVA = CByte(cbTIPOIVA.Text)
lblPVPIVA.Caption = Format(Replace(ioPREVEN.Valor, ".", ",") + (Replace(ioPREVEN.Valor, ".", ",") * devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & AJ_TipoIVA) / 100), "Currency")
End If

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With ioPREVEN
    .dspFormat = "Currency"
    .SoloNumeros = True
    .Alineacion = 1
    .PermitirBlanco = False
End With

With cbTIPOIVA
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, CAST(IVA AS char(" & .LenCodigo & ")) + ' - ' + DESCRIPCION FROM IVA WHERE MBAJA = 0 ORDER BY CODIGO"
    DoEvents
    .CodigoWidth = 500
    .carga
End With

lblMargen.Caption = AJ_Margen_Tarifa
lblPRECOM.Caption = Format(AJ_Precio_Compra, "Currency")
lblPreVenEx.Caption = Format(AJ_Precio_Compra + ((AJ_Precio_Compra * AJ_Margen_Tarifa) / 100), "Currency")
lblPreVenSug.Caption = Format(aproxima_precio(AJ_Precio_Compra + ((AJ_Precio_Compra * AJ_Margen_Tarifa) / 100)), "Currency")
ioPREVEN.Text = aproxima_precio(Format(AJ_Precio_Compra + ((AJ_Precio_Compra * AJ_Margen_Tarifa) / 100), "Currency"))
cbTIPOIVA.Text = AJ_TipoIVA
lblPVPIVA.Caption = Format(Replace(ioPREVEN.Valor, ".", ",") + (Replace(ioPREVEN.Valor, ".", ",") * devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & AJ_TipoIVA) / 100), "Currency")

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : aproxima_precio
' Fecha/Hora  : 30/11/2003 14:49
' Autor       : JCASTILLO
' Propósito   :
'
' Algoritmo de aproximacion a precios psicologicos
' Si tiene centimos:
'
'         Si son mas de 50 centimos, ajustar hacia arriba poniendo
'         ,90
'
'         Si son menos de 50 centimos, ajustar hacia abajo poniendo
'         ,99

' Si es un entero:
'
'         Restar un centimo de euro para poner a ,99
'---------------------------------------------------------------------------------------
Private Function aproxima_precio(Preven As Double) As Double

'redondear a 2 decimales (euro)
Preven = Round(Preven, 2)

'determinar si tiene decimales:
If Int(Preven) = Preven Then
    'si no tiene decimales:

    'redondear hacia abajo (restar 0.01)
    aproxima_precio = Int(Preven) - 0.01

Else 'si tiene decimales:

 'obtener la parte decimal. Si son >= 50 cens ...
 If Preven - Int(Preven) >= 50 Then
 
     'ponerle .90 centimos (redondear hacia arriba)
     aproxima_precio = Int(Preven) + 0.9
      
 Else
 
    'redondear hacia abajo (una unidad menos y ponerle 0.99)
     aproxima_precio = Int(Preven) - 0.01
 
 End If
 
 
End If
 
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not AJ_Cancelado Then AJ_PreVenta_Final = ioPREVEN.Valor

End Sub
