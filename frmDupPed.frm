VERSION 5.00
Begin VB.Form frmDupPed 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Duplicar Pedido ..."
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7125
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
   ScaleHeight     =   2790
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin PCGestion.bsGradientLabel lblIVA 
      Height          =   465
      Left            =   45
      Top             =   870
      Width           =   2085
      _ExtentX        =   2778
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
   Begin PCGestion.miText ioTOTAL_CON_IVA 
      Height          =   555
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   893
      TabIndex        =   1
      Top             =   1935
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
      MICON           =   "frmDupPed.frx":0000
      PICN            =   "frmDupPed.frx":001C
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
      Left            =   1868
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1935
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
      MICON           =   "frmDupPed.frx":0CF6
      PICN            =   "frmDupPed.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      Height          =   285
      Left            =   15
      TabIndex        =   2
      Top             =   315
      Width           =   765
   End
End
Attribute VB_Name = "frmDupPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo         :  frmCalcIVA
' Fecha/Hora   :   03/12/2003 12:54
' Autor           :   JCastillo
' Propósito     :   Calcular el %IVA que se le aplicacan a los articulos del pedido
'                       cuando el proveedor envia los articulos sin el IVA especificado
'                       para cada articulo, solo en el TOTAL GENERAL
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
Public Total_General_Con_IVA As Double
'---------------------------------------------------------------------------------------
'Despues de introducir sus unidades y precios sin IVA
Public Total_General_Sin_IVA As Double
'Porcentaje de IVA aplicado
Public Porcen_IVA As Byte
'---------------------------------------------------------------------------------------


Private Sub cbAceptar_Click()

If MsgBox("El porcentaje de IVA ha cambiado, ¿desea aplicarlo a todos los artículos de este pedido?", vbQuestion + vbYesNo) = vbNo Then
    Porcen_IVA = 0
End If

Unload Me

End Sub

Private Sub cbCancelar_Click()

Porcen_IVA = 0
Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With ioTOTAL_CON_IVA
    .PermitirBlanco = False
    .SoloNumeros = True
    .dspFormat = "Currency"
    .Alineacion = 1
End With

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : ioTOTAL_CON_IVA_Validate
' Fecha/Hora     : 03/12/2003 13:15
' Autor             : JCastillo
' Propósito        : Validación y cálculo del % de iva aplicado
'---------------------------------------------------------------------------------------
Private Sub ioTOTAL_CON_IVA_Validate(Cancel As Boolean)
Dim tmpimp_iva As Double

   On Error GoTo ioTOTAL_CON_IVA_Validate_Error

    With ioTOTAL_CON_IVA
    
        'no puede ser menor que el total sin IVA
        If .Valor < Total_General_Sin_IVA Then
            MsgBox "El total es incorrecto", vbInformation
            .SetFocus
            .CancelarValidacion
            Cancel = True
            Porcen_IVA = 0
            Exit Sub
        End If
        
        'sacar el importe de IVA en €
        tmpimp_iva = .Valor - Total_General_Sin_IVA
        
        'sacar el % del IVA
        If Round(tmpimp_iva * 100 / Total_General_Sin_IVA) > 100 Then
        'porcentaje no valido
            MsgBox "El % de IVA no es valido (mayor de 100%)", vbInformation
            .SetFocus
            .CancelarValidacion
            Cancel = True
            Porcen_IVA = 0
            Exit Sub
        Else
        
            Porcen_IVA = Round(tmpimp_iva * 100 / Total_General_Sin_IVA)
        
        End If
        
        lblIVA.Caption = CStr(Porcen_IVA) & " %"
        
        
    End With

   On Error GoTo 0
   Exit Sub

ioTOTAL_CON_IVA_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioTOTAL_CON_IVA_Validate de Formulario frmCalcIVA"

End Sub
