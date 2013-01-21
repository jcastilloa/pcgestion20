VERSION 5.00
Begin VB.Form frmCalcIVA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calcular % IVA desde TOTAL"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
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
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin PCGestion.bsGradientLabel lblIVA 
      Height          =   465
      Left            =   1575
      Top             =   1335
      Width           =   2085
      _extentx        =   2778
      _extenty        =   820
      caption         =   ""
      fount           =   "frmCalcIVA.frx":0000
      captioncolour   =   192
      colour1         =   14737632
      colour2         =   12632256
      captionalignment=   2
   End
   Begin PCGestion.bsGradientLabel lblTot_sin_IVA 
      Height          =   465
      Left            =   1575
      Top             =   75
      Width           =   2085
      _extentx        =   2778
      _extenty        =   820
      caption         =   ""
      fount           =   "frmCalcIVA.frx":002E
      captioncolour   =   16711680
      colour1         =   14737632
      colour2         =   12632256
      captionalignment=   2
   End
   Begin PCGestion.miText ioTOTAL_CON_IVA 
      Height          =   525
      Left            =   1560
      TabIndex        =   0
      Top             =   705
      Width           =   2130
      _extentx        =   3757
      _extenty        =   926
      font            =   "frmCalcIVA.frx":005C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   893
      TabIndex        =   1
      Top             =   1935
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmCalcIVA.frx":0088
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCalcIVA.frx":00B4
      picn            =   "frmCalcIVA.frx":00D2
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
      Left            =   1868
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1935
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmCalcIVA.frx":0DAE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCalcIVA.frx":0DDA
      picn            =   "frmCalcIVA.frx":0DF8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "% IVA aplicado"
      Height          =   315
      Left            =   45
      TabIndex        =   4
      Top             =   1410
      Width           =   1470
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total general CON IVA"
      Height          =   555
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total SIN IVA"
      Height          =   345
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmCalcIVA"
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

'cargar total sin IVA
lblTot_sin_IVA.Caption = Format(Total_General_Sin_IVA, "Currency")

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
