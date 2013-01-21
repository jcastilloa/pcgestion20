VERSION 5.00
Begin VB.Form frmImpEtiq 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir etiquetas"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3360
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
   ScaleHeight     =   2220
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin PCGestion.miText ioDIGITOS 
      Height          =   525
      Left            =   1845
      TabIndex        =   1
      Top             =   660
      Width           =   810
      _extentx        =   1429
      _extenty        =   926
      font            =   "frmImpEtiq.frx":0000
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   1260
      TabIndex        =   2
      Top             =   1395
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
      MICON           =   "frmImpEtiq.frx":002C
      PICN            =   "frmImpEtiq.frx":0048
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioSALTAR 
      Height          =   450
      Left            =   1830
      TabIndex        =   0
      Top             =   120
      Width           =   825
      _extentx        =   1455
      _extenty        =   794
      font            =   "frmImpEtiq.frx":0D22
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saltar etiquetas"
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   1680
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dos dígitos a imprimir"
      Height          =   555
      Left            =   345
      TabIndex        =   3
      Top             =   555
      Width           =   1455
   End
End
Attribute VB_Name = "frmImpEtiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmImpEtiq
' Fecha/Hora  : 06/01/2004 15:12
' Autor       : JCASTILLO
' Propósito   : Imprimir etiquetas. Recoge los valores XX (dos digitos para disimular
'               el precio de venta. Y el numero de etiquetas a saltar.
'---------------------------------------------------------------------------------------
Option Explicit

'digitos a imprimir
Public digitos As String
'saltar n etiquetas
Public saltar As Byte

Private Sub cbAceptar_Click()

If Len(ioDIGITOS.Text) <> 2 Then
    MsgBox "DIGITOS incorrecto", vbInformation, titulo
    ioDIGITOS.SetFocus
    ioDIGITOS.CancelarValidacion
    Exit Sub
End If

If ioSALTAR.Text <> "" Then
    If CLng(ioSALTAR.Text) > 254 Then
        MsgBox "Numero de etiquetas a saltar incorrecto", vbExclamation, titulo
        ioSALTAR.SetFocus
        ioSALTAR.CancelarValidacion
    End If
    saltar = ioSALTAR.Text
    
End If

digitos = ioDIGITOS.Text


Unload Me

End Sub

Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With ioDIGITOS
    .PermitirBlanco = False
    .SoloNumeros = True
    .LongMaxima = 2
    .Text = "11"
End With

With ioSALTAR
    .PermitirBlanco = True
    .SoloNumeros = True
    .LongMaxima = 3
End With

End Sub

