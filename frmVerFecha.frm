VERSION 5.00
Begin VB.Form frmVerFecha 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fecha en curso ..."
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6420
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
   ScaleHeight     =   3105
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PCGestion.bsGradientLabel lblFecha 
      Height          =   900
      Left            =   45
      Top             =   360
      Width           =   6315
      _extentx        =   11139
      _extenty        =   1588
      caption         =   ""
      fount           =   "frmVerFecha.frx":0000
      captioncolour   =   16711680
      colour1         =   14737632
      colour2         =   12632256
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   2700
      TabIndex        =   0
      Top             =   2250
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmVerFecha.frx":002E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmVerFecha.frx":005A
      picn            =   "frmVerFecha.frx":0078
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblFecha2 
      Height          =   330
      Left            =   45
      Top             =   1290
      Width           =   6315
      _extentx        =   11139
      _extenty        =   582
      caption         =   ""
      fount           =   "frmVerFecha.frx":0D54
      captioncolour   =   16711680
      colour1         =   14737632
      colour2         =   12632256
      captionalignment=   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Si la fecha del sistema no fuera correcta, cámbiela, cierre y vuelva a abrir PCGestion."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   60
      TabIndex        =   2
      Top             =   1620
      Width           =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "La fecha en curso del sistema es:"
      Height          =   360
      Left            =   75
      TabIndex        =   1
      Top             =   0
      Width           =   6240
   End
End
Attribute VB_Name = "frmVerFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmVerFecha
' Fecha/Hora  : 19/09/2004 20:16
' Autor       : JCASTILLO
' Propósito   : Ver la fecha en curso para que el usuario se de cuenta
'               de posibles cambios de fecha.
'---------------------------------------------------------------------------------------

Option Explicit


Private Sub cbAceptar_Click()

Unload Me

End Sub
Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
lblFecha.Caption = Format(Now, "dddddd")
lblFecha2.Caption = Format(Now, "dd/mm/yyyy hh:mm:ss")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set frmVerFecha = Nothing

End Sub
