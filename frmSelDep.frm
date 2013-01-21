VERSION 5.00
Begin VB.Form frmSelDep 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione Dependiente"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7230
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
   ScaleHeight     =   3945
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDependientes 
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7200
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   5310
      TabIndex        =   1
      Top             =   3120
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmSelDep.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmSelDep.frx":002C
      picn            =   "frmSelDep.frx":004A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   345
      Left            =   30
      Top             =   2760
      Width           =   7185
      _extentx        =   12674
      _extenty        =   609
      caption         =   ""
      fount           =   "frmSelDep.frx":0D26
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6285
      TabIndex        =   2
      Top             =   3120
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmSelDep.frx":0D54
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmSelDep.frx":0D80
      picn            =   "frmSelDep.frx":0D9E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   345
      Left            =   0
      Top             =   15315
      Width           =   7185
      _extentx        =   12674
      _extenty        =   609
      caption         =   ""
      fount           =   "frmSelDep.frx":167A
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   345
      Left            =   2445
      Top             =   3570
      Width           =   2835
      _extentx        =   5001
      _extenty        =   609
      caption         =   "- C -  Cancelar"
      fount           =   "frmSelDep.frx":16A8
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
End
Attribute VB_Name = "frmSelDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmSelDep
' Fecha/Hora  : 17/01/2004 21:41
' Autor       : JCASTILLO
' Propósito   : Selecciona un dependiente para la nueva venta
'---------------------------------------------------------------------------------------
Option Explicit

Public ID_Dependiente As Integer
Public Nombre_Dep As String
Public S_Cancelado As Boolean

Dim rc As New ADODB.Recordset


Private Sub cbAceptar_Click()

If ID_Dependiente = 0 Then
    lblStatus.Caption = "Debe seleccionar un Dependiente"
    Exit Sub
End If

S_Cancelado = False

Unload Me

End Sub


Private Sub cbCancelar_Click()

If rc.State = 1 Then rc.Close
Set rc = Nothing

S_Cancelado = True
Unload Me

End Sub

Private Sub Form_Activate()
lstDependientes.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'cancelar
If KeyCode = vbKeyC Then
    Call cbCancelar_Click
    KeyCode = 0
End If


End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

ID_Dependiente = 0

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

'cargar el list de dependientes del centro actual
rc.Open "SELECT CODIGO, NOMBRE from PERSONAL WHERE (CODCEN = " & CentroActual & ") AND (TIPPERM = 0) AND (MBAJA = 0)", locCnn, adOpenDynamic

With lstDependientes

Do

    .AddItem Format(rc.fields("CODIGO"), "00000") & " - " & Trim(rc.fields("NOMBRE"))
    If Not rc.EOF Then rc.MoveNext
    
Loop Until rc.EOF

    
    If .ListCount > 0 Then .ListIndex = 0
    DoEvents
   
    
End With


rc.Close
'cargar el list de con los supervisores de la aplicación
rc.Open "SELECT CODIGO, NOMBRE from PERSONAL WHERE (TIPPERM = 1) AND (MBAJA = 0)", locCnn, adOpenDynamic

With lstDependientes

Do

    .AddItem Format(rc.fields("CODIGO"), "00000") & " - " & Trim(rc.fields("NOMBRE"))
    If Not rc.EOF Then rc.MoveNext
    
        
Loop Until rc.EOF

    
    If .ListCount > 0 Then .ListIndex = 0
    DoEvents
   
    
End With

rc.Close
Set rc = Nothing


   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    If rc.State = 1 Then rc.Close
    Set rc = Nothing
    If Err.Number = 3705 Then Unload Me
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario frmSelDep"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If rc.State = 1 Then rc.Close
Set rc = Nothing

End Sub

Private Sub lstDependientes_DblClick()

    If lstDependientes.Text = "" Then Exit Sub
    'obtener el id dependiente seleccionado
    ID_Dependiente = Left(lstDependientes.Text, 5)
    Nombre_Dep = Mid(lstDependientes.Text, 9, Len(lstDependientes.Text) - 8)
    
    Call cbAceptar_Click

End Sub

Private Sub lstDependientes_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
    If lstDependientes.Text = "" Then Exit Sub
    'obtener el id dependiente seleccionado
    ID_Dependiente = Left(lstDependientes.Text, 5)
    Nombre_Dep = Mid(lstDependientes.Text, 9, Len(lstDependientes.Text) - 8)
    
    Call cbAceptar_Click

End If

End Sub
