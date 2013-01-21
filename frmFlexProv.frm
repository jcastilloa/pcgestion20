VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFlexProv 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
   ClipControls    =   0   'False
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
   ScaleHeight     =   6750
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miCombo cbSECTOR 
      Height          =   495
      Left            =   3615
      TabIndex        =   3
      Top             =   495
      Width           =   3660
      _extentx        =   6456
      _extenty        =   873
      font            =   "frmFlexProv.frx":0000
   End
   Begin PCGestion.miText ioCODIGO 
      Height          =   450
      Left            =   1050
      TabIndex        =   0
      Top             =   15
      Width           =   1050
      _extentx        =   1852
      _extenty        =   794
      font            =   "frmFlexProv.frx":002C
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5700
      Left            =   0
      TabIndex        =   5
      Top             =   1005
      Visible         =   0   'False
      Width           =   11145
      _cx             =   19659
      _cy             =   10054
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14331047
      ForeColor       =   -2147483640
      BackColorFixed  =   15120763
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14859077
      ForeColorSel    =   -2147483635
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15573900
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFlexProv.frx":0058
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   3
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin PCGestion.miText ioCIF 
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   495
      Width           =   1725
      _extentx        =   3043
      _extenty        =   794
      font            =   "frmFlexProv.frx":0136
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.miText ioNOMBRE 
      Height          =   450
      Left            =   3615
      TabIndex        =   1
      Top             =   15
      Width           =   5745
      _extentx        =   10134
      _extenty        =   794
      font            =   "frmFlexProv.frx":0162
      dspformat       =   ""
      enabled         =   -1
      espassword      =   -1
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   405
      Left            =   7290
      TabIndex        =   10
      Top             =   540
      Width           =   1260
      _extentx        =   2223
      _extenty        =   714
      caption         =   "&Consultar"
      font            =   "frmFlexProv.frx":018E
      image           =   "frmFlexProv.frx":01BA
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   405
      Left            =   8580
      TabIndex        =   11
      Top             =   540
      Width           =   750
      _extentx        =   1323
      _extenty        =   714
      caption         =   "&Borrar"
      font            =   "frmFlexProv.frx":01D8
      image           =   "frmFlexProv.frx":0204
   End
   Begin PCGestion.chameleonButton chameleonButton1 
      Height          =   555
      Left            =   9900
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   735
      _extentx        =   1296
      _extenty        =   979
      btype           =   9
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmFlexProv.frx":0222
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmFlexProv.frx":024E
      picn            =   "frmFlexProv.frx":026C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin MSForms.CheckBox fwbajas 
      Height          =   435
      Left            =   9435
      TabIndex        =   9
      Top             =   600
      Width           =   1755
      VariousPropertyBits=   746588183
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3096;767"
      Value           =   "0"
      Caption         =   "Ocultar BAJAS"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SECTOR"
      Height          =   330
      Left            =   2760
      TabIndex        =   4
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   165
      TabIndex        =   8
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   300
      Left            =   2745
      TabIndex        =   7
      Top             =   105
      Width           =   840
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CIF"
      Height          =   330
      Left            =   630
      TabIndex        =   6
      Top             =   570
      Width           =   360
   End
End
Attribute VB_Name = "frmFlexProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim first As Boolean

Dim tmprc As New ADODB.Recordset
Dim tmpstrcombo As String

Public miosql As New clsSmartSQL
Public miRc As New ADODB.Recordset
Dim nif As New clsNIF

Private Sub cbBorrar_click()

ioCODIGO.Text = ""
ioCIF.Text = ""
ioNOMBRE.Text = ""
cbSECTOR.Text = ""
fwbajas.Value = True

Call cbLista_click

End Sub

Private Sub cbLista_click()

miosql.ClearWhereClause

If ioCODIGO.Text <> "" Then
    miosql.AddSimpleWhereClause "CODIGO", CLng(ioCODIGO.Text)
End If

If ioNOMBRE.Text <> "" Then
    miosql.AddSimpleWhereClause "NOMBRE", ioNOMBRE.Text, , CLAUSE_LIKE
End If

If ioCIF.Text <> "" Then
    miosql.AddSimpleWhereClause "CIF", ioCIF.Text
End If

If cbSECTOR.Text <> "" Then
    miosql.AddSimpleWhereClause "SECTOR", CLng(cbSECTOR.Text)
End If

'si decimos que ocultar bajas,
'MBAJA = FALSE
If fwbajas.Value = True Then
    miosql.AddSimpleWhereClause "MBAJA", 0
End If

miRc.Close

If locCnn.State = 1 Then
    miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic
ElseIf locCnnSP.State = 1 Then
    miRc.Open miosql.SQL, locCnnSP, adOpenStatic, adLockOptimistic
End If

Set fg.DataSource = miRc
DoEvents

'ocultar la columna rowguid
fg.ColHidden(fg.Cols - 1) = True
        
fg.ColComboList(4) = tmpstrcombo
fg.ColFormat(1) = "0000"
fg.AutoSize 1, fg.Cols - 1

End Sub

Private Sub cbSECTOR_Validate(Cancel As Boolean)
    Call cbLista_click
End Sub

Private Sub chameleonButton1_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpalm As String
         

   On Error GoTo chameleonButton1_Click_Error

    DoEvents

   
    linea1 = "Proveedores.  Nomobre: " & ioNOMBRE.Text & ". CIF: " & ioCIF.Text
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexProv"

End Sub

Private Sub fg_dblClick()
    Unload Me
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13, vbKeyEscape
    KeyAscii = 0
    Unload Me
    
End Select

End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    DoEvents
       
    
    If Not first Then
    
        With ioCODIGO
            .SoloNumeros = True
            .LongMaxima = 4
            .dspFormat = "0000"
            .Alineacion = 1
        End With
                
               
        Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortShow
        
  
  If locCnn.State = 1 Then
    tmprc.Open "SELECT CODST, SECTOR FROM SECTORES WHERE MBAJA = 0 ORDER BY CODST", locCnn, adOpenDynamic, adLockReadOnly
  ElseIf locCnnSP.State = 1 Then
    tmprc.Open "SELECT CODST, SECTOR FROM SECTORES WHERE MBAJA = 0 ORDER BY CODST", locCnnSP, adOpenDynamic, adLockReadOnly
  End If
  
    tmpstrcombo = fg.BuildComboList(tmprc, "SECTOR", "CODST", vbBlue)
    fg.ColComboList(4) = tmpstrcombo
    fg.ColFormat(1) = "0000"
    fg.AutoSize 1, fg.Cols - 1
    
    tmprc.Close
    Set tmprc = Nothing
    

        first = True
    End If
       
    
End Sub


Private Sub Form_Load()

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
  fg.Visible = False

          'Cargar el micombo sectores
  With cbSECTOR
    .ConexionString = locCnn
    .SQLString = "SELECT CODST, SECTOR FROM SECTORES WHERE MBAJA = 0 ORDER BY CODST"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmpstrcombo = ""
    Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    If miRc.EOF Then Call cbBorrar_click
    
    Set frmFlexProv = Nothing
End Sub

Private Sub ioCIF_Validate(Cancel As Boolean)

'si esta a blancos salir
If Trim(ioCIF.Text) = "" Then
   ' ioCIF.CancelarValidacion
   ' Cancel = True
    Exit Sub
End If

nif.DarFormato = True
nif.nif = ioCIF.Text

If nif.Err Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioCIF.Text = nif.nif
End If


Call cbLista_click

'If ioCIF.Text <> "" Then Call comprueba_DNI(ioCIF.Text, ioCIF)
End Sub



Private Sub ioNOMBRE_Validate(Cancel As Boolean)
Call cbLista_click
End Sub
