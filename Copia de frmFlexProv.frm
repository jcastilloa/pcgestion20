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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miCombo cbSECTOR 
      Height          =   495
      Left            =   3615
      TabIndex        =   3
      Top             =   495
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miText ioCODIGO 
      Height          =   450
      Left            =   1050
      TabIndex        =   0
      Top             =   15
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5700
      Left            =   0
      TabIndex        =   5
      Top             =   1035
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
      FormatString    =   $"frmFlexProv.frx":0000
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
      _ExtentX        =   3043
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
   End
   Begin PCGestion.miText ioNOMBRE 
      Height          =   450
      Left            =   3615
      TabIndex        =   1
      Top             =   15
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   405
      Left            =   7290
      TabIndex        =   10
      Top             =   540
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "&Consultar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexProv.frx":00DE
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   405
      Left            =   8580
      TabIndex        =   11
      Top             =   540
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   714
      Caption         =   "&Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexProv.frx":00FA
   End
   Begin MSForms.CheckBox fwbajas 
      Height          =   435
      Left            =   9435
      TabIndex        =   9
      Top             =   555
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

Public miOsql As New clsSmartSQL
Public miRc As New ADODB.Recordset
Dim Nif As New clsNIF

Private Sub cbBorrar_click()

ioCODIGO.Text = ""
ioCIF.Text = ""
ioNOMBRE.Text = ""
cbSECTOR.Text = ""
fwbajas.Value = True

Call cbLista_click

End Sub

Private Sub cbLista_click()

miOsql.ClearWhereClause

If ioCODIGO.Text <> "" Then
    miOsql.AddSimpleWhereClause "CODIGO", CLng(ioCODIGO.Text)
End If

If ioNOMBRE.Text <> "" Then
    miOsql.AddSimpleWhereClause "NOMBRE", ioNOMBRE.Text, , CLAUSE_LIKE
End If

If ioCIF.Text <> "" Then
    miOsql.AddSimpleWhereClause "CIF", ioCIF.Text
End If

If cbSECTOR.Text <> "" Then
    miOsql.AddSimpleWhereClause "SECTOR", CLng(cbSECTOR.Text)
End If

'si decimos que ocultar bajas,
'MBAJA = FALSE
If fwbajas.Value = True Then
    miOsql.AddSimpleWhereClause "MBAJA", 0
End If

miRc.Close
miRc.Open miOsql.SQL, locCnn, adOpenStatic, adLockOptimistic

Set fg.DataSource = miRc
DoEvents
fg.ColComboList(4) = tmpstrcombo
fg.ColFormat(1) = "0000"
fg.AutoSize 1, fg.Cols - 1

End Sub

Private Sub fg_DblClick()
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
        End With
                
               
        Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortShow
        

  
  tmprc.Open "SELECT CODST, SECTOR FROM SECTORES WHERE MBAJA = 0 ORDER BY CODST", locCnn, adOpenDynamic, adLockReadOnly
  
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
    Set Nif = Nothing
    Set frmFlexProv = Nothing
    
End Sub

Private Sub ioCIF_Validate(Cancel As Boolean)

'si esta a blancos salir
If Trim(ioCIF.Text) = "" Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
End If

Nif.DarFormato = True
Nif.Nif = ioCIF.Text

If Nif.Err Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioCIF.Text = Nif.Nif
End If

'If ioCIF.Text <> "" Then Call comprueba_DNI(ioCIF.Text, ioCIF)
End Sub

