VERSION 5.00
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#2.1#0"; "fpFlp20.ocx"
Begin VB.UserControl miCombo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   6990
   ToolboxBitmap   =   "miCombo.ctx":0000
   Begin LpADOLib.fpComboADO cbDESC 
      Height          =   390
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   5625
      _Version        =   131073
      _ExtentX        =   9922
      _ExtentY        =   688
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   2
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   14737632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ColDesigner     =   "miCombo.ctx":0312
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   -1
      ComboGap        =   7
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   -1  'True
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
      ExtendRow       =   0
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1245
   End
End
Attribute VB_Name = "miCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' Module    : miCombo
' DateTime  : 09/06/2003 12:14
' Author    : José Castillo
' Purpose   : combo multiuso
'---------------------------------------------------------------------------------------

Option Explicit

Const fondo_original = vbWhite
Dim fondo_celda As Long
Dim abierto As Boolean

'Const LenCodigo = 8
'Const ConexionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=SRVDB;UID=sysdba;PWD=masterkey;DB=100.100.100.161:/mnt/disco2/estadistica/contratos.gdb;CHARSET=ISO8859_1;"
'Const SQLString = "select DNIPER || '   ' ||  NOMBRE || '   ' || APELL1 || '   ' || APELL2 from cardnif"
'Default Property Values:
Const No_Existe_Text = "No Existe ..."
Const m_def_BackColor = vbBlue
Const m_def_ForeColor = vbBlack
Const m_def_Enabled = True
Const m_def_Appearance = 1
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Text = "00000000"
Const m_def_LenCodigo = 8
Const m_def_ConexionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=SRVDB;UID=sysdba;PWD=masterkey;DB=100.100.100.161:/mnt/disco2/estadistica/contratos.gdb;CHARSET=ISO8859_1;"
Const m_def_SQLString = "select DNIPER,  NOMBRE || '   ' || APELL1 || '   ' || APELL2 from cardnif"
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_enabled As Boolean
Dim m_Font As Font
Dim m_Appearance As Integer
Dim m_Text As String
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_LenCodigo As Integer
Dim m_ConexionString As String
Dim m_SQLString As String
Dim m_Locked As Boolean

Dim m_DataSource As ADODB.Recordset
Dim m_DataField As String

'Event Declarations:
Event Click()
'Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private m_CodigoWidth As Single

Const m_def_DataSource = Null
Const m_def_DataField = ""

'Para que se despliegue:
'Call SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, 0&)
'Para que se contraiga:
'Call SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 0, 0&)

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Se posiciona en un fpComboADO buscando el texto
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function posiciona_combo(texto As String, com As fpComboADO) As Boolean
    Dim cB As Long
    Dim FindString As String
    Const CB_ERR = (-1)
    Const CB_FINDSTRING = &H14C
    
    cB = SendMessage(com.hwnd, CB_FINDSTRING, -1, ByVal texto)
    
    If cB <> CB_ERR Then
        com.ListIndex = cB
        'com.SelStart = Len(texto)
        'com.SelLength = Len(com.Text) - com.SelStart
        Else
        'si no se encuentra, devolver true para poder cancelar
        'la validacion
        
        'si esta a blancos, salir
        If txtCodigo.Text = "" Then Exit Function
        
        cbDESC.Text = No_Existe_Text
        
        With txtCodigo
            .SelStart = 0
            .SelLength = Len(txtCodigo.Text)
        End With
        
        posiciona_combo = True
    End If

End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub desplegar_combo(com As fpComboADO)
    Call SendMessage(com.hwnd, &H14F, 1, 0&)
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub contraer_combo(com As fpComboADO)
    Call SendMessage(com.hwnd, &H14F, 0, 0&)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
    BackColor = m_BackColor
    
    txtCodigo.BackColor = BackColor
    cbDESC.BackColor = BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
    txtCodigo.ForeColor = ForeColor
    cbDESC.ForeColor = ForeColor
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Locked() As Boolean
    Locked = m_Locked
    'txtCodigo.Locked = Locked
    'cbDESC.Locked = Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    txtCodigo.Locked = Locked
    cbDESC.Enabled = Not Locked
    'cbDESC.Locked = Locked
    PropertyChanged "Locked"
End Property


Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_enabled
    'txtCodigo.Enabled = Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_enabled = New_Enabled
    PropertyChanged "Enabled"
    txtCodigo.Enabled = Enabled
    cbDESC.Enabled = Enabled
    
    
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=6,0,0,0
Public Property Get Font() As Font

    Set Font = m_Font
    Set txtCodigo.Font = Font
    Set cbDESC.Font = Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,1
Public Property Get Appearance() As Integer
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get LenCodigo() As Integer
    LenCodigo = m_LenCodigo
    txtCodigo.MaxLength = LenCodigo
End Property

Public Property Let LenCodigo(ByVal New_LenCodigo As Integer)
    m_LenCodigo = New_LenCodigo
    PropertyChanged "LenCodigo"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ConexionString() As String
    ConexionString = m_ConexionString
End Property

Public Property Let ConexionString(ByVal New_ConexionString As String)
    m_ConexionString = New_ConexionString
    PropertyChanged "ConexionString"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get SQLString() As Variant
    SQLString = m_SQLString
End Property

Public Property Let SQLString(ByVal New_SQLString As Variant)
    m_SQLString = New_SQLString
    PropertyChanged "SQLString"
End Property

'Recoger el codigo q pone el usuario
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
'Public Property Get Codigo() As Variant
     'LenCodigo = m_LenCodigo
   'Me.Codigo = m_Codigo
   'm_Codigo = Me.Codigo
'   Codigo = m_Codigo
   'txtCodigo.Text = Codigo
'End Property

Private Sub cbDESC_Click()

On Error Resume Next

If txtCodigo.Locked Then Exit Sub

        With cbDESC
        
        If Not abierto And .Text <> No_Existe_Text Then
            txtCodigo.Text = Trim(Left(.Text, LenCodigo))
            txtCodigo.TabStop = True
'           txtCodigo.SetFocus
            .TabStop = False
        ElseIf Not abierto And .Text = No_Existe_Text Then
            txtCodigo.Text = ""
            txtCodigo.TabStop = True
            txtCodigo.SetFocus
            .TabStop = False
            abierto = False
        End If
        
        End With
End Sub

Private Sub cbDESC_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode

Case 13
    'Call desplegar_combo(cbDESC)
    With cbDESC
    If .Text <> No_Existe_Text Then
        txtCodigo.Text = Trim(Left(.Text, LenCodigo))
        txtCodigo.TabStop = True
        txtCodigo.SetFocus
        .TabStop = False
        abierto = False
    Else
        txtCodigo.Text = ""
        txtCodigo.TabStop = True
        txtCodigo.SetFocus
        .TabStop = False
        abierto = False
    End If
    
    End With
    
End Select
End Sub




Private Sub txtCodigo_GotFocus()
    
    If m_DataField = "XCENTRO" Then
    Debug.Print "p"
    End If
    
    txtCodigo.BackColor = &HFFFF80
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyUp
        
   ' SendKeys "+{TAB}"
    
    FocusNextMask Up
    
   ' Case vbKeyDown
   ' FocusNextMask Down
    

Case vbKeyDown
                
                If txtCodigo.Locked Then Exit Sub
                abierto = True
                cbDESC.TabStop = True
                cbDESC.SetFocus
                txtCodigo.TabStop = False
                Call desplegar_combo(cbDESC)
                KeyCode = 0

End Select


End Sub

Private Sub txtCodigo_lostFocus()
    txtCodigo.BackColor = fondo_original
    If txtCodigo.Text <> "" Then posiciona_combo txtCodigo.Text, cbDESC
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Dim tmpformato As String
    
    Select Case KeyAscii
    Case vbKeyBack
                Exit Sub
    Case vbKeySpace
                
                cbDESC.TabStop = True
                'cbDESC.SetFocus
                SendKeys "{tab}"
                Call desplegar_combo(cbDESC)
                KeyAscii = 0
                Exit Sub
                
    Case 13
    
                                    
                If m_DataField = "" Then
                    tmpformato = String(LenCodigo, "0")
                    txtCodigo.Text = Format(txtCodigo, tmpformato)
                    posiciona_combo txtCodigo.Text, cbDESC
                End If
                
                    tmpformato = ""
                    cbDESC.TabStop = False
                    DoEvents
                    SendKeys "{tab}"
                    Exit Sub
                    
            
     End Select
     
    'If Len(Trim(txtCodigo)) = (LenCodigo) Then
    '    SendKeys "{tab}"
     '   KeyAscii = 0
    '    posiciona_combo txtCodigo.Text, cbDESC
    '    Exit Sub
   ' End If
 
End Sub

'Ponerle el codigo al usuario
Public Property Get Text() As Variant
                Text = txtCodigo.Text
End Property

'Obtener el codigo del usuario
Public Property Let Text(ByVal iCodigo As Variant)
Dim tmpformato As String
    tmpformato = String(LenCodigo, "0")
    
    'si es distinto de blanco
    If Trim(iCodigo) <> "" And Trim(iCodigo) <> "0" Then
    
    
    If Not IsNumeric(iCodigo) Then Exit Property
    
    'si el codigo es mayor de 0
    If CLng(iCodigo) > 0 Then
        
        If m_DataField = "" Then
            txtCodigo.Text = Format(iCodigo, tmpformato)
        Else
            txtCodigo.Text = iCodigo
        End If
        posiciona_combo txtCodigo.Text, cbDESC
    Else
        cbDESC.ListIndex = -1
    End If
    
    'si es = ""
    Else
        cbDESC.ListIndex = -1
        txtCodigo.Text = ""
    End If
    
    tmpformato = ""
End Property

Public Property Get CodigoWidth() As Single
    CodigoWidth = m_CodigoWidth
End Property

Public Property Let CodigoWidth(ByVal fCodigoWidth As Single)
    m_CodigoWidth = fCodigoWidth
    txtCodigo.Width = fCodigoWidth
    Call UserControl_Resize
    Call UserControl.PropertyChanged("CodigoWidth")
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get DataSource() As ADODB.Recordset
   Set DataSource = m_DataSource
   'Set txtCodigo.DataSource = m_DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As ADODB.Recordset)
    Set m_DataSource = New_DataSource
    Set txtCodigo.DataSource = m_DataSource
    PropertyChanged "DataSource"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get DataField() As String
    DataField = m_DataField
    'txtCodigo.DataField = DataField
End Property

Public Property Let DataField(ByVal New_DataField As String)
    m_DataField = New_DataField
    PropertyChanged "DataField"
    txtCodigo.DataField = DataField
End Property

Private Sub txtCodigo_change()
    
    If m_DataField <> "" Then
        'para ignorar los valores 0 y "" recibidos por el
        'datasource (no se controlan en la propiedad
        'Let de Text).
        If txtCodigo.Text = "0" Or txtCodigo.Text = "" Then
                txtCodigo.Text = ""
                cbDESC.ListIndex = -1
        End If
    
     If Not m_DataSource Is Nothing Then
        If Not m_DataSource.EOF And Not m_DataSource.BOF Then
            If (cbDESC.Enabled = False Or txtCodigo.Enabled = False) Or _
                m_DataSource.EditMode = 0 Then _
                Call posiciona_combo(txtCodigo.Text, cbDESC)
            End If
     End If
     
    End If
    
  If Len(Trim(txtCodigo.Text)) = m_LenCodigo Then posiciona_combo txtCodigo.Text, cbDESC
  
End Sub

Private Sub txtCodigo_Validate(Cancel As Boolean)
    
    If Trim(txtCodigo.Text) = "" Then
       Cancel = False
       cbDESC.Text = ""
       DoEvents
    ElseIf posiciona_combo(txtCodigo.Text, cbDESC) = True Then
        If txtCodigo.Locked = False Then Cancel = True
    End If
    
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_enabled = m_def_Enabled
    m_Text = m_def_Text
    Set m_Font = Ambient.Font
    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_LenCodigo = m_def_LenCodigo
    m_ConexionString = m_def_ConexionString
    m_SQLString = m_def_SQLString
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Locked = PropBag.ReadProperty("Locked", False)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_LenCodigo = PropBag.ReadProperty("LenCodigo", m_def_LenCodigo)
    m_ConexionString = PropBag.ReadProperty("ConexionString", m_def_ConexionString)
    m_SQLString = PropBag.ReadProperty("SQLString", m_def_SQLString)
    
    If Not IsNull(m_def_DataSource) Then _
    Set m_DataSource = PropBag.ReadProperty("DataSource", m_def_DataSource)
    
    If m_def_DataField <> "" Then _
    m_DataField = PropBag.ReadProperty("DataField", m_def_DataSource)
    
    
End Sub

Private Sub UserControl_Resize()
cbDESC.Left = txtCodigo.Width + 50
If (UserControl.Width - (txtCodigo.Width + 50)) > 0 Then
    cbDESC.Width = UserControl.Width - (txtCodigo.Width + 50)
End If
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("LenCodigo", m_LenCodigo, m_def_LenCodigo)
    Call PropBag.WriteProperty("ConexionString", m_ConexionString, m_def_ConexionString)
    Call PropBag.WriteProperty("SQLString", m_SQLString, m_def_SQLString)
    
    Call PropBag.WriteProperty("DataSource", m_DataSource, m_def_DataSource)
    Call PropBag.WriteProperty("DataField", m_DataField, m_def_DataField)
End Sub


'Añadir ITEMS Manualmente
Public Function añade_item(ItemString As String, Optional Indice As Long)
    cbDESC.AddItem ItemString, Indice
End Function

'Borrar combo (especialmente para items añadidos de manera manual)
Public Function borra_combo()
    cbDESC.Clear
End Function

Public Function carga()
Dim tmpformato As String
Dim rc As New ADODB.Recordset

On Error GoTo errores


tmpformato = String(LenCodigo, "0")


txtCodigo.Text = ""
cbDESC.Clear
Screen.MousePointer = vbHourglass

rc.Open m_SQLString, m_ConexionString

If Not rc.EOF Then

Do Until rc.EOF
    
    If m_DataField <> "" Then
    cbDESC.AddItem rc.fields(0).Value & Space(LenCodigo) & " " & rc.fields(1).Value
    Else
    cbDESC.AddItem Format(rc.fields(0).Value, tmpformato) & " " & rc.fields(1).Value
    End If
    
    If Not rc.EOF Then rc.MoveNext
Loop

cbDESC.ListIndex = -1
End If

rc.Close
Set rc = Nothing

Screen.MousePointer = vbDefault

'tmpformato = ""

Exit Function
errores:
    'tmpformato = ""
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation, "¡Atención!"
End Function



Public Property Get hwnd() As Long
hwnd = UserControl.hwnd
End Property


Private Sub pfocusnext(ByVal Direction As flxDirect)
Dim control As Object
Dim tabbus As Long
Dim tabactual As Integer

tabactual = Extender.TabIndex


If Direction = Down Then
    tabbus = tabactual + 1
Else
    tabbus = tabactual - 1
End If

    For Each control In Extender.Parent

        If (TypeOf control Is miText) Or (TypeOf control Is miCombo) Then
        
         If control.Visible Then
         
                If Direction = Down Then
                
                    If control.TabIndex = tabbus Then
                        control.SetFocus
                        Exit Sub
                    End If
                
                Else
                
                    If control.TabIndex <= tabbus Then
                        control.SetFocus
                        Exit Sub
                    End If
                
                End If
            
         End If
        
        End If

    Next control


End Sub

Private Sub FocusNextMask(ByVal Direction As flxDirect, Optional bReturnKey As Boolean)
Dim xObject          As Object
Dim ObjHwnds         As New Collection
Dim ObjTabIndex      As New Collection
Dim iNextTabIndex    As Integer
Dim iCurrTabIndex    As Integer
Dim iTabIndex        As Variant
Dim Cancel           As Boolean
Dim l                As Long

iCurrTabIndex = Extender.TabIndex
iNextTabIndex = iCurrTabIndex

For Each xObject In Extender.Parent
If (TypeOf xObject Is FlexMaskEditBox) Or (TypeOf xObject Is miText) Or (TypeOf xObject Is miCombo) Then
If xObject.Enabled And xObject.Visible Then
         ObjHwnds.Add xObject.hwnd, CStr(xObject.TabIndex)
         ObjTabIndex.Add xObject.TabIndex, CStr(xObject.TabIndex)
      End If
   End If
Next

If Direction = Down Then
   For Each iTabIndex In ObjTabIndex
      If iTabIndex > iCurrTabIndex Then
         If iTabIndex <= iNextTabIndex Or iNextTabIndex = iCurrTabIndex Then
            iNextTabIndex = iTabIndex
         End If
      End If
   Next
   If iNextTabIndex = iCurrTabIndex Then
      If bReturnKey Then
         If Not txtCodigo.Enabled Then
            Set ObjHwnds = Nothing
            Set ObjTabIndex = Nothing
            Set xObject = Nothing
            SendKeys "{tab}"
         End If
         Exit Sub
      Else
         For Each iTabIndex In ObjTabIndex
            If iTabIndex < iNextTabIndex Then
               iNextTabIndex = iTabIndex
            End If
         Next
      End If
   End If
ElseIf Direction = Up Then
   For Each iTabIndex In ObjTabIndex
      If iTabIndex < iCurrTabIndex Then
         If iTabIndex >= iNextTabIndex Or iNextTabIndex = iCurrTabIndex Then
            iNextTabIndex = iTabIndex
         End If
      End If
   Next
   If iNextTabIndex = iCurrTabIndex Then
      For Each iTabIndex In ObjTabIndex
         If iTabIndex > iNextTabIndex Then
            iNextTabIndex = iTabIndex
         End If
      Next
   End If
End If

'WithOut UseIng a API
For Each xObject In Extender.Parent
   If (TypeOf xObject Is miText) Or (TypeOf xObject Is miCombo) Then
      If xObject.TabIndex = iNextTabIndex Then xObject.SetFocus
   End If
Next

'If ObjHwnds.Count > 0 Then
'   RaiseEvent ExitOnArrowKeys(Cancel)
'   If Not Cancel Then
'      l = ObjHwnds.Item(CStr(iNextTabIndex))
'      Set ObjHwnds = Nothing
'      Set ObjTabIndex = Nothing
'      Set xObject = Nothing
'      SetFocusAPI l
'   End If
'End If

Set ObjHwnds = Nothing
Set ObjTabIndex = Nothing
Set xObject = Nothing
End Sub




