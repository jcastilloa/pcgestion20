VERSION 5.00
Object = "{9A326101-2FD8-11D2-9BC0-00A024695830}#1.0#0"; "tinumbl8.ocx"
Begin VB.UserControl miTextNumLite 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ScaleHeight     =   495
   ScaleWidth      =   1335
   ToolboxBitmap   =   "miTextNumLite.ctx":0000
   Begin TDBNumLite6Ctl.TDBNumLite txtCodigo 
      Height          =   450
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   794
      Caption         =   "miTextNumLite.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "miTextNumLite.ctx":037E
      Spin            =   "miTextNumLite.ctx":03BA
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClearAction     =   0
      DecimalPlaces   =   0
      DecimalPoint    =   ","
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      HighlightText   =   0
      IntegerPlaces   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinusSign       =   1
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      ShowContextMenu =   -1
      ValueVT         =   1179653
      Value           =   0
      MaxValueVT      =   1885667333
      MinValueVT      =   1818820613
   End
End
Attribute VB_Name = "miTextNumLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const fondo_original = vbWhite
Dim fondo_celda As Long

'Default Property Values:
Const m_def_BackColor = vbBlue
Const m_def_ForeColor = vbBlack
Const m_def_Appearance = 1
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Text = "00"
Const m_def_DataSource = Null
Const m_def_DataField = ""

Const m_def_DecimalPlaces = 0
Const m_def_MaxValue = 999999999
Const m_def_MinValue = 0
Const m_def_IntegerPlaces = 0


'Const m_def_format = "##.##"
'Const m_def_displayformat = "00.00"

'Const m_def_LongMaxima = 0
Const m_def_Value = ""
Const m_def_Enabled = True

Const m_def_PermitirCero = True

Dim m_DecimalPlaces As Byte
Dim m_IntegerPlaces As Byte
Dim m_MaxValue As Double
Dim m_MinValue As Double


Dim m_PermitirCero  As Boolean

Dim m_BackColor As Long
Dim m_enabled As Boolean
Dim m_ForeColor As Long
Dim m_Font As Font
Dim m_Appearance As Integer
Dim m_Text As String
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_DataSource As adodb.Recordset
Dim m_DataField As String
Dim m_Locked As Boolean

'Event Declarations:
Event Click()
'Event DblClick()
Event KeyDown(KeyCode As Integer, shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, shift As Integer)
Event MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)

Private m_Value As String

Dim tmptexto As String
Dim esfecha As Boolean



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
    BackColor = m_BackColor
    txtCodigo.BackColor = BackColor
    'cbDESC.BackColor = BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Locked() As Boolean
    Locked = m_Locked
    'txtCodigo.Enabled = Not Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_enabled
    txtCodigo.Enabled = Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_enabled = New_Enabled
    PropertyChanged "Enabled"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get PermitirCero() As Boolean
    PermitirCero = m_PermitirCero
End Property

Public Property Let PermitirCero(ByVal New_PermitirCero As Boolean)
    m_PermitirCero = New_PermitirCero
'    txtCodigo.MaxLength = m_PermitirCero
    PropertyChanged "PermitirCero"
End Property


'Solo Numericos

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
    txtCodigo.ForeColor = ForeColor
    'cbDESC.ForeColor = ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=6,0,0,0
Public Property Get Font() As Font

    Set Font = m_Font
    Set txtCodigo.Font = Font
    'Set cbDESC.Font = Font
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
Public Property Get DataSource() As adodb.Recordset
   Set DataSource = m_DataSource
   'Set txtCodigo.DataSource = m_DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As adodb.Recordset)
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

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get MaxValue() As String
    MaxValue = m_MaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As String)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
    txtCodigo.MaxValue = m_MaxValue
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get MinValue() As String
    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As String)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
    txtCodigo.MinValue = m_MinValue
End Property


Public Property Get DecimalPlaces() As String
    DecimalPlaces = m_DecimalPlaces
End Property

Public Property Let DecimalPlaces(ByVal New_DecimalPlaces As String)
    m_DecimalPlaces = New_DecimalPlaces
    PropertyChanged "DecimalPlaces"
    txtCodigo.DecimalPlaces = m_DecimalPlaces
End Property

Public Property Get IntegerPlaces() As String
    IntegerPlaces = m_IntegerPlaces
End Property

Public Property Let IntegerPlaces(ByVal New_IntegerPlaces As String)
    m_IntegerPlaces = New_IntegerPlaces
    PropertyChanged "IntegerPlaces"
    txtCodigo.IntegerPlaces = m_IntegerPlaces
End Property





Private Sub txtCodigo_GotFocus()
    
    'If txtCodigo.Locked = False Then
       
      '  If (DataSource.EditMode <> adEditAdd) And (DataSource.EditMode <> adEditInProgress) Then
       '
        'txtCodigo.Enabled = False
        '
  '      Else
   '
    '    txtCodigo.Enabled = True
     '
      '  End If
        
        txtCodigo.BackColor = &HFFFF80
        SendKeys "{end}"
    
    'End If
    
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, shift As Integer)
 
 Select Case KeyCode
  '  Case vbKeyF1
  '  Case vbKeyF2
  '  Case vbKeyF3
  '  Case vbKeyF4
'
'
'
'    Case vbKeyF5
'    Case vbKeyF6
'    Case vbKeyF7
'    Case vbKeyF8
'    Case vbKeyF9
'    Case vbKeyF10
 Case vbKeyTab
 
   If valida_datos = False Then
    KeyCode = 0
    txtCodigo.SetFocus
   End If
               ' If Trim(txtCodigo.Text = "0") And Not m_PermitirCero Then
                '    KeyCode = 0
                 '   txtCodigo.SetFocus
                  '  Exit Sub
               ' End If
 End Select
End Sub


'Si devuelve TRUE, los datos se han validado correctamente
'Si devuelve FALSE, los datos no se han validado correctamente,
'impedir el paso.

Private Function valida_datos() As Boolean

    'si esta en blanco, pero no es permitido, impedir validacion ...
    If Trim(txtCodigo.Value = "0") And (m_PermitirCero = False) Then
                    valida_datos = False
                    Exit Function
    End If
    
    valida_datos = True

End Function

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
       
    Select Case KeyAscii
    Case vbKeyBack
                Exit Sub
   
    Case 13
    
            'If Trim(txtCodigo.Text = "") And Not m_PermitirCero Then
            'KeyAscii = 0
            'Exit Sub
            'End If
            
            KeyAscii = 0
            SendKeys "{tab}"
            
        
    End Select
           
End Sub

'Ponerle el codigo al usuario
Public Property Get Text() As String
   Text = txtCodigo.Text
  ' tmptexto = Text
End Property

'Obtener el .Text del usuario
Public Property Let Text(ByVal iTexto As String)
  
  'tmptexto = format(iTexto, format)
  'm_Value = tmptexto
  txtCodigo.Text = iTexto
  'txtCodigo.Value = iTexto
  
    'si es <> "" y tiene formato ...
  '  If Trim(iTexto) <> "" And Trim(displayformat) <> "" Then
  '      tmptexto = Format(iTexto, format)
  '      txtCodigo.Text = Format(iTexto, displayformat)
  '  'si es <> "" y no tiene formato ...
  '  ElseIf Trim(iTexto) <> "" And Trim(displayformat) = "" Then
  '      txtCodigo.Text = iTexto
  '  End If

Call UserControl.PropertyChanged("Text")
End Property

Public Property Get Value() As String

            If Not IsNull(txtCodigo.Value) Then
            Value = txtCodigo.Value
            End If

End Property

Public Property Let Value(ByVal fValue As String)
  
  m_Value = fValue
  txtCodigo.Value = fValue
        
    Call UserControl.PropertyChanged("Value")
End Property


Private Sub txtCodigo_lostFocus()

       'establecer solo numericos para formatos : 000 ...
   'If (m_DataField <> "") And m_SoloNumeros > 0 And m_Locked = False Then
        'reemplazar la coma por el punto
    
    'If valida_datos Then
     '   txtCodigo.Text = Replace(txtCodigo.Text, ",", ".")
    'End If

   

'If Not validado Then
 '   Call txtCodigo_Validate(False)
  '  Exit Sub
'End If

    With txtCodigo
           .BackColor = fondo_original
        End With
    

    
End Sub

Private Sub txtCodigo_Validate(Cancel As Boolean)
    
If valida_datos Then
   
        txtCodigo.BackColor = fondo_original
        Cancel = False
       
            
Else
        Cancel = True
        txtCodigo.BackColor = vbYellow
        
End If
    
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    
    m_enabled = m_def_Enabled
    
    m_Text = m_def_Text
    m_Value = m_def_Value
    
    Set m_Font = Ambient.Font
    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    
    If Not IsNull(m_def_DataSource) Then
    Set m_DataSource = m_def_DataSource
    Else
    Set m_DataSource = Nothing
    End If
    
    m_DataField = m_def_DataField
    
    m_DecimalPlaces = m_def_DecimalPlaces
    m_MaxValue = m_def_MaxValue
    m_MinValue = m_def_MinValue
    m_IntegerPlaces = m_def_IntegerPlaces
       
    m_PermitirCero = m_def_PermitirCero
    
    txtCodigo.BackColor = fondo_original
    
        
End Sub

'Cargar Valuees de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    
    If Not IsNull(m_def_DataSource) Then _
    Set m_DataSource = PropBag.ReadProperty("DataSource", m_def_DataSource)
    
    If m_def_DataField <> "" Then _
    m_DataField = PropBag.ReadProperty("DataField", m_def_DataSource)
    
    m_DecimalPlaces = PropBag.ReadProperty("DecimalPlaces", m_DecimalPlaces)
    m_IntegerPlaces = PropBag.ReadProperty("IntegerPlaces", m_IntegerPlaces)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_PermitirCero = PropBag.ReadProperty("PermitirCero", m_def_PermitirCero)
    
End Sub

Private Sub UserControl_Resize()
If UserControl.Width - 50 > 0 Then
       txtCodigo.Width = UserControl.Width - 50
End If
End Sub

'Escribir Valuees de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
       
    Call PropBag.WriteProperty("Enabled", m_enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("DataSource", m_DataSource, m_def_DataSource)
    Call PropBag.WriteProperty("DataField", m_DataField, m_def_DataField)
    
    
    m_DecimalPlaces = PropBag.ReadProperty("DecimalPlaces", m_DecimalPlaces)
    m_IntegerPlaces = PropBag.ReadProperty("IntegerPlaces", m_IntegerPlaces)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    
    
    Call PropBag.WriteProperty("DecimalPlaces", m_DecimalPlaces, m_def_DecimalPlaces)
    Call PropBag.WriteProperty("IntegerPlaces", m_IntegerPlaces, m_def_IntegerPlaces)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)


    Call PropBag.WriteProperty("PermitirCero", m_PermitirCero, m_def_PermitirCero)

    
End Sub

Private Sub solo_numerico(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyF1
    Case vbKeyF2
    Case vbKeyF3
    Case vbKeyF4
    Case vbKeyF5
    Case vbKeyF6
    Case vbKeyF7
    Case vbKeyF8
    Case vbKeyF9
    Case vbKeyF10
    Case vbKeyHome
    Case vbKeyEnd
    Case vbKeyPageDown
    Case vbKeyPageUp

    Case 13
    Case 8
    Case vbKeyTab
    Case Else
        
        'If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
  
    End Select
    
    'If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
    
End Sub


Public Function CancelarValidacion()

'txtCodigo.SetFocus
txtCodigo.BackColor = vbYellow
DoEvents
SendKeys "{end}"

End Function





