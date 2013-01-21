VERSION 5.00
Begin VB.UserControl miTextMask 
   BackStyle       =   0  'Transparent
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ScaleHeight     =   465
   ScaleWidth      =   1305
   ToolboxBitmap   =   "miTextMask.ctx":0000
   Begin PCGestion.FlexMaskEditBox txtCodigo 
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   2
      MaskCharInclude =   0   'False
      FontSize        =   11,25
      BeginProperty FormatFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DecimalSeperator=   ","
   End
End
Attribute VB_Name = "miTextMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : miText
' DateTime  : 20/09/2003 12:14
' Author    : José Castillo
' Purpose   : TextBox con algunas características añadidas
'---------------------------------------------------------------------------------------

'*****************************
' POR HACER:
' Establecer propiedades Fechas maxima y minima para campos fecha
' con valores predefinidos: min:  01/01/2003, max: 31/12/2099
' personalizar colores.
'*****************************

Option Explicit

Const fondo_original = vbWhite
Dim fondo_celda As Long

'Default Property Values:
Const m_def_BackColor = vbBlue
Const m_def_ForeColor = vbBlack
Const m_def_Appearance = 1
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Text = "00000000"
Const m_def_DataSource = Null
Const m_def_DataField = ""
Const m_def_intFormat = ""
Const m_def_dspFormat = "Currency"

Const m_def_Mask = "##.##"
Const m_def_LongMaxima = 0
Const m_def_SoloNumeros = False
Const m_def_Alineacion = 0
Const m_def_Valor = ""
Const m_def_PermitirBlanco = True


'Property Variables:
Dim m_LongMaxima As Integer
Dim m_SoloNumeros As Boolean
Dim m_Alineacion As Byte
Dim m_PermitirBlanco  As Boolean

Dim m_mask As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_enabled As Boolean
Dim m_Font As Font
Dim m_Appearance As Integer
Dim m_Text As String
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_DataSource As ADODB.Recordset
Dim m_DataField As String
Dim m_intFormat As String
Dim m_dspFormat As String
Dim m_Locked As Boolean

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
'Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."

Private m_Valor As String

Dim tmptexto As String
Dim esfecha As Boolean



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
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
    txtCodigo.Locked = Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_enabled
    txtCodigo.Enabled = Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_enabled = New_Enabled
    PropertyChanged "Enabled"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get PermitirBlanco() As Boolean
    PermitirBlanco = m_PermitirBlanco
End Property

Public Property Let PermitirBlanco(ByVal New_PermitirBlanco As Boolean)
    m_PermitirBlanco = New_PermitirBlanco
'    txtCodigo.MaxLength = m_PermitirBlanco
    PropertyChanged "PermitirBlanco"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get LongMaxima() As Long
    LongMaxima = m_LongMaxima
End Property

Public Property Let LongMaxima(ByVal New_LongMaxima As Long)
    m_LongMaxima = New_LongMaxima
    txtCodigo.MaxLength = m_LongMaxima
    PropertyChanged "LongMaxima"
End Property

'Alineacion
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Alineacion() As Byte
    Alineacion = m_Alineacion
End Property

Public Property Let Alineacion(ByVal New_Alineacion As Byte)
    m_Alineacion = New_Alineacion
    txtCodigo.Alignment = Alineacion
    PropertyChanged "Alineacion"
End Property

'Mask
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get Mask() As String
    Mask = m_mask
End Property

Public Property Let Mask(ByVal New_Mask As String)
    m_mask = New_Mask
    txtCodigo.Mask = Mask
    PropertyChanged "Mask"
End Property

'Solo Numericos
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get SoloNumeros() As Long
    SoloNumeros = m_SoloNumeros
    'txtCodigo.SoloNumeros = SoloNumeros
End Property

Public Property Let SoloNumeros(ByVal New_SoloNumeros As Long)
    m_SoloNumeros = New_SoloNumeros
    PropertyChanged "SoloNumeros"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
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
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512

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
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indica si un control Label o el color de fondo de un control Shape es transparente u opaco."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
'Public Property Get DataSource() As ADODB.Recordset
'   Set DataSource = m_DataSource
   'Set txtCodigo.DataSource = m_DataSource
'End Property

'Public Property Set DataSource(ByVal New_DataSource As ADODB.Recordset)
'    Set m_DataSource = New_DataSource
'    Set txtCodigo.DataSource = m_DataSource
'    PropertyChanged "DataSource"
'End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
'Public Property Get DataField() As String
'    DataField = m_DataField
'    'txtCodigo.DataField = DataField
'End Property

'Public Property Let DataField(ByVal New_DataField As String)
   ' m_DataField = New_DataField
  '  PropertyChanged "DataField"
 '   txtCodigo.DataField = DataField
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get intFormat() As String
    intFormat = m_intFormat
End Property

Public Property Let intFormat(ByVal New_intFormat As String)
    m_intFormat = New_intFormat
    PropertyChanged "intFormat"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get dspFormat() As String
    dspFormat = m_dspFormat
End Property

Public Property Let dspFormat(ByVal New_dspFormat As String)
    m_dspFormat = New_dspFormat
    PropertyChanged "dspFormat"
    
   
End Property


Private Sub txtCodigo_GotFocus()
    
    If txtCodigo.Locked = False Then
    
    
           
        txtCodigo.BackColor = &HFFFF80
    
        If Trim(tmptexto) <> "" And m_DataField = "" Then
            txtCodigo.Text = tmptexto
            txtCodigo.Refresh
     
        ElseIf m_DataField <> "" Then
        
            txtCodigo.Text = Trim(txtCodigo.Text)
    
        End If
    
        SendKeys "{end}"
    
    End If
    
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
 
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
                If Not Trim(tmptexto = "") And Trim(txtCodigo.Text = "") And Not m_PermitirBlanco Then
                    KeyCode = 0
                    txtCodigo.SetFocus
                    Exit Sub
                End If
 End Select
End Sub


'Si devuelve TRUE, los datos se han validado correctamente
'Si devuelve FALSE, los datos no se han validado correctamente,
'impedir el paso.

Private Function valida_datos() As Boolean
Dim tmpstr As String
Dim tmpstr2 As String
Dim tmpvalida As Boolean

On Error GoTo errores
'    txtCodigo.BackColor = fondo_original

tmptexto = txtCodigo.Text

    'si esta en blanco, pero no es permitido, impedir validacion ...
    If Trim(tmptexto = "") And (m_PermitirBlanco = False) Then
                    valida_datos = False
                    tmptexto = ""
                    Exit Function
    End If
    
    'si esta en blanco y ademas es permitido, permitir validación ...
    If Trim(tmptexto = "") And (m_PermitirBlanco = True) Then
                    valida_datos = True
                    tmptexto = ""
                    Exit Function
    End If
    
    
          'establecer solo numericos para formatos : 000 ...
   If InStr(1, m_dspFormat, "0", vbTextCompare) > 0 Then
    m_SoloNumeros = True
   ' 'reemplazar la coma por el punto
    txtCodigo.Text = Replace(txtCodigo.Text, ",", ".")
   
   
   End If
   
    tmpstr = tmptexto
    tmpstr2 = Trim(txtCodigo.Text)
        
        
    'si es numérico
    If m_SoloNumeros Or UCase(dspFormat) = "CURRENCY" Then  'si esta como solo numerico, comprobar
               If Not IsNumeric(tmpstr2) Then
                        txtCodigo.Text = tmpstr2
                        m_SoloNumeros = True
'vbRed
                        valida_datos = False
                        Exit Function
                End If
                
    Else
                If InStr(1, dspFormat, "dd", vbTextCompare) > 0 Then esfecha = True
                If Not esfecha Then If InStr(1, dspFormat, "mm", vbTextCompare) > 0 Then esfecha = True
                If Not esfecha Then If InStr(1, dspFormat, "yy", vbTextCompare) > 0 Then esfecha = True
                
                If esfecha Then
                    If Not IsDate(tmpstr2) Then
                        txtCodigo.Text = tmpstr2
'                        txtCodigo.BackColor = vbRed
                        'txtCodigo.SetFocus
                        valida_datos = False
                        Exit Function
                    End If
                End If
                
    End If
                                     
tmptexto = tmpstr
txtCodigo.Text = tmpstr2

tmpstr = ""
tmpstr2 = ""

valida_datos = True

Exit Function
errores:

    valida_datos = False
    tmptexto = tmpstr
    txtCodigo.Text = tmpstr2
    tmpstr = ""
    tmpstr2 = ""

End Function

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
       
    Select Case KeyAscii
    Case vbKeyBack
                Exit Sub
   
    Case 13
    
            'If Trim(txtCodigo.Text = "") And Not m_PermitirBlanco Then
            'KeyAscii = 0
            'Exit Sub
            'End If
            
            KeyAscii = 0
            SendKeys "{tab}"
            
    Case Else
    

    
            If m_SoloNumeros Then
            
            If Chr(KeyAscii) = "." Then
                
                    'KeyAscii = 0
                    'SendKeys "{.}"
                    KeyAscii = Asc(",")
            
            End If
            
                Call solo_numerico(KeyAscii)
            End If
          '  tmptexto = txtCodigo.Text
        
    End Select
           
End Sub

'Ponerle el codigo al usuario
Public Property Get Text() As String
   
   If (m_SoloNumeros = True And m_Locked = False) And UCase(m_dspFormat) <> "CURRENCY" Then
        'reemplazar la coma por el punto
    txtCodigo.Text = Replace(txtCodigo.Text, ",", ".")
   End If
   
   Text = txtCodigo.Text
   
  ' tmptexto = Text
End Property

'Obtener el .Text del usuario
Public Property Let Text(ByVal iTexto As String)
  
  tmptexto = Format(iTexto, intFormat)
  m_Valor = tmptexto
  txtCodigo.Text = Format(iTexto, dspFormat)
  
    'si es <> "" y tiene formato ...
  '  If Trim(iTexto) <> "" And Trim(dspFormat) <> "" Then
  '      tmptexto = Format(iTexto, intFormat)
  '      txtCodigo.Text = Format(iTexto, dspFormat)
  '  'si es <> "" y no tiene formato ...
  '  ElseIf Trim(iTexto) <> "" And Trim(dspFormat) = "" Then
  '      txtCodigo.Text = iTexto
  '  End If

Call UserControl.PropertyChanged("Text")
End Property

Public Property Get Valor() As String
Dim tmpval As String

           ' If m_SoloNumeros = True Then
            
                'si es numerico, cambiarle el punto por la coma
                'para que se obtenga bien al recoger el valor
            '    tmpval = tmptexto
            '    tmpval = Replace(tmpval, ".", ",")
             '   Valor = tmpval
             '   tmpval = ""
            
           ' Else
                        
            Valor = Format(tmptexto, intFormat)
          '  End If
End Property

Public Property Let Valor(ByVal fvalor As String)
  m_Valor = fvalor
    
  tmptexto = Format(fvalor, intFormat)
  txtCodigo.Text = Format(fvalor, dspFormat)
        
    Call UserControl.PropertyChanged("Valor")
End Property


Private Sub txtCodigo_lostFocus()

       'establecer solo numericos para formatos : 000 ...
 '  If (m_DataField <> "") And m_SoloNumeros > 0 And m_Locked = False Then
 '       'reemplazar la coma por el punto
 '   txtCodigo.Text = Replace(txtCodigo.Text, ",", ".")
 '  End If
 
   If (m_SoloNumeros = True And m_Locked = False) And UCase(m_dspFormat) <> "CURRENCY" Then
        'reemplazar la coma por el punto
    txtCodigo.Text = Replace(txtCodigo.Text, ",", ".")
   End If
   

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

        If Not Trim(tmptexto) = "" Then
                        
                        If m_DataField = "" Then
                        
                        If UCase(dspFormat) = "CURRENCY" Then
   
                        'si es currency, hacer el replace al reves para que
                        'cambie el punto por la coma, porque sino no coge bien
                        'los decimales
                        txtCodigo.Text = Replace(txtCodigo.Text, ".", ",")
                        tmptexto = txtCodigo.Text
                        
                        End If
   
                        txtCodigo.Text = Format(tmptexto, m_dspFormat)
                        'si es fecha que reemplace (para fechas parcialmente
                        'introducidas
                        If esfecha Then tmptexto = txtCodigo
                        
                                           
                        End If
        End If
    
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
    
    m_Text = m_def_Text
    Set m_Font = Ambient.Font
    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle

    m_mask = m_def_Mask
    m_intFormat = m_def_intFormat
    m_dspFormat = m_def_dspFormat
    m_Locked = False
    m_enabled = True
    
    m_PermitirBlanco = m_def_PermitirBlanco

    m_LongMaxima = m_def_LongMaxima
    m_SoloNumeros = m_def_SoloNumeros
    m_Alineacion = m_def_Alineacion
    
    txtCodigo.BackColor = fondo_original
    
        
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
     
   
    m_mask = PropBag.ReadProperty("Mask", m_def_Mask)
    
    m_intFormat = PropBag.ReadProperty("intFormat", m_def_intFormat)
    m_dspFormat = PropBag.ReadProperty("dspFormat", m_def_dspFormat)
    m_dspFormat = PropBag.ReadProperty("Valor", m_def_Valor)
    m_LongMaxima = PropBag.ReadProperty("LongMaxima", m_def_LongMaxima)
    m_SoloNumeros = PropBag.ReadProperty("SoloNumeros", m_def_SoloNumeros)
    m_Alineacion = PropBag.ReadProperty("Alineacion", m_def_Alineacion)
    m_PermitirBlanco = PropBag.ReadProperty("PermitirBlanco", m_def_PermitirBlanco)
    m_Locked = PropBag.ReadProperty("Locked", False)
    m_enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

Private Sub UserControl_Resize()
If UserControl.Width - 50 > 0 Then
       txtCodigo.Width = UserControl.Width - 50
End If
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Mask", m_mask, m_def_Mask)
    Call PropBag.WriteProperty("intFormat", m_intFormat, m_def_intFormat)
    Call PropBag.WriteProperty("dspFormat", m_dspFormat, m_def_dspFormat)

    Call PropBag.WriteProperty("LongMaxima", m_LongMaxima, m_def_LongMaxima)
    Call PropBag.WriteProperty("SoloNumeros", m_SoloNumeros, m_def_SoloNumeros)
    Call PropBag.WriteProperty("Alineacion", m_Alineacion, m_def_Alineacion)
    Call PropBag.WriteProperty("PermitirBlanco", m_PermitirBlanco, m_def_PermitirBlanco)

    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("Enabled", m_enabled, False)
    
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
        
        If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
  
    End Select
    
    'If KeyAscii = 13 Or KeyAscii = 8 Then Exit Sub
    
End Sub


Public Function CancelarValidacion()

'txtCodigo.SetFocus
txtCodigo.BackColor = vbYellow
SendKeys "{end}"

End Function

