VERSION 5.00
Begin VB.UserControl ucGrdBttn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BackStyle       =   0  'Transparent
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   PropertyPages   =   "GradCtrl.ctx":0000
   ScaleHeight     =   630
   ScaleWidth      =   2490
   ToolboxBitmap   =   "GradCtrl.ctx":004D
   Begin VB.PictureBox picImageHold 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picGradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   90
      ScaleHeight     =   510
      ScaleWidth      =   1245
      TabIndex        =   0
      Tag             =   "GrdButton"
      Top             =   90
      Width           =   1245
   End
End
Attribute VB_Name = "ucGrdBttn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private bLoad As Boolean

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER = &H1

Public Enum geGrdBorderWidth
    gbwBorderWidthNormal = 2
    gbwBorderWidthStrong = 3
End Enum
Public Enum geGrdButtonColors
    gbcGradient
    gbcOneColor
End Enum
Public Enum geImageCaptionPosition
    icpImageCaption_LeftRight
    icpImageCaption_TopBottom
    icpCaptionImage_LeftRight
    icpCaptionImage_TopBottom
End Enum
Public Enum geDisabledEmbossLevel
    delEmbossNormal
    delEmbossStrong
End Enum

Private mtRectImgTW As gtypeRect    '   holds data in twips
'Public mtRectImgPX As gtypeRect    '   holds data in pixel
Private mtRectCaptn As gtypeRect    '   holds data in twips (caption)
'Public mtRectMnemc As gtypeRect    '   holds data in twips (underline for mnemonic)
Private mc_clsGradient       As New clsGradient
Private mbLoadPicture        As Boolean

Dim bDoCaption          As Boolean
Dim bMouseDown          As Boolean
Dim bHasFocus           As Boolean
Dim bInit               As Boolean
Dim bReadProp           As Boolean
Dim bWriteProp          As Boolean

Private Const conDefCaption$ = "GrdButton"
Private Const conDefGtColor1& = vbButtonFace       '&H8000000F  '  COLOR_WINDOWTEXT - old
Private Const conDefGtColor2& = vbButtonFace       '&H8000000F  '  COLOR_BTNFACE
Private Const conDefBdColor1& = vbHighlightText    '&H8000000E  '  COLOR_HIGHLIGHTTEXT
Private Const conDefBdColor2& = vb3DDKShadow       '&H80000015  '  COLOR_DARKSHADOW
Private Const conDefGtRatio1# = 1#
Private Const conDefGtRatio2# = 0.2

Private Const conDefForeColor& = vbActiveTitleBarText   '   &H80000009&
Private Const conDefAngle! = 120!

Private Const conShift! = 8!

Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event KeyDown(KeyCode%, Shift%)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(KeyAscii%)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyUp(KeyCode%, Shift%)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseDown(Button%, Shift%, X!, Y!)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseUp(Button%, Shift%, X!, Y!)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607

Private me_BorderWidth      As geGrdBorderWidth
Private me_ImgCapPos        As geImageCaptionPosition
Private me_DisEmboss        As geDisabledEmbossLevel
Private me_GrdBttnColor     As geGrdButtonColors
Private mp_Angle            As Single
Private mp_Percent          As Integer
Private mp_GColor1          As OLE_COLOR
Private mp_GColor2          As OLE_COLOR
Private mp_BColor1          As OLE_COLOR
Private mp_BColor2          As OLE_COLOR
Private mp_FocusColor       As OLE_COLOR
Private mp_GColor_Boost0    As Boolean

Private mp_GColor1_Ratio    As Double
Private mp_GColor2_Ratio    As Double

Private Sub picGradient_Click()
    RaiseEvent Click
End Sub

Private Sub picGradient_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picGradient_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picGradient_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'    Call fRaiseEvent("Click")
'End Sub

Private Sub fDrawFocus()
    Dim R As RECT
    
    Const conS! = 60!
    Const conOffset = 3&
        
    If Not UserControl.Enabled Then Exit Sub
    
    With picGradient
        Call SetRect(R, conOffset, conOffset, _
                    .Width / Screen.TwipsPerPixelX - conOffset, _
                    .Height / Screen.TwipsPerPixelY - conOffset)
        'Call OffsetRect(R, 3, 3)

        Call DrawFocusRect(.hdc, R)
    End With
    
    Exit Sub
        
'        With picGradient
'    '        .DrawStyle = vbDot
'    '        .DrawWidth = 1
'    '        picGradient.Line (conS!, conS!)-(.Width - conS!, .Height - conS!), lColor, B
'            Call fDrawDottedRect(picGradient, mp_FocusColor, conS, conS, _
'                           .Width - conS!, .Height - conS!)
'    '        .DrawStyle = vbSolid
'    '        .DrawWidth = 2
'
'        End With
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii%)
    RaiseEvent Click
'    Select Case KeyAscii
'        Case vbKeyReturn
'            RaiseEvent Click    '   'Default' behaviour
'        Case vbKeyEscape
'            RaiseEvent Click    '   'Cancel' behaviour
'    End Select
End Sub

'Private Sub UserControl_AmbientChanged(PropertyName$)
'    Select Case PropertyName
'        Case "DisplayAsDefault"
'            If AmbientProperties.DisplayAsDefault Then
'
'            Else
'
'            End If
'    End Select
'End Sub

Private Sub UserControl_EnterFocus()
    If bHasFocus Then Exit Sub
    
    bHasFocus = True
    
    Call fDrawFocus
End Sub

Private Sub UserControl_ExitFocus()
    If Not bHasFocus Then Exit Sub
    
    bHasFocus = False
    
    'Debug.Print "ExitFocus"
    Call fDrawButton
End Sub

Private Sub UserControl_GotFocus()
    If bHasFocus Then Exit Sub
    
    bHasFocus = True
    
    Call fDrawFocus
End Sub

Private Sub UserControl_LostFocus()
    If Not bHasFocus Then Exit Sub
    
    bHasFocus = False
    
    'Debug.Print "LostFocus"
    Call fDrawButton
End Sub

Private Sub picGradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    If bMouseDown Then Exit Sub
    
    If Button = vbLeftButton Then
        Call fMouseDown
        bMouseDown = True
    End If
End Sub

Private Sub picGradient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    If bMouseDown Then
        bMouseDown = False
        Call fMouseUp
    End If
End Sub

Private Sub fMouseDown()
    Const conStartLine! = 20!
    
    With picGradient
        .Move .Left + conShift, .Top + conShift
        
        picGradient.Line (conStartLine, conStartLine)- _
                         (.Width - conStartLine, .Height - conStartLine), _
                         vbBlack, B
    End With
End Sub

Private Sub fMouseUp()
    With picGradient
        .Move .Left - conShift, .Top - conShift
    End With
    'Debug.Print "Mouseup"
    Call fDrawButton    '   overdraw lines
End Sub

Private Sub UserControl_Initialize()
    bInit = True
    'Debug.Print "UserControl_Initialize start"
    mp_Percent = 50
    bMouseDown = False
    picGradient.Left = 0
    picGradient.Top = 0
    UserControl.Height = picGradient.Height + conShift
    UserControl.Width = picGradient.Width + conShift
    
    bInit = False
    'Debug.Print "UserControl_Initialize end"
    'Call UserControl_Resize
End Sub

Private Sub fDrawButton(Optional bShift As Boolean = False)
    Dim sngImgW!, sngShiftCaption!
    Dim lTemp&, R As RECT, lTpPX&, lTpPY&
    
    If bInit Or bReadProp Or bWriteProp Then Exit Sub
    If mp_GColor1_Ratio = False Then Exit Sub

#If TestDraw Then
    Static i%
    i% = i% + 1
    Debug.Print "fDrawButton", i%
#End If
    Const conI% = 128
    Const conT& = 1
    
    If bMouseDown Then Exit Sub
    
    With UserControl
        If .Width < conShift Then .Width = conShift
        If .Height < conShift Then .Height = conShift
    
        picGradient.Width = .Width - conShift
        picGradient.Height = .Height - conShift
    End With
    
    Call fGetCaptionImageSize
    
    With picGradient
        
        Call fSetGradient
            
        If bDoCaption Then
            .CurrentX = mtRectCaptn.Left
            .CurrentY = mtRectCaptn.Top
                        
            If UserControl.Enabled Then
                If fGetBitsPerPixel(UserControl.hdc) > 8 Then
                    .ForeColor = picImageHold.ForeColor
                Else
                    .ForeColor = fGetSysColor(COLOR_BTNTEXT)
                End If
            Else
                .ForeColor = vbWhite
            End If
            
            'picGradient.Print .Tag
            
            lTpPX& = Screen.TwipsPerPixelX
            lTpPY& = Screen.TwipsPerPixelY
            With mtRectCaptn
                SetRect R, .Left / lTpPX&, .Top / lTpPY&, (.Left + .Width) / lTpPX&, (.Top + .Height) / lTpPY&  '+ 10
            End With
                DrawText .hdc, .Tag, Len(.Tag), R, DT_CENTER
        End If
        
        If mbLoadPicture Then
            .PaintPicture picImageHold, mtRectImgTW.Left, mtRectImgTW.Top
        End If
            
        If Not UserControl.Enabled Then
            Call fMakeEmboss(.Image, fGetSysColor(COLOR_BTNFACE), mProgress, me_DisEmboss)
'            lTemp = .ScaleWidth
'            .ScaleWidth = 3
'
'            picGradient.Line (conT&, conT&)-(.Width - conT&, .Height - conT&), _
'                              fGetSysColor(COLOR_BTNSHADOW), B
'            .ScaleWidth = lTemp&
        End If
        .Refresh
    End With
End Sub

'Private Sub UserControl_Paint()
'    Call fDrawButton
'End Sub

Private Sub UserControl_Resize()
    If Not bInit Then
        'Debug.Print "resize"
        Call fDrawButton
    End If
End Sub

'Public Static Property Let About(sAbout As String)
'    '
'End Property

Public Property Get About() As String
Attribute About.VB_ProcData.VB_Invoke_Property = "ppAbout"
    About = fAbout
End Property

'   FocusColor
Public Property Get FocusColor() As OLE_COLOR
Attribute FocusColor.VB_Description = "Color of the dotted line, which appears when the control is getting focus (default - black)"
    FocusColor = mp_FocusColor
End Property
Public Property Let FocusColor(ByVal NewColor As OLE_COLOR)
    mp_FocusColor = NewColor
        
    'Debug.Print "focuscolor"
    Call fDrawButton
    
    PropertyChanged "FocusColor"
End Property

'   Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
    UserControl.Enabled = bEnabled
        
    'Debug.Print "enabled"
    Call fDrawButton
    
    PropertyChanged "Enabled"
End Property

'   Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Image.VB_ProcData.VB_Invoke_Property = "StandardPicture"
    Set Image = picImageHold
End Property
Public Property Set Image(ByVal NewButtonImage As Picture)
    On Error GoTo ErrOccurred
    Set picImageHold = NewButtonImage
    
    'Debug.Print "image"
    Call fDrawButton
Exit_:
    PropertyChanged "Image"
    
    Exit Property
    
ErrOccurred:
    MsgBox "Error loading image...", vbCritical, App.Title
    
    Set picImageHold = Nothing
    
    Resume Exit_
End Property

'   image-caption position
Public Property Get ImageCaptionPos() As geImageCaptionPosition
Attribute ImageCaptionPos.VB_Description = "Changes order of caption and image as it appears on the button"
Attribute ImageCaptionPos.VB_ProcData.VB_Invoke_Property = "ppGrdButton"
    ImageCaptionPos = me_ImgCapPos
End Property
Public Property Let ImageCaptionPos(ByVal NewPos As geImageCaptionPosition)
    
    If NewPos > geImageCaptionPosition.icpCaptionImage_TopBottom Then
        NewPos = geImageCaptionPosition.icpCaptionImage_TopBottom
    End If
    If NewPos < geImageCaptionPosition.icpImageCaption_LeftRight Then
        NewPos = geImageCaptionPosition.icpImageCaption_LeftRight
    End If
    
    me_ImgCapPos = NewPos
    
    'Debug.Print "position"
    Call fDrawButton
    
    PropertyChanged "ImageCaptionPos"
End Property

'
Public Property Get ButtonColors() As geGrdButtonColors
Attribute ButtonColors.VB_Description = "You can choose the mode of button:\r\ngbcGradient  - (default) gradient filling\r\ngbcOneColor - one color filling (in this case GradientColor2 is ignored)"
    ButtonColors = me_GrdBttnColor
End Property
Public Property Let ButtonColors(ByVal NewBehavior As geGrdButtonColors)
    
    If NewBehavior > geGrdButtonColors.gbcOneColor Then
        NewBehavior = geGrdButtonColors.gbcOneColor
    End If
    If NewBehavior < geGrdButtonColors.gbcGradient Then
        NewBehavior = geGrdButtonColors.gbcGradient
    End If
    
    me_GrdBttnColor = NewBehavior
    
    'Debug.Print "buttoncolors"
    Call fDrawButton
    
    PropertyChanged "ButtonColors"
End Property

Private Function fAssignProprtyInDisignMode() As Boolean
    fAssignProprtyInDisignMode = False
    
    On Error Resume Next
    If (Not bReadProp) And (Not bWriteProp) And (Not bInit) Then
        If Not UserControl.Ambient.UserMode Then
            fAssignProprtyInDisignMode = True
        End If
    End If
End Function

Public Property Get BorderWidth() As geGrdBorderWidth
    BorderWidth = me_BorderWidth
End Property
Public Property Let BorderWidth(ByVal NewWidth As geGrdBorderWidth)
    '   disable the property
    '_________________________________________________________________'
    '¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
    me_BorderWidth = gbwBorderWidthNormal
    
    If fAssignProprtyInDisignMode Then
        MsgBox LoadResString(102), vbExclamation, App.ProductName
    End If
    
    Exit Property
    '_________________________________________________________________'
    '¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
    If NewWidth < geGrdBorderWidth.gbwBorderWidthNormal Then
        NewWidth = geGrdBorderWidth.gbwBorderWidthNormal
    End If
    If NewWidth > geGrdBorderWidth.gbwBorderWidthStrong Then
        NewWidth = geGrdBorderWidth.gbwBorderWidthStrong
    End If
    
    me_BorderWidth = NewWidth
    picGradient.DrawWidth = me_BorderWidth
    
    'Debug.Print "borderwidth"
    Call fDrawButton
    
    PropertyChanged "BorderWidth"
End Property

'   disable Emboss level    (property suspended yet)
Public Property Get DisabledEmboss() As geDisabledEmbossLevel
Attribute DisabledEmboss.VB_Description = "Sets the Emboss level of the disabled control: \r\ndelEmbossNormal or delEmbossStrong\r\n(temporary unavailable)\r\n"
    ImageCaptionPos = me_ImgCapPos
End Property
Public Property Let DisabledEmboss(ByVal NewLevel As geDisabledEmbossLevel)
    '   disable the property
    '_________________________________________________________________'
    '¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
    me_DisEmboss = delEmbossNormal
    
    If fAssignProprtyInDisignMode Then
        MsgBox LoadResString(102), vbExclamation, App.ProductName
    End If
    
    Exit Property
    '_________________________________________________________________'
    '¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
    
    If NewLevel < geDisabledEmbossLevel.delEmbossNormal Then
        NewLevel = geDisabledEmbossLevel.delEmbossNormal
    End If
    If NewLevel > geDisabledEmbossLevel.delEmbossStrong Then
        NewLevel = geDisabledEmbossLevel.delEmbossStrong
    End If
    
    me_DisEmboss = NewLevel
    
    'Debug.Print "disabledemboss"
    Call fDrawButton
    
    PropertyChanged "DisabledEmboss"
End Property

'   GradientColor1
Public Property Get GradientColor1() As OLE_COLOR
Attribute GradientColor1.VB_Description = "Color #1 for gradient filling. \r\nIf ButtonColor = gbcOneColor then this'll be a color of the button"
Attribute GradientColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientColor1 = mp_GColor1
End Property
Public Property Let GradientColor1(ByVal NewColor As OLE_COLOR)
    mp_GColor1 = NewColor
    
    'Debug.Print "gradientcolor1"
    Call fDrawButton
    
    PropertyChanged "GradientColor1"
End Property

'   GradientColor1_Percent
Public Property Get GradientColor1_Percent%()
Attribute GradientColor1_Percent.VB_Description = "Set percent of the Gradient Color 1\r\n(less the number - darker the color)"
    GradientColor1_Percent = mp_GColor1_Ratio * 100
End Property
Public Property Let GradientColor1_Percent(ByVal NewPercent%)
    'If NewRatio# <= 0# Or NewRatio# > 1# Then
    If NewPercent <= 0 Then
        mp_GColor1_Ratio = 0.01
    ElseIf NewPercent > 25500 Then
        mp_GColor1_Ratio = 255#
    Else
        mp_GColor1_Ratio = Round(NewPercent / 100#, 2)
    End If
    
    Call fDrawButton
    
    PropertyChanged "GradientColor1_Percent"
End Property

'   GradientColor2
Public Property Get GradientColor2() As OLE_COLOR
Attribute GradientColor2.VB_Description = "Color #2 for gradient filling"
Attribute GradientColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientColor2 = mp_GColor2
End Property
Public Property Let GradientColor2(ByVal NewColor As OLE_COLOR)
    
    mp_GColor2 = NewColor
    
    If me_GrdBttnColor = gbcOneColor Then
        If fAssignProprtyInDisignMode Then
            MsgBox LoadResString(103), vbInformation, App.ProductName
        End If
    Else
        'Debug.Print "gradientcolor2"
        Call fDrawButton
    End If
    
    PropertyChanged "GradientColor2"
End Property

'   GradientColor2_Percent
Public Property Get GradientColor2_Percent%()
Attribute GradientColor2_Percent.VB_Description = "Set percent of the Gradient Color 2\r\n(less the number - darker the color)"
    GradientColor2_Percent = mp_GColor2_Ratio * 100
End Property
Public Property Let GradientColor2_Percent(ByVal NewPercent%)
    Dim dRatio#
    
    'If NewRatio# <= 0# Or NewRatio# > 1# Then
    If NewPercent <= 0 Then
        mp_GColor2_Ratio = 0.01
    ElseIf NewPercent > 25500 Then
        mp_GColor2_Ratio = 255#
    Else
        mp_GColor2_Ratio = Round(NewPercent / 100#, 2)
    End If
    
    Call fDrawButton
    
    PropertyChanged "GradientColor2_Percent"
End Property

'   BorderColor1
Public Property Get BorderColor1() As OLE_COLOR
Attribute BorderColor1.VB_Description = "Color of left and top borders"
Attribute BorderColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor1 = mp_BColor1
End Property
Public Property Let BorderColor1(ByVal NewColor As OLE_COLOR)
    mp_BColor1 = NewColor
    
    'Debug.Print "border color 1"
    Call fDrawButton
    
    PropertyChanged "BorderColor1"
End Property

'   BorderColor2
Public Property Get BorderColor2() As OLE_COLOR
Attribute BorderColor2.VB_Description = "Color of right and bottom borders"
Attribute BorderColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderColor2 = mp_BColor2
End Property
Public Property Let BorderColor2(ByVal NewColor As OLE_COLOR)
    mp_BColor2 = NewColor
    
    'Debug.Print "border color 2"
    Call fDrawButton
    
    PropertyChanged "BorderColor2"
End Property

'   Angle
Public Property Get Angle() As Single
Attribute Angle.VB_Description = "Gradient filling angle"
Attribute Angle.VB_ProcData.VB_Invoke_Property = "ppGrdButton;Appearance"
    Angle = mp_Angle
End Property
Public Property Let Angle(ByVal NewAngle!)
    Dim l&
    Const conDegreeCircle& = 360&
    
    l = CLng(NewAngle)
    mp_Angle = CSng(((l Mod conDegreeCircle&) + conDegreeCircle&) Mod conDegreeCircle&) + (NewAngle - CSng(l))
    
    'Debug.Print "angle"
    Call fDrawButton
    
    PropertyChanged "Angle"
End Property

'   Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/Set a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = picGradient.Font
End Property
Public Property Set Font(ByVal newFont As Font)
    Set picGradient.Font = newFont
    
    'Debug.Print "font"
    Call fDrawButton
    
    PropertyChanged "Font"
End Property

'   caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "ppGrdButton"
Attribute Caption.VB_UserMemId = -518
    Caption = picGradient.Tag
End Property
Public Property Let Caption(ByVal NewCaption As String)
    Dim k%, D$
    
    Const conAmpersand = "&"
    picGradient.Tag = NewCaption
        
    bDoCaption = CBool(Len(NewCaption))
    
    If bDoCaption Then
        k% = InStr(1, NewCaption, conAmpersand, vbBinaryCompare)
        If (k% > 0) And (k% < Len(NewCaption)) Then
            D$ = Mid$(NewCaption, k% + 1, 1)
            If D$ <> conAmpersand Then
                UserControl.AccessKeys = D$
            End If
        End If
    End If
    'Debug.Print "caption"
    Call fDrawButton
    
    PropertyChanged "Caption"
End Property

'   forecolor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = picImageHold.ForeColor
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    picImageHold.ForeColor = NewColor
    'Debug.Print "forecolor"
    Call fDrawButton
    
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Terminate()
    bInit = True
    Set mc_clsGradient = Nothing
End Sub

'   hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hwnd.VB_UserMemId = -515
Attribute hwnd.VB_MemberFlags = "440"
    hwnd = picGradient.hwnd
End Property

'   hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
Attribute hdc.VB_MemberFlags = "440"
    hdc = picGradient.hdc
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    bWriteProp = True
    With PropBag
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("Caption", Me.Caption, conDefCaption)
        Call .WriteProperty("ForeColor", Me.ForeColor, conDefForeColor)
        Call .WriteProperty("Font", Me.Font, fDefaultFont)
        Call .WriteProperty("Image", Me.Image, Nothing)
        Call .WriteProperty("GColor_Boost0", Me.GColor_Boost0, False)
        Call .WriteProperty("GradientColor1", Me.GradientColor1, conDefGtColor1)
        Call .WriteProperty("GradientColor1_Percent", Me.GradientColor1_Percent, conDefGtRatio1 * 100)
        Call .WriteProperty("GradientColor2", Me.GradientColor2, conDefGtColor2)
        Call .WriteProperty("GradientColor2_Percent", Me.GradientColor2_Percent, conDefGtRatio2 * 100)
        Call .WriteProperty("BorderColor1", Me.BorderColor1, conDefBdColor1)
        Call .WriteProperty("BorderColor2", Me.BorderColor2, conDefBdColor2)
        Call .WriteProperty("Angle", Me.Angle, conDefAngle)
        Call .WriteProperty("ImageCaptionPos", Me.ImageCaptionPos, icpImageCaption_LeftRight)
        Call .WriteProperty("FocusColor", Me.FocusColor, vbBlack)
        Call .WriteProperty("DisabledEmboss", Me.DisabledEmboss, delEmbossNormal)
        Call .WriteProperty("ButtonColors", Me.ButtonColors, gbcGradient)
        Call .WriteProperty("BorderWidth", Me.BorderWidth, gbwBorderWidthNormal)
    End With
    bWriteProp = False
    'Debug.Print "writeP"
    'Call fDrawButton
End Sub

Private Function fDefaultFont() As Font
    On Error Resume Next
    Set fDefaultFont = Ambient.Font
    fDefaultFont.Bold = True
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    bInit = True
    bReadProp = True
    
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        Caption = .ReadProperty("Caption", conDefCaption)
        ForeColor = .ReadProperty("ForeColor", conDefForeColor)
        GColor_Boost0 = .ReadProperty("GColor_Boost0", False)
        GradientColor1 = .ReadProperty("GradientColor1", conDefGtColor1)
        GradientColor1_Percent = .ReadProperty("GradientColor1_Percent", conDefGtRatio1 * 100)
        GradientColor2 = .ReadProperty("GradientColor2", conDefGtColor2)
        GradientColor2_Percent = .ReadProperty("GradientColor2_Percent", conDefGtRatio2 * 100)
        BorderColor1 = .ReadProperty("BorderColor1", conDefBdColor1)
        BorderColor2 = .ReadProperty("BorderColor2", conDefBdColor2)
        Angle = .ReadProperty("Angle", conDefAngle)
        Set Font = .ReadProperty("Font", fDefaultFont)
        Set Image = .ReadProperty("Image", Nothing)
        ImageCaptionPos = .ReadProperty("ImageCaptionPos", icpImageCaption_LeftRight)
        FocusColor = .ReadProperty("FocusColor", vbBlack)
        DisabledEmboss = .ReadProperty("DisabledEmboss", delEmbossNormal)
        ButtonColors = .ReadProperty("ButtonColors", gbcGradient)
        BorderWidth = .ReadProperty("BorderWidth", gbwBorderWidthNormal)
    End With
    
    bInit = False
    bReadProp = False
    'Debug.Print "read P"
    Call fDrawButton
End Sub

Private Function fDefaultCaption$()
    With UserControl.Parent
        fDefaultCaption = .Controls(.Controls.Count - 1).Name
    End With
End Function

Private Sub UserControl_InitProperties()
    bInit = True
    'Debug.Print "UserControl_InitProp start"
    With UserControl.Parent
        Caption = Mid(.Controls(.Controls.Count - 1).Name, 3)
    End With
    Enabled = True
    Angle = conDefAngle
    BorderColor1 = conDefBdColor1
    BorderColor2 = conDefBdColor2
    GradientColor1 = conDefGtColor1
    GradientColor2 = conDefGtColor2
    ForeColor = conDefForeColor
    Set Font = fDefaultFont
    ImageCaptionPos = icpImageCaption_LeftRight
    DisabledEmboss = delEmbossNormal
    BorderWidth = gbwBorderWidthNormal
    
    mp_GColor1_Ratio = conDefGtRatio1
    mp_GColor2_Ratio = conDefGtRatio2
        
    mp_GColor_Boost0 = False
        
    FocusColor = vbBlack
    
    'Debug.Print "UserControl_InitProp end"
    bInit = False
End Sub

Private Sub fSetGradient()
    Dim lBitsPerPixel&, lTempStepX&, lTempStepY&, lTempStep&, lTempDW&
    Dim lShadowColor&, lHighlightColor&
    Dim BC As gtypeRect
    
    lBitsPerPixel& = fGetBitsPerPixel&(UserControl.hdc)
    
    'lTempStepX = Screen.TwipsPerPixelX / me_BorderWidth - 1
    'lTempStepY = Screen.TwipsPerPixelY / me_BorderWidth - 1
    
    If me_BorderWidth = gbwBorderWidthStrong Then
        lTempStep = me_BorderWidth * 2
    Else
        lTempStep = me_BorderWidth - 1
    End If
    On Error GoTo ExitSub
    
    With picGradient
'        BC.Left = Screen.TwipsPerPixelX * me_BorderWidth - 1
'        BC.Top = Screen.TwipsPerPixelY * me_BorderWidth - 1
'        BC.Width = .Width - BC.Left
'        BC.Height = .Height - BC.Top
        
        BC.Left = lTempStep
        BC.Top = lTempStep
        BC.Width = .Width - lTempStep * 3
        BC.Height = .Height - lTempStep * 3
        
        If UserControl.Enabled Then
            If lBitsPerPixel > 8 Then   '   > 256
                If me_GrdBttnColor = gbcGradient Then
                    With mc_clsGradient
                        .Angle = mp_Angle
                        .Color1 = fGetGradientColorRatio(mp_GColor1, mp_GColor1_Ratio, mp_GColor_Boost0)
                        .Color2 = fGetGradientColorRatio(mp_GColor2, mp_GColor2_Ratio, mp_GColor_Boost0)
                        '.Color1 = mp_GColor1
                        '.Color2 = mp_GColor2
                        
                        .Draw picGradient
                        
                        'Call gfDrawGradient(picGradient.hWnd, mp_GColor1, mp_GColor2, CLng(mp_Angle), CLng(mp_Percent))
                    End With
                Else    '   me_GrdBttnColor = gbcOneColor
                    .BackColor = mp_GColor1
                End If
                
                With BC
                    picGradient.Line (.Left, .Top)-(.Width, .Top), mp_BColor1
                    picGradient.Line (.Left, .Left)-(.Left, .Height), mp_BColor1
                    picGradient.Line (.Left, .Height)-(.Width, .Height), mp_BColor2
                    picGradient.Line (.Width, .Top)-(.Width, .Height), mp_BColor2
                End With
            ElseIf lBitsPerPixel = 8 Then '   256
                '.BackColor = fMixTwoColors256&(mp_GColor1, mp_GColor2)
                GoSub DoItFor16ColorsOrDisabled
            Else    '   16 or less
                GoSub DoItFor16ColorsOrDisabled
            End If
        Else
            GoSub DoItFor16ColorsOrDisabled
        End If
        .Refresh
    End With
    
    If bHasFocus Then Call fDrawFocus
ExitSub:
    Exit Sub
    
DoItFor16ColorsOrDisabled:
    lShadowColor = fGetSysColor(COLOR_BTNSHADOW)
    lHighlightColor = fGetSysColor(COLOR_BTNHIGHLIGHT)
    
    With picGradient
        .BackColor = fGetSysColor(COLOR_BTNFACE)
        
        picGradient.Line (BC.Left, BC.Top)-(BC.Width, BC.Top), lHighlightColor&
        picGradient.Line (BC.Left, BC.Left)-(BC.Left, BC.Height), lHighlightColor&
        
        lTempDW = .DrawWidth
        .DrawWidth = lTempDW + 1
        
        picGradient.Line (BC.Left, BC.Height)-(BC.Width, BC.Height), lShadowColor
        picGradient.Line (BC.Width, BC.Top)-(BC.Width, BC.Height), lShadowColor
        
        .DrawWidth = lTempDW
    End With
    
    Return
End Sub

Private Sub fGetCaptionImageSize()
    Const conSpotH = 80 '   twips
    Const conSpotV = 40 '   twips
    
    With picGradient
        If bDoCaption Then
            mtRectCaptn.Width = .TextWidth(.Tag)
            mtRectCaptn.Height = .TextHeight(.Tag)
        Else
            mtRectCaptn.Width = 0
            mtRectCaptn.Height = 0
        End If
    End With
    With picImageHold
        If picImageHold.Picture Then
            mtRectImgTW.Width = .ScaleWidth
            mtRectImgTW.Height = .ScaleHeight
            With picGradient
                If me_ImgCapPos = icpImageCaption_LeftRight Then   '   default
                    If bDoCaption Then
                        mtRectImgTW.Left = (.Width - mtRectCaptn.Width - _
                                            mtRectImgTW.Width - conSpotH) / 2
                    Else
                        mtRectImgTW.Left = (.Width - mtRectImgTW.Width) / 2
                    End If
                    mtRectImgTW.Top = (.Height - mtRectImgTW.Height) / 2
                    
                    mtRectCaptn.Left = mtRectImgTW.Left + mtRectImgTW.Width + conSpotH
                    mtRectCaptn.Top = (.Height - mtRectCaptn.Height) / 2
                ElseIf me_ImgCapPos = icpImageCaption_TopBottom Then
                    If bDoCaption Then
                        mtRectImgTW.Top = (.Height - mtRectCaptn.Height - _
                                            mtRectImgTW.Height - conSpotV) / 2
                    Else
                        mtRectImgTW.Top = (.Height - mtRectImgTW.Height) / 2
                    End If
                    mtRectImgTW.Left = (.Width - mtRectImgTW.Width) / 2
                    
                    mtRectCaptn.Top = mtRectImgTW.Top + mtRectImgTW.Height + conSpotV
                    mtRectCaptn.Left = (.Width - mtRectCaptn.Width) / 2
                ElseIf me_ImgCapPos = icpCaptionImage_LeftRight Then
                    mtRectCaptn.Left = (.Width - mtRectImgTW.Width - _
                                        mtRectCaptn.Width - conSpotH) / 2
                    mtRectCaptn.Top = (.Height - mtRectCaptn.Height) / 2
                    
                    If bDoCaption Then
                        mtRectImgTW.Left = mtRectCaptn.Left + mtRectCaptn.Width + conSpotH
                    Else
                        mtRectImgTW.Left = (.Width - mtRectImgTW.Width) / 2
                    End If
                    mtRectImgTW.Top = (.Height - mtRectImgTW.Height) / 2
                Else    '   icpCaptionImage_TopBottom
                    mtRectCaptn.Top = (.Height - mtRectImgTW.Height - _
                                        mtRectCaptn.Height - conSpotV) / 2
                    mtRectCaptn.Left = (.Width - mtRectCaptn.Width) / 2
                    
                    If bDoCaption Then
                        mtRectImgTW.Top = mtRectCaptn.Top + mtRectCaptn.Height + conSpotV
                    Else
                        mtRectImgTW.Top = (.Width - mtRectImgTW.Width) / 2
                    End If
                    mtRectImgTW.Left = (.Width - mtRectImgTW.Width) / 2
                End If
            End With
            
            mbLoadPicture = True
            
'            With Screen
'                mtRectImgPX.Width = mtRectImgTW.Width / .TwipsPerPixelX
'                mtRectImgPX.Height = mtRectImgTW.Height / .TwipsPerPixelY
'                mtRectImgPX.Left = mtRectImgTW.Left / .TwipsPerPixelX
'                mtRectImgPX.Top = mtRectImgTW.Top / .TwipsPerPixelY
'            End With
        Else
            mtRectImgTW.Width = 0
            mtRectImgTW.Height = 0
            With picGradient
                mtRectCaptn.Left = (.Width - mtRectCaptn.Width) / 2
                mtRectCaptn.Top = (.Height - mtRectCaptn.Height) / 2
            End With
            
            mbLoadPicture = False
        End If
    End With
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Can be used to redraw the button"
    On Error Resume Next
    Call fDrawButton
End Sub

Public Property Get GColor_Boost0() As Boolean
Attribute GColor_Boost0.VB_Description = "Sample of use:\r\n   GradientColor = &h8000C0 (R = 128, G = 0, B = 192)\r\n   GradientColor_Percent = 300%\r\nIf False: ResultColor = &hFF00FF (G remains to be 0)\r\nIf True: ResultColor = &hFF81FF\r\nIn this case G = 129 (SmallestColor(128) * Percent(300%) - 255)"
    GColor_Boost0 = mp_GColor_Boost0
End Property

Public Property Let GColor_Boost0(ByVal bNewValue As Boolean)
    mp_GColor_Boost0 = bNewValue
    
    Call fDrawButton
    
    PropertyChanged "GColor_Boost0"
End Property
