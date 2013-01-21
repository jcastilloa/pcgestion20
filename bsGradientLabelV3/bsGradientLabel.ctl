VERSION 5.00
Begin VB.UserControl bsGradientLabel 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "bsGradientLabel.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "bsGradientLabel.ctx":000C
End
Attribute VB_Name = "bsGradientLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'-------------------------------------
' BadSoft bsGradientLabel control
' Version 3
' ©2002 BadSoft Entertainment, all rights reserved.
'
' WWW: http://www.badsoft.co.uk
'      badhart@hotpop.com
'-------------------------------------

' I've finally decided to give in to consumer demand and make
' this a multiline control. But it'll take some work...

Option Explicit

'Default Property Values:
Const m_def_NonTTError = True
Const m_def_Offset = 6
Const m_def_WordWrap = 0
Const m_def_TextShadowYOffset = 2
Const m_def_TextShadowXOffset = 2
Const m_def_BorderStyle = 0
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_FlatBorderColour = vbBlack
Const m_def_TextShadowColour = vbBlack
Const m_def_TextShadow = False
Const m_def_LabelType = 0
Const m_def_CaptionAlignment = 0
Const m_def_Colour1 = 0
Const m_def_Colour2 = vbBlue
Const m_def_Colour3 = vbYellow
Const m_def_Colour4 = vbRed
Const m_def_CaptionColour = vbWhite
Const m_def_GradientType = 0

'Property Variables:
Dim m_NonTTError As Boolean
Dim m_Offset As Integer
Dim m_WordWrap As Boolean
Dim m_TextShadowYOffset As Integer
Dim m_TextShadowXOffset As Integer
Dim m_BorderStyle As bsBorderStyle
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_FlatBorderColour As OLE_COLOR
Dim m_TextShadowColour As OLE_COLOR
Dim m_TextShadow As Boolean
Dim m_LabelType As bsLabelType
Dim m_CaptionAlignment As bsCaptionAlign
Dim m_Colour1 As OLE_COLOR
Dim m_Colour2 As OLE_COLOR
Dim m_Colour3 As OLE_COLOR
Dim m_Colour4 As OLE_COLOR
Dim m_CaptionColour As OLE_COLOR
Dim m_GradientType As bsGradient
Dim m_Caption As String
Dim m_Fount As Font


' API CALLS
'-------------------------------------
' The star of the show is GradientFillRect, an API call I came
' across in API Guide. It turned out to be a bit hard to use,
' but with someone's help I managed to get it to work. Bloody
' C++ people...
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long


' CONSTANTS
' GradientFillTriangle()
'-------------------------------------
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2

' CreateFontIndirect()
'-------------------------------------
Private Const PROOF_QUALITY = 2
Private Const OUT_TT_ONLY_PRECIS = 7

' DrawText()
'-------------------------------------
Private Const TA_BASELINE = 24
Private Const TA_BOTTOM = 8
Private Const TA_CENTER = 6
Private Const TA_LEFT = 0
Private Const TA_NOUPDATECP = 0
Private Const TA_RIGHT = 2
Private Const TA_TOP = 0
Private Const TA_UPDATECP = 1
Private Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_CALCRECT = &H400
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100

' GetTextMetrics()
'-------------------------------------
Private Const TMPF_TRUETYPE = &H4

' CreateFontIndirect()
'-------------------------------------
Private Const LF_FACESIZE = 32


' TYPES
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFacename(1 To LF_FACESIZE) As Byte
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer   'Red, Green, Blue and Alpha are "unsigned
   Green As Integer 'short", or UShort, variables.
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UpperLeft As Long  'In reality this is a UNSIGNED Long
   LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type Colour
   Red As Long
   Green As Long
   Blue As Long
   Alpha As Long
End Type


' ENUMS
Enum bsCaptionAlign
   [AlignLeft]
   [AlignCentre]
   [AlignRight]
End Enum

Enum bsGradient
   [Horizontal]
   [Vertical]
   [4 Way]
End Enum

Enum bsLabelType
   glHorizontal
   glVertical
End Enum

Enum bsBorderStyle
   [None]
   [Flat]
   [Raised Thin]
   [Raised 3D]
   [Sunken Thin]
   [Sunken 3D]
   [Etched]
   [Bump]
End Enum


' DrawLabel()
' ------------------------------
' This sub draws the background of the label first, then calls
' other routines to do the text and border.

Private Sub DrawLabel()

   ScaleMode = vbPixels
   AutoRedraw = True
   
   Dim vert(4) As TRIVERTEX
   Dim gTRi(1) As GRADIENT_TRIANGLE
   Dim gRect As GRADIENT_RECT
   Dim temp As Colour
   Dim iTemp As Integer
   
   Cls
   
' It would make sense to make the label control taller than it
' is wider if it is vertical, and vice versa. So a check is
' done here.
   If (m_LabelType = glVertical And ScaleWidth > ScaleHeight) Or _
      (m_LabelType = glHorizontal And ScaleHeight > ScaleWidth) _
      Then
      iTemp = UserControl.Width
      UserControl.Width = UserControl.Height
      UserControl.Height = iTemp
      Exit Sub
   End If
   
' Only if we're satisified with the above do we start drawing
' gradients.
   
   Select Case m_GradientType
      Case [4 Way]
         
         vert(0).X = 0
         vert(0).Y = 0
         temp = LongToRGB(m_Colour1)
         vert(0).Red = temp.Red
         vert(0).Green = temp.Green
         vert(0).Blue = temp.Blue
         
         vert(1).X = ScaleWidth
         vert(1).Y = 0
         temp = LongToRGB(m_Colour2)
         vert(1).Red = temp.Red
         vert(1).Green = temp.Green
         vert(1).Blue = temp.Blue
    
         vert(2).X = ScaleWidth
         vert(2).Y = ScaleHeight
         temp = LongToRGB(m_Colour3)
         vert(2).Red = temp.Red
         vert(2).Green = temp.Green
         vert(2).Blue = temp.Blue
         
         vert(3).X = 0
         vert(3).Y = ScaleHeight
         temp = LongToRGB(m_Colour4)
         vert(3).Red = temp.Red
         vert(3).Green = temp.Green
         vert(3).Blue = temp.Blue
         
         gTRi(0).Vertex1 = 0
         gTRi(0).Vertex2 = 1
         gTRi(0).Vertex3 = 2
         
         gTRi(1).Vertex1 = 0
         gTRi(1).Vertex2 = 2
         gTRi(1).Vertex3 = 3
         GradientFillTriangle UserControl.hdc, vert(0), 4, _
            gTRi(0), 2, GRADIENT_FILL_TRIANGLE

      Case Else
      
         temp = LongToRGB(m_Colour1)
         With vert(0)
            .X = 0
            .Y = 0
            .Red = temp.Red
            .Green = temp.Green
            .Blue = temp.Blue
         End With
         
         temp = LongToRGB(m_Colour2)
         With vert(1)
            .X = ScaleWidth
            .Y = ScaleHeight
            .Red = temp.Red
            .Green = temp.Green
            .Blue = temp.Blue
         End With
         
         gRect.UpperLeft = 0
         gRect.LowerRight = 1

         GradientFillRect UserControl.hdc, vert(0), 2, _
            gRect, 1, IIf(m_GradientType = Horizontal, _
            GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
   End Select
   
   
' Draw the text...
   DrawLabelText IIf(m_LabelType = glHorizontal, 0, 90)
   
' ... and the edges (this function determines whether or not
' we need to).
   DrawEdges
   
End Sub


' LongToRGB()
' ------------------------------
' This is my own function for converting colours as Long values
' into red, green and blue values. Of course I needed a bit of
' help, or should I say reminding.

Private Function LongToRGB(ByVal lColour As Long) As Colour
   Dim iTemp As Byte
   
   'Don't forget to convert those system colours...
   TranslateColor lColour, 0, lColour
   
   'Red
   iTemp = CByte(lColour And vbRed)
   LongToRGB.Red = ByteToUShort(iTemp)
   
   'Green
   iTemp = CByte((lColour And vbGreen) / 256)
   LongToRGB.Green = ByteToUShort(iTemp)
   
   'Blue
   iTemp = CByte((lColour And vbBlue) / 65536)
   LongToRGB.Blue = ByteToUShort(iTemp)

End Function


' DrawLabelText()
' ------------------------------
' The bsGradientLabel control was never meant to be a multiline
' label replacement - or a least, I tried to explain to one of
' those Planet Source Coders who don't seem to listen. I saw no
' reason why it should be implemented, as most uses of such a
' control would be single-lined. But I guess it was to prove a
' challenge.

' There's two problems this had to face - firstly only TrueType
' fonts can be rotated by this method. If you try to rotate a
' non-TrueType font, you'll find it won't work. Luckily I
' managed to find a way of detecting if the user has selected a
' TrueType font or not.
' Secondly, rotating multiple lines of text was proved not to be
' as straightforward as rotating single lines. See below...

Private Sub DrawLabelText(ByVal Angle As Integer)

   On Error GoTo GetOut
   Dim F As LOGFONT, hPrevFont As Long, hFont As Long
   Dim lColour As Long
   Dim tmp As RECT
   Dim iFontHeight As Integer
   Dim px As Integer, py As Integer
   Dim i As Integer, N As Integer, MaxLines As Integer
   Dim tmpArray() As Byte
   Dim tmpCaption As String
   Dim MLines() As String
   Dim MLAlign As Long
   Dim RectWidth As Integer
   
' Check for no caption!
   If m_Caption = "" Then Exit Sub
   
' Set up font
' ----------------------------
' To get the height of the font (in pixels) using the
' UserControl's TextHeight method, we need to set the
' UserControl font to the one the user specified.
   UserControl.FontName = m_Fount.Name
   UserControl.FontSize = m_Fount.Size
   
' Font name is converted to a byte array for API reasons. (Null
' termination of the fount name is important.)
   tmpArray = StrConv(m_Fount.Name & vbNullString, _
      vbFromUnicode)
   On Error GoTo 0
   For i = 0 To UBound(tmpArray)
       F.lfFacename(i + 1) = tmpArray(i)
   Next
   
   F.lfEscapement = 10 * Angle
   F.lfHeight = (m_Fount.Size * -20) / Screen.TwipsPerPixelY
   F.lfItalic = m_Fount.Italic
   F.lfUnderline = m_Fount.Underline
   F.lfWeight = IIf(m_Fount.Bold, 700, 0)
   F.lfQuality = PROOF_QUALITY
   
   hFont = CreateFontIndirect(F)
   hPrevFont = SelectObject(UserControl.hdc, hFont)
   
   Select Case m_CaptionAlignment
      Case [AlignCentre]
         SetTextAlign UserControl.hdc, TA_CENTER
      Case [AlignLeft]
         SetTextAlign UserControl.hdc, TA_LEFT
      Case [AlignRight]
         SetTextAlign UserControl.hdc, TA_RIGHT
   End Select
   
' Get text height
'-------------------------------------
' To get the correct height of the fount we can use the DrawText
' API function.

   If m_LabelType = glHorizontal Then
      tmp.Left = m_Offset
      tmp.Right = ScaleWidth
   Else
      tmp.Bottom = ScaleWidth
   End If
   
   DrawText UserControl.hdc, m_Caption, Len(m_Caption), tmp, _
      IIf(m_WordWrap, DT_WORDBREAK, 0) + DT_CALCRECT
   iFontHeight = tmp.Bottom
       
   If m_LabelType = glHorizontal Then
      Select Case m_CaptionAlignment
         Case [AlignLeft]
            CurrentX = m_Offset
         Case [AlignRight]
            CurrentX = ScaleWidth - m_Offset
         Case [AlignCentre]
            CurrentX = ScaleWidth / 2
      End Select
      CurrentY = (ScaleHeight - iFontHeight) / 2
      
   Else
      Select Case m_CaptionAlignment
         Case [AlignLeft]
            CurrentY = ScaleHeight - m_Offset
         Case [AlignRight]
            CurrentY = m_Offset
         Case [AlignCentre]
            CurrentY = ScaleHeight / 2
      End Select
      CurrentX = (ScaleWidth - iFontHeight) / 2
   End If
   
   
' Draw text + text shadows
' -------------------------------------
' We need to use three different methods for drawing the
' text, depending on WordWrap and LabelType.
   
' Our job is made infinitely easy if the label is a non-
' wordwrapped one. We just use a single Print statement,
' regardless of whether it's horizontal or vertical.

' The variables px and py are needed because after each Print
' command the UserControl's CurrentX and CurrentY completely
' reset themselves.

' For a horizontal wordwrapped label, we can use the DrawText
' API call easily.

' But for vertical wordwrapped labels, we have to do the
' word wrapping ourselves! I tried to use the DrawText API call,
' but the lines aligned themselves to the left of the rect and
' consequently drew themselves over each other. So, we go
' through the whole caption and pick out the lines based on
' spaces and carriage returns. This took some doing, so please
' show your appreciation and leave feedback.

   If WordWrap = False Then
      If m_TextShadow = True Then
         px = CurrentX
         py = CurrentY
         CurrentX = CurrentX + m_TextShadowXOffset
         CurrentY = CurrentY + m_TextShadowYOffset
         TranslateColor m_TextShadowColour, 0, lColour
         SetTextColor UserControl.hdc, lColour
         Print m_Caption
         CurrentX = px
         CurrentY = py
      End If
      TranslateColor m_CaptionColour, 0, lColour
      SetTextColor UserControl.hdc, lColour
      Print m_Caption
      
   ElseIf LabelType = glHorizontal Then
      ShiftRect tmp, 0, CurrentY
      Select Case m_CaptionAlignment
         Case AlignLeft
            MLAlign = TA_LEFT
         Case AlignRight
            MLAlign = TA_RIGHT
            ShiftRect tmp, ScaleWidth - m_Offset * 2, 0
         Case AlignCentre
            MLAlign = TA_CENTER
            ShiftRect tmp, ScaleWidth / 2 - m_Offset, 0
      End Select
      If m_TextShadow = True Then
         ShiftRect tmp, m_TextShadowXOffset, m_TextShadowYOffset
         TranslateColor m_TextShadowColour, 0, lColour
         SetTextColor UserControl.hdc, lColour
         DrawText UserControl.hdc, m_Caption, Len(m_Caption), _
            tmp, DT_WORDBREAK + DT_NOCLIP
         ShiftRect tmp, -m_TextShadowXOffset, -m_TextShadowYOffset
      End If
      TranslateColor m_CaptionColour, 0, lColour
      SetTextColor UserControl.hdc, lColour
      DrawText UserControl.hdc, m_Caption, Len(m_Caption), _
         tmp, DT_WORDBREAK + DT_NOCLIP
            
   Else
      RectWidth = ScaleHeight - m_Offset
      tmpCaption = m_Caption
      i = 1
      
      While i < Len(tmpCaption)
         If Mid(tmpCaption, i, 1) = vbCr Then
            N = N + 1
            ReDim Preserve MLines(1 To N)
            MLines(N) = Left(tmpCaption, i - 1)
            tmpCaption = Right(tmpCaption, Len(tmpCaption) - i - 1)
            i = 1
         ElseIf TextWidth(Left(tmpCaption, i)) > RectWidth Then
            Do Until i = 1 Or Mid(tmpCaption, i, 1) = " "
               i = i - 1
            Loop
            N = N + 1
            ReDim Preserve MLines(1 To N)
            MLines(N) = Left(tmpCaption, i - 1)
            tmpCaption = Right(tmpCaption, Len(tmpCaption) - i)
            i = 1
         Else
            i = i + 1
         End If
      Wend
      If Len(tmpCaption) > 0 Then
         N = N + 1
         ReDim Preserve MLines(1 To N)
         MLines(N) = tmpCaption
      End If
           
      N = TextHeight(" ")
      MaxLines = ScaleWidth / N
      MaxLines = IIf(MaxLines > UBound(MLines()), _
         UBound(MLines()), MaxLines)
      
      If m_TextShadow = True Then
         TranslateColor m_TextShadowColour, 0, lColour
         SetTextColor UserControl.hdc, lColour
         px = (ScaleWidth - MaxLines * N) / 2
         For i = 1 To MaxLines
            CurrentX = px + m_TextShadowXOffset
            Select Case m_CaptionAlignment
               Case [AlignCentre]
                  SetTextAlign UserControl.hdc, TA_CENTER
                  CurrentY = ScaleHeight / 2
               Case [AlignLeft]
                  SetTextAlign UserControl.hdc, TA_LEFT
                  CurrentY = ScaleHeight - m_Offset
               Case [AlignRight]
                  SetTextAlign UserControl.hdc, TA_RIGHT
                  CurrentY = m_Offset
            End Select
            CurrentY = CurrentY + m_TextShadowYOffset
            Print MLines(i)
            px = px + N
         Next
      End If
      
      TranslateColor m_CaptionColour, 0, lColour
      SetTextColor UserControl.hdc, lColour
      px = (ScaleWidth - MaxLines * N) / 2
      For i = 1 To MaxLines
         CurrentX = px
         Select Case m_CaptionAlignment
            Case [AlignCentre]
               SetTextAlign UserControl.hdc, TA_CENTER
               CurrentY = ScaleHeight / 2
            Case [AlignLeft]
               SetTextAlign UserControl.hdc, TA_LEFT
               CurrentY = ScaleHeight - m_Offset
            Case [AlignRight]
               SetTextAlign UserControl.hdc, TA_RIGHT
               CurrentY = m_Offset
         End Select
         Print MLines(i)
         px = px + N
      Next
   End If
   
   
' Clean up and restore original fount.
   hFont = SelectObject(UserControl.hdc, hPrevFont)
   DeleteObject hFont
   
   Exit Sub
GetOut:
   Exit Sub

End Sub

' DrawEdges()
' ------------------------------
' Alongside many BadSoft controls, the bsGradientLabel has 7
' colour-customisable edge styles.

Sub DrawEdges()

   Dim lPen As Long
   
   If m_BorderStyle = None Then Exit Sub
   
   Select Case m_BorderStyle
      Case [Flat]
         lPen = CreatePen(0, 0, TranslateColour(m_FlatBorderColour))
         SelectObject UserControl.hdc, lPen
         Rectangle UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight
         DeleteObject lPen
      
      Case [Raised Thin]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         
      Case [Sunken Thin]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
   
      Case [Raised 3D]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hdc, ScaleWidth - 2, 1, 0
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Sunken 3D]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hdc, ScaleWidth - 2, 1, 0
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Etched]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hdc, ScaleWidth - 2, 1, 0
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Bump]
         MoveToEx UserControl.hdc, ScaleWidth, 0, 0
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hdc, ScaleWidth - 2, 1, 0
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, 0
         DeleteObject lPen
   End Select
End Sub

' ShiftRect()
' ------------------------------
' A sub for quickly shifting a rect by a certain amount in
' either direction.

Private Sub ShiftRect(ByRef whichRect As RECT, X As Integer, Y As Integer)
   whichRect.Top = whichRect.Top + Y
   whichRect.Bottom = whichRect.Bottom + Y
   whichRect.Right = whichRect.Right + X
   whichRect.Left = whichRect.Left + X
End Sub

' TranslateColour()
' ------------------------------
' This translates any long value into an RGB colour, for use
' with drawing functions. I object to being forced to use
' American words so I renamed it myself.

Function TranslateColour(lColour As Long) As Long
   TranslateColor lColour, 0, TranslateColour
End Function

' ByteToUShort()
' ------------------------------
' Thanks to a guy who I only know as Ark, from a Visual Basic
' message board, I can use this function to convert byte values
' into unsigned short (ushort) variables. Again, bloody C++
' people...

Private Function ByteToUShort(ByVal bt As Byte) As Integer
   If bt < 128 Then
      ByteToUShort = CInt(CLng("&H" & Hex(bt) & "00"))
   Else
      ByteToUShort = CInt(CLng("&H" & Hex(bt) & "00") - &H10000)
   End If
End Function

' ShowAbout()
' ------------------------------
' A small sub for showing the About screen.

'Sub ShowAbout()
   'frmAbout.Show vbModal
'End Sub


' IsFontTrueType()
' ------------------------------
' At last, a way of telling if a font is TrueType or not. This
' came from James Crowley.

Public Function IsFontTrueType(sFontName As String) As Boolean
    Dim lf As LOGFONT
    Dim tm As TEXTMETRIC
    Dim oldfont As Long, newFont As Long
    Dim tmpArray() As Byte
    Dim dummy As Long
    Dim i As Integer
    
    tmpArray = StrConv(sFontName & vbNullString, vbFromUnicode)
    For i = 0 To UBound(tmpArray)
        lf.lfFacename(i + 1) = tmpArray(i)
    Next
    
    newFont = CreateFontIndirect(lf)
    oldfont = SelectObject(UserControl.hdc, newFont)
    dummy = GetTextMetrics(UserControl.hdc, tm)
    IsFontTrueType = (tm.tmPitchAndFamily And TMPF_TRUETYPE)
    dummy = SelectObject(UserControl.hdc, oldfont)
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get GradientType() As bsGradient
Attribute GradientType.VB_Description = "The direction the gradient follows."
Attribute GradientType.VB_ProcData.VB_Invoke_Property = ";Appearance"
   GradientType = m_GradientType
End Property

Public Property Let GradientType(ByVal New_GradientType As bsGradient)
   m_GradientType = New_GradientType
   PropertyChanged "GradientType"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text the GradientLabel contains."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0

' Fount()
' -----------------------------
' A check is made when setting the fount to see if the user has
' selected a Vertical type label and a non-TrueType font.
Public Property Get Fount() As Font
Attribute Fount.VB_Description = "The fount used by the Caption property."
Attribute Fount.VB_ProcData.VB_Invoke_Property = ";Font"
   Set Fount = m_Fount
End Property

Public Property Set Fount(ByVal New_Fount As Font)
   Set m_Fount = New_Fount

   If m_LabelType = glVertical And IsFontTrueType(New_Fount.Name) = False Then
      If m_NonTTError Then
         MsgBox "The LabelType property can only be Vertical when the Fount is a TrueType fount.", vbExclamation
      End If
      LabelType = glHorizontal
   End If
   
   PropertyChanged "Fount"
   DrawLabel
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_GradientType = m_def_GradientType
   m_Caption = UserControl.Extender.Name
   Set m_Fount = Ambient.Font
   m_CaptionColour = m_def_CaptionColour
   m_Colour1 = m_def_Colour1
   m_Colour2 = m_def_Colour2
   m_Colour3 = m_def_Colour3
   m_Colour4 = m_def_Colour4
   m_LabelType = m_def_LabelType
   m_CaptionAlignment = m_def_CaptionAlignment
   m_BorderStyle = m_def_BorderStyle
   m_HighlightColour = m_def_HighlightColour
   m_HighlightDKColour = m_def_HighlightDKColour
   m_ShadowColour = m_def_ShadowColour
   m_ShadowDKColour = m_def_ShadowDKColour
   m_FlatBorderColour = m_def_FlatBorderColour
   m_TextShadowColour = m_def_TextShadowColour
   m_TextShadow = m_def_TextShadow
   m_TextShadowYOffset = m_def_TextShadowYOffset
   m_TextShadowXOffset = m_def_TextShadowXOffset
   m_WordWrap = m_def_WordWrap
   m_Offset = m_def_Offset
   m_NonTTError = m_def_NonTTError
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
   Set m_Fount = PropBag.ReadProperty("Fount", Ambient.Font)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   m_Colour1 = PropBag.ReadProperty("Colour1", m_def_Colour1)
   m_Colour2 = PropBag.ReadProperty("Colour2", m_def_Colour2)
   m_Colour3 = PropBag.ReadProperty("Colour3", m_def_Colour3)
   m_Colour4 = PropBag.ReadProperty("Colour4", m_def_Colour4)
   m_LabelType = PropBag.ReadProperty("LabelType", m_def_LabelType)
   m_CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", m_def_CaptionAlignment)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDKColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDKColour", m_def_ShadowDKColour)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
   m_TextShadowColour = PropBag.ReadProperty("TextShadowColour", m_def_TextShadowColour)
   m_TextShadow = PropBag.ReadProperty("TextShadow", m_def_TextShadow)
   m_TextShadowYOffset = PropBag.ReadProperty("TextShadowYOffset", m_def_TextShadowYOffset)
   m_TextShadowXOffset = PropBag.ReadProperty("TextShadowXOffset", m_def_TextShadowXOffset)
   m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
   m_Offset = PropBag.ReadProperty("Offset", m_def_Offset)
   m_NonTTError = PropBag.ReadProperty("NonTTError", m_def_NonTTError)
End Sub

Private Sub UserControl_Resize()
   DrawLabel
End Sub

Private Sub UserControl_Show()
   DrawLabel
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
   Call PropBag.WriteProperty("Fount", m_Fount, Ambient.Font)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("Colour1", m_Colour1, m_def_Colour1)
   Call PropBag.WriteProperty("Colour2", m_Colour2, m_def_Colour2)
   Call PropBag.WriteProperty("Colour3", m_Colour3, m_def_Colour3)
   Call PropBag.WriteProperty("Colour4", m_Colour4, m_def_Colour4)
   Call PropBag.WriteProperty("LabelType", m_LabelType, m_def_LabelType)
   Call PropBag.WriteProperty("CaptionAlignment", m_CaptionAlignment, m_def_CaptionAlignment)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDKColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDKColour", m_ShadowDKColour, m_def_ShadowDKColour)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
   Call PropBag.WriteProperty("TextShadowColour", m_TextShadowColour, m_def_TextShadowColour)
   Call PropBag.WriteProperty("TextShadow", m_TextShadow, m_def_TextShadow)
   Call PropBag.WriteProperty("TextShadowYOffset", m_TextShadowYOffset, m_def_TextShadowYOffset)
   Call PropBag.WriteProperty("TextShadowXOffset", m_TextShadowXOffset, m_def_TextShadowXOffset)
   Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
   Call PropBag.WriteProperty("Offset", m_Offset, m_def_Offset)
   Call PropBag.WriteProperty("NonTTError", m_NonTTError, m_def_NonTTError)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the Caption text."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Colours"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour1() As OLE_COLOR
Attribute Colour1.VB_Description = "The first gradient colour."
Attribute Colour1.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour1 = m_Colour1
End Property

Public Property Let Colour1(ByVal New_Colour1 As OLE_COLOR)
   m_Colour1 = New_Colour1
   PropertyChanged "Colour1"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour2() As OLE_COLOR
Attribute Colour2.VB_Description = "The second gradient colour."
Attribute Colour2.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour2 = m_Colour2
End Property

Public Property Let Colour2(ByVal New_Colour2 As OLE_COLOR)
   m_Colour2 = New_Colour2
   PropertyChanged "Colour2"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour3() As OLE_COLOR
Attribute Colour3.VB_Description = "The third gradient colour."
Attribute Colour3.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour3 = m_Colour3
End Property

Public Property Let Colour3(ByVal New_Colour3 As OLE_COLOR)
   m_Colour3 = New_Colour3
   PropertyChanged "Colour3"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour4() As OLE_COLOR
Attribute Colour4.VB_Description = "The fourth gradient colour."
Attribute Colour4.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour4 = m_Colour4
End Property

Public Property Let Colour4(ByVal New_Colour4 As OLE_COLOR)
   m_Colour4 = New_Colour4
   PropertyChanged "Colour4"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get LabelType() As bsLabelType
Attribute LabelType.VB_Description = "The alignment of the Caption."
Attribute LabelType.VB_ProcData.VB_Invoke_Property = ";Appearance"
   LabelType = m_LabelType
End Property

Public Property Let LabelType(ByVal New_LabelType As bsLabelType)
   m_LabelType = New_LabelType
   
   If m_LabelType = glVertical And IsFontTrueType(m_Fount.Name) = False Then
      If m_NonTTError Then
         MsgBox "The LabelType property can only be Vertical when the Fount is a TrueType fount.", vbExclamation
      End If
      m_LabelType = glHorizontal
   End If
   
   PropertyChanged "LabelType"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get CaptionAlignment() As bsCaptionAlign
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionAlignment = m_CaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As bsCaptionAlign)
   m_CaptionAlignment = New_CaptionAlignment
   PropertyChanged "CaptionAlignment"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BorderStyle() As bsBorderStyle
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsBorderStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightColour() As OLE_COLOR
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightDKColour() As OLE_COLOR
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDKColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowColour() As OLE_COLOR
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowDKColour() As OLE_COLOR
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDKColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FlatBorderColour() As OLE_COLOR
   FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
   m_FlatBorderColour = New_FlatBorderColour
   PropertyChanged "FlatBorderColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextShadowColour() As OLE_COLOR
Attribute TextShadowColour.VB_Description = "The colour of the shadow under the text when TextShadow is set to True."
   TextShadowColour = m_TextShadowColour
End Property

Public Property Let TextShadowColour(ByVal New_TextShadowColour As OLE_COLOR)
   m_TextShadowColour = New_TextShadowColour
   PropertyChanged "TextShadowColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get TextShadow() As Boolean
Attribute TextShadow.VB_Description = "Determines whether or not a shadow is drawn under the caption."
   TextShadow = m_TextShadow
End Property

Public Property Let TextShadow(ByVal New_TextShadow As Boolean)
   m_TextShadow = New_TextShadow
   PropertyChanged "TextShadow"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get TextShadowYOffset() As Integer
Attribute TextShadowYOffset.VB_Description = "The distance between the text shadow and the Caption vertically."
   TextShadowYOffset = m_TextShadowYOffset
End Property

Public Property Let TextShadowYOffset(ByVal New_TextShadowYOffset As Integer)
   m_TextShadowYOffset = New_TextShadowYOffset
   PropertyChanged "TextShadowYOffset"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get TextShadowXOffset() As Integer
Attribute TextShadowXOffset.VB_Description = "The distance between the text shadow and the Caption horizontally."
   TextShadowXOffset = m_TextShadowXOffset
End Property

Public Property Let TextShadowXOffset(ByVal New_TextShadowXOffset As Integer)
   m_TextShadowXOffset = New_TextShadowXOffset
   PropertyChanged "TextShadowXOffset"
   DrawLabel
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Enables and disabled multiple label lines."
   WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
   m_WordWrap = New_WordWrap
   PropertyChanged "WordWrap"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,6
Public Property Get offset() As Integer
Attribute offset.VB_Description = "The text offset from the left."
   offset = m_Offset
End Property

Public Property Let offset(ByVal New_Offset As Integer)
   m_Offset = New_Offset
   PropertyChanged "Offset"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get NonTTError() As Boolean
Attribute NonTTError.VB_Description = "Decides whether or not to warn the user that a non-TrueType font cannot be rotated."
   NonTTError = m_NonTTError
End Property

Public Property Let NonTTError(ByVal New_NonTTError As Boolean)
   m_NonTTError = New_NonTTError
   PropertyChanged "NonTTError"
End Property

