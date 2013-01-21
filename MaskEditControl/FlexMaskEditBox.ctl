VERSION 5.00
Begin VB.UserControl FlexMaskEditBox 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   KeyPreview      =   -1  'True
   PaletteMode     =   0  'Halftone
   PropertyPages   =   "FlexMaskEditBox.ctx":0000
   ScaleHeight     =   1410
   ScaleWidth      =   3540
   ToolboxBitmap   =   "FlexMaskEditBox.ctx":003D
   Begin VB.TextBox txtFlex 
      Height          =   684
      IMEMode         =   3  'DISABLE
      Left            =   624
      TabIndex        =   0
      Top             =   336
      Width           =   2412
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "calculator"
      End
      Begin VB.Menu mnuPopUpOwnDef 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "FlexMaskEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12Move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                           and  Douwe Konings dkonings@xs4all.nl

'This sources may be used freely without the intention of commercial distribution.
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

'In other words, when you are developing a program for yourself or for
'a company without selling this product to thirt party's it's allowed to
'use this source code. When you, or the company you work for, sells the
'program then permission is needed!!!!!


Event PopUpItems(MenuItemsArray() As Variant, TextMenu As Boolean)
Event PopUpItemsClick(MenuIndex As Integer)

Event ExitOnArrowKeys(Cancel As Boolean)
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event Resize()
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event LinkOpenFlex(Cancel As Integer)
Event LinkErrorFlex(LinkErr As Integer)
Event LinkCloseFlex()
Event LinkNotifyFlex()

Enum flxOleDragMode
   Manual
   Automatic
End Enum

Enum flxOleDropMode
   NoDropMode
   Manual
   Automatic
End Enum

Enum flxLinkMode
   NoLinkMode
   Automatic
   Manual
   Notify
End Enum

Enum flxFieldType
   AlfaNumericField
   DateField
   NumericField
End Enum

Enum flxAlignMent
   LeftJustify = vbLeftJustify
   RightJustify = vbRightJustify
   Center = vbCenter
End Enum

Enum flxDirect
   Down = 1
   Up = -1
End Enum

Enum flxBorderStyle
   NoBorder
   FixedSingle
End Enum

Enum flxAction
   NoValidPos
   UpperCase
   LowerCase
   NoCase
End Enum

Enum flxDateFormat
   LocalSetting
   ddmmyyyy
   ddyyyymm
   yyyyddmm
   yyyymmdd
   mmddyyyy
   mmyyyydd
   ddmmyy
   ddyymm
   yyddmm
   yymmdd
   mmddyy
   mmyydd
   mmdd
   ddmm
End Enum

Enum flxMask '"&AaCc#9? ."
   [Ampersand (&)] = 1
   UpperA = 2
   LowerA = 3
   UpperC = 4
   LowerC = 5
   [Number Sign (#)] = 6
   [Nine Sign (9)] = 7
   [Question Mark (?)] = 8
   NoPos = 9
   [DecPoint (.)] = 10
End Enum

'Default Property Values:
Const m_def_DecimalSeperator = ""
Const m_def_AutoSelect = False
Const m_def_BorderIfFocus = False
Const m_def_BorderNoFocus = False
Const m_def_Char2ActivateTranNextChar2Upper = ""
Const m_def_InOutDateSeperator = "-"
Const m_def_InOutDateFormat = 0
Const m_def_DateSeperator = "-"
Const m_def_InsertZerosInNumField = True
Const m_def_DateFormat = 0
Const m_def_AutoNextFlexInput = False
Const m_def_ForeColor = &H80000008
Const m_def_ForeColorInActive = 0
Const m_def_BackColorInActive = 0
Const m_def_BackColor = &H80000005
Const m_def_UpAndDownKeys2NextFlexMask = True
Const m_def_CalculatorInMenu = True
Const m_def_HideSelection = 0
Const m_def_Alignment = 0
Const m_def_FormatAlignment = 0
Const m_def_FormatString = ""
Const m_def_DataHasChanged = False
Const m_def_SpecialChars = ""
Const m_def_Century = 0
Const m_def_ExitOnEnter = True
Const m_def_CenturyON = 0
Const m_def_MaskCharInclude = True
Const m_def_BeepOnError = False
Const m_def_PromptInclude = 0
Const m_def_MaskChars = "&AaCc#9?Dd"
Const m_def_Mask = ""
Const m_def_FieldType = 0
Const m_def_PromptChar = "_"
Const m_def_AutoToLastPos = True

Const cnstKEYDecPoint = 190
Const cnstASCiiDecPoint = 46
Const VK_INSERT = &H2D

'Property Variables:
Dim m_DecimalSeperator              As String * 1
Dim m_iDecimalSeperator             As Integer
Dim m_SelLength                     As Integer
Dim m_SelStart                      As Integer
Dim bAutoSelect                     As Boolean
Dim m_AutoSelect                    As Boolean
Dim m_BorderIfFocus                 As Boolean
Dim m_BorderNoFocus                 As Boolean
Dim bCharTran2Upper                 As Boolean
Dim iCharPos                        As Integer
Dim m_Char2ActivateTranNextChar2Upper As String
Dim m_InOutDateSeperator                As String
Dim m_InOutDateFormat               As flxDateFormat
Dim m_DateSeperator                 As String
Dim m_InsertZerosInNumField         As Boolean
Dim m_DateFormat                    As flxDateFormat
Dim m_FormatFont                    As Font
Dim m_AutoNextFlexInput             As Boolean
Dim m_ForeColor                     As OLE_COLOR
Dim m_ForeColorInActive             As OLE_COLOR
Dim m_BackColorInActive             As OLE_COLOR
Dim m_BackColor                     As OLE_COLOR
Dim m_UpAndDownKeys2NextFlexMask    As Boolean
Dim m_CalculatorInMenu              As Boolean
Dim m_HideSelection                 As Boolean
Dim m_Alignment                     As flxAlignMent
Dim m_FormatAlignment               As flxAlignMent
Dim m_FormatString                  As String
Dim iDecPoint                       As Integer
Dim iMaxLen                         As Integer
Dim m_DataHasChanged                As Boolean
Dim m_SpecialChars                  As String
Dim m_Century                       As Integer
Dim m_ExitOnEnter                   As Boolean
Dim m_CenturyON                     As Boolean
Dim m_MaskCharInclude               As Boolean
Dim m_BeepOnError                   As Boolean
Dim m_PromptInclude                 As Boolean
Dim m_mask                          As String
Dim mi_CursorPos                    As Integer
Dim mi_FieldType                    As flxFieldType
Dim mb_AutoToLastPos                As Boolean
Dim ms_Text                         As String
Dim ms_PromptChar                   As String
Dim bMaskUsed                       As Boolean
Dim WhatAction()                    As flxAction
Dim MaskSign()                      As flxMask
Dim ms_TxtCopy                      As String
Dim mi_Asc_ms_PromptChar            As Integer
Dim bMayRaiseChange                 As Boolean
Dim sOldTxt                         As String
Dim bNoEval                         As Boolean
Dim bInsertOn                       As Boolean
Dim bEnterFocus                     As Boolean
Dim sYear                           As String
Dim bDelete                         As Boolean
Dim bFormatView                     As Boolean
Dim lInitStyle                      As Long
Dim bRefresh                        As Boolean

Private Sub BuildMask()
Dim i                   As Integer
Dim i2                  As Integer
Dim bBackslash          As Boolean
Dim s                   As String
Dim eUL                 As flxAction
Dim eLast               As flxMask
Dim iLenMask            As Integer
Dim sDisplayMask        As String
Dim me_NextChar         As flxAction

ms_Text = ""
ms_TxtCopy = ""
iDecPoint = 0
If mi_FieldType = DateField Then txtFlex.MaxLength = 0
eUL = NoCase
me_NextChar = NoCase

For i = 1 To Len(m_mask)
   If InStr("{}<>\", Mid$(m_mask, i, 1)) = 0 Then
      iLenMask = iLenMask + 1
   End If
Next

For i = 1 To Len(m_mask) - 1
   If InStr("\> \< \\ \{ \} ", Mid$(m_mask, i, 2) & " ") Then
      iLenMask = iLenMask + 1
   End If
Next

If m_mask <> "" Then
   ReDim WhatAction(1 To Max(iLenMask, txtFlex.MaxLength)) As flxAction
   ReDim MaskSign(1 To Max(iLenMask, txtFlex.MaxLength)) As flxMask
End If

i2 = 1

For i = 1 To Len(m_mask)
   s = Mid$(m_mask, i, 1)
   If s = ">" And Not bBackslash Then 'UPPERCASE
      eUL = UpperCase  '"U"
   ElseIf s = "}" And Not bBackslash Then
      me_NextChar = UpperCase
   ElseIf s = "{" And Not bBackslash Then
      me_NextChar = LowerCase
   ElseIf s = "<" And Not bBackslash Then  'LOWERCASE
      eUL = LowerCase '"L"
   ElseIf s = "\" And Not bBackslash Then  'Next Char is NO mask Char
      bBackslash = True
   ElseIf InStr(m_def_MaskChars, s) > 0 And Not bBackslash Then
      MaskSign(i2) = InStr(m_def_MaskChars, s)
      ms_Text = ms_Text & ms_PromptChar
      If me_NextChar <> NoCase Then
         WhatAction(i2) = me_NextChar
         me_NextChar = NoCase
      Else
         WhatAction(i2) = eUL
      End If
      i2 = i2 + 1
   ElseIf s = m_DecimalSeperator And Not bBackslash Then  'DECIMAL POINT
      MaskSign(i2) = [DecPoint (.)]
      ms_Text = ms_Text & m_DecimalSeperator
      WhatAction(i2) = NoValidPos
      If mi_FieldType = NumericField And iDecPoint = 0 Then
         iDecPoint = i2
      End If
      i2 = i2 + 1
   Else
      MaskSign(i2) = NoPos
      ms_Text = ms_Text & s
      bBackslash = False
      eUL = NoCase
      WhatAction(i2) = NoValidPos
      i2 = i2 + 1
   End If
Next

If i > 1 And i2 > 1 Then
   If iLenMask < txtFlex.MaxLength Then
      eLast = MaskSign(i2 - 1)
      s = Right$(m_mask, 1)
      For i = i2 To UBound(WhatAction)
         WhatAction(i) = eUL
         MaskSign(i) = eLast
         m_mask = m_mask & s
      Next
      ms_Text = ms_Text & String$(txtFlex.MaxLength - iLenMask, ms_PromptChar)
   End If
   bMaskUsed = True
   iMaxLen = Len(ms_Text)
   txtFlex.MaxLength = iMaxLen
   If Not Ambient.UserMode And mi_FieldType = DateField Then
      If m_CenturyON Then txtFlex.MaxLength = txtFlex.MaxLength + 2
   End If
   ms_TxtCopy = ms_Text
   sDisplayMask = ms_TxtCopy
   s = "&AaCc#9? " & m_DecimalSeperator
   For i = 1 To iMaxLen
      If MaskSign(i) <> NoPos Then
         Mid$(sDisplayMask, i, 1) = Mid$(s, MaskSign(i), 1)
      End If
   Next
   If Ambient.UserMode Then
      txtFlex = ms_Text
      GotoFirstPrompChar
   Else
      txtFlex = sDisplayMask
   End If
ElseIf Not Ambient.UserMode Then
   txtFlex = ""
End If
End Sub

Private Sub DelNul(Optional HasFocus As Boolean, Optional NewText As String)
Dim s          As String
Dim i          As Integer
Dim s2         As String
Dim iPos       As Integer
Dim iLen       As Integer
Dim MinusSign  As Boolean
Dim bNumDetect As Boolean

On Error Resume Next   'If len s < 3  e.g. s = "."

If mi_FieldType = NumericField Then

   iPos = txtFlex.SelStart
   
   If NewText <> "" Then
      s = NewText
   Else
      s = ms_Text
      For i = 1 To iMaxLen
         If WhatAction(i) = NoValidPos And i <> iDecPoint Then
            Mid$(s, i, 1) = ms_PromptChar
         End If
      Next
   End If
   
   If iDecPoint Then
      If InStr(s, m_DecimalSeperator) = 0 Then s = s & m_DecimalSeperator
   End If
   
   iLen = Len(s)
      
   For i = 1 To iLen Step 1
      s2 = Mid$(s, i, 1)
      If s2 <> ms_PromptChar Then
         If s2 = "0" Then
            If MinusSign Then
               If i < iLen Then
                  If Mid$(s, i + 1, 1) <> m_DecimalSeperator Or HasFocus Then
                     Mid$(s, i, 1) = ms_PromptChar
                  End If
               End If
               MinusSign = False
            Else
               Mid$(s, i, 1) = ms_PromptChar
            End If
         ElseIf s2 = "-" Then
            MinusSign = True
         Else
            Exit For
         End If
      End If
   Next
         
   If iDecPoint Then
      For i = iLen To 1 Step -1
         s2 = Mid$(s, i, 1)
         If s2 <> ms_PromptChar Then
            If s2 = "0" Then
               Mid$(s, i, 1) = ms_PromptChar
            Else
               Exit For
            End If
         End If
      Next
   End If
      
   s = Replace(s, ms_PromptChar, "")
      
   If HasFocus Then
      For i = 1 To Len(s)
         If IsNumeric(Mid$(s, i, 1)) Then
            bNumDetect = True
            Exit For
         End If
      Next
      If Not bNumDetect Then s = ""
   Else
      If iDecPoint Then
         If Mid$(s, 1, 1) = m_DecimalSeperator Then
            s = "0" & s
         ElseIf Mid$(s, 1, 2) = "-" & m_DecimalSeperator Then
            s = "-0" & m_DecimalSeperator & Mid$(s, 3)
         End If
      End If
   End If
   
   For i = 1 To iMaxLen Step 1
      If WhatAction(i) <> NoValidPos And i <> iDecPoint Then
         Mid$(ms_Text, i, 1) = ms_PromptChar
      End If
   Next
   
   s2 = Token(s, 0, m_DecimalSeperator)
      
   HideCaret txtFlex.hwnd
   If iDecPoint Then
      For i = 1 To Len(s2)
         txtFlex.SelStart = iDecPoint - 1
         Call InsertCharToLeft(Asc(Mid$(s2, i, 1)))
      Next
   Else
      For i = 1 To Len(s2) Step 1
         txtFlex.SelStart = i - 1
         Call InsertChar(Asc(Mid$(s2, i, 1)))
      Next
   End If
   
   s2 = Token(s, 1, m_DecimalSeperator)
   
   For i = Min(Len(s2), iMaxLen - iDecPoint) To 1 Step -1
      txtFlex.SelStart = iDecPoint
      Call InsertChar(Asc(Mid$(s2, i, 1)))
   Next
   
   If m_InsertZerosInNumField And iDecPoint > 0 Then
      
      For i = iDecPoint To iMaxLen Step 1
         If WhatAction(i) <> NoValidPos And Mid$(ms_Text, i, 1) = ms_PromptChar Then
            Mid$(ms_Text, i, 1) = "0"
         End If
      Next
   
   End If
   
   txtFlex = ms_Text
   txtFlex.SelStart = iPos
   ShowCaret txtFlex.hwnd
   
End If

End Sub

Private Sub FitInMask(ByVal sNewText As String)
Dim i                As Integer
Dim i2               As Integer
Dim s                As String
Dim s2               As String

If m_mask = "" Then
   ms_Text = sNewText
   bMaskUsed = False
   Exit Sub
End If

If sNewText = "" Then
   ms_Text = ms_TxtCopy
   Exit Sub
End If

If mi_FieldType = NumericField Then
   Call DelNul(NewText:=sNewText)
   Exit Sub
End If

For i = 1 To Len(sNewText)
   s = Mid$(sNewText, i, 1)
   For i2 = i2 + 1 To iMaxLen
      If WhatAction(i2) <> NoValidPos Then
         s2 = ValidChar(Asc(s), i2)
         If s2 <> "" Then
            Mid$(ms_Text, i2, 1) = s2
         Else
            i2 = i2 - 1
         End If
         Exit For
      ElseIf Mid$(ms_TxtCopy, i2, 1) = s Then
         Exit For
      End If
   Next
Next

End Sub

Private Function ValidChar(ByVal KeyAscii As Integer, ByVal iPos As Integer) As String
Dim i As Integer
If KeyAscii = mi_Asc_ms_PromptChar Then
   ValidChar = ms_PromptChar
   If InStr(m_Char2ActivateTranNextChar2Upper, ValidChar) Then
      bCharTran2Upper = True
   End If
   Exit Function
ElseIf m_SpecialChars <> "" Then
   If InStr(m_SpecialChars, Chr(KeyAscii)) Then
      ValidChar = Chr(KeyAscii)
      Exit Function
   End If
End If

Select Case MaskSign(iPos)
   Case [Number Sign (#)]
      If KeyAscii > 47 And KeyAscii < 58 Then
         ValidChar = Chr(KeyAscii)
      End If
   Case [Nine Sign (9)]
      If KeyAscii > 47 And KeyAscii < 58 Then
         ValidChar = Chr(KeyAscii)
      ElseIf KeyAscii = 45 Then
         If mi_FieldType = NumericField Then
            If Not bDelete Then
               If iDecPoint > 0 Then
                  If iDecPoint < iPos Then Exit Function
               End If
               For i = 1 To iPos
                  If Mid$(ms_Text, i, 1) <> ms_PromptChar And WhatAction(i) <> NoValidPos Then Exit Function
               Next
            End If
         End If
         ValidChar = Chr(KeyAscii)
      End If
   Case [Question Mark (?)]
      Select Case KeyAscii
         Case 65 To 90, 97 To 122
            ValidChar = xCase(Chr(KeyAscii), WhatAction(iPos))
      End Select
   Case UpperA, UpperC
      Select Case KeyAscii
         Case 65 To 90, 97 To 122, 48 To 57
            ValidChar = xCase(Chr(KeyAscii), WhatAction(iPos))
      End Select
   Case [Ampersand (&)]
      Select Case KeyAscii
         Case 32 To 126, 128 To 255
            ValidChar = xCase(Chr(KeyAscii), WhatAction(iPos))
      End Select
   Case LowerA, LowerC
      Select Case KeyAscii
         Case 65 To 90, 32, 97 To 122, 48 To 57
            ValidChar = xCase(Chr(KeyAscii), WhatAction(iPos))
      End Select
End Select

If bCharTran2Upper Then
   ValidChar = UCase$(ValidChar)
   bCharTran2Upper = InStr(m_Char2ActivateTranNextChar2Upper, ValidChar) <> 0
ElseIf InStr(m_Char2ActivateTranNextChar2Upper, ValidChar) And ValidChar <> "" Then
   bCharTran2Upper = True
End If

End Function

Private Function xCase(ByVal s As String, ByVal xAction As flxAction) As String
Select Case xAction
   Case NoCase
      xCase = s
   Case UpperCase
      xCase = UCase$(s)
   Case LowerCase
      xCase = LCase$(s)
End Select
End Function

Private Function ShiftBackSpace() As Boolean
Dim i       As Integer
Dim iPos    As Integer

iPos = txtFlex.SelStart

For i = iPos To 1 Step -1
   If WhatAction(i) <> NoValidPos Then
      Mid$(ms_Text, i, 1) = ms_PromptChar
      txtFlex = ms_Text
      txtFlex.SelStart = i - 1
      ShiftBackSpace = True
      Exit Function
   End If
Next
End Function

Private Function OverWriteChar(ByVal KeyAscii As Integer) As Boolean
Dim s          As String
Dim i          As Integer
Dim iPos       As Integer

iPos = txtFlex.SelStart + 1

For i = iPos To iMaxLen
   s = ValidChar(KeyAscii, i)
   If s <> "" Then
      Mid$(ms_Text, i, 1) = s
      txtFlex.Text = ms_Text
      txtFlex.SelStart = i
      OverWriteChar = True
      Exit For
   End If
Next
End Function

Private Function InsertCharToLeft(ByVal KeyAscii As Integer) As Boolean
Dim i          As Integer
Dim iPos       As Integer
Dim sTmp       As String
Dim s          As String
Dim s2         As String
Dim s3 As String

iPos = txtFlex.SelStart

For i = iPos + 1 To iMaxLen Step 1
   If i = iDecPoint Or (WhatAction(i) <> NoValidPos And Mid$(ms_Text, i, 1) <> ms_PromptChar) Then
      Exit For
   Else
      iPos = iPos + 1
   End If
Next

For i = iPos To 1 Step -1
   If WhatAction(i) = NoValidPos Then
      iPos = iPos - 1
   Else
      Exit For
   End If
Next

If iPos < 2 Then
   If mi_FieldType = NumericField Then
      If iDecPoint < 3 Then
         If Mid$(ms_Text, 1, 1) = ms_PromptChar Then
            txtFlex.SelStart = 0
         End If
      End If
   End If
   InsertCharToLeft = InsertChar(KeyAscii)
   Exit Function
End If

s2 = ValidChar(KeyAscii, iPos)

If s2 <> "" Then
   sTmp = ms_Text
   For i = iPos To 2 Step -1
      If WhatAction(i) <> NoValidPos And s = "" Then
         s = Mid(ms_Text, i, 1)
      End If
      If WhatAction(i - 1) <> NoValidPos And s <> "" Then
         s3 = ValidChar(Asc(s), i - 1)
         If s3 <> "" Then
            If s = ms_PromptChar Then
               Exit For
            ElseIf i = 2 And Mid$(ms_Text, 1, 1) <> ms_PromptChar Then
               InsertCharToLeft = InsertChar(KeyAscii)
               Exit Function
            Else
               Mid$(sTmp, i - 1, 1) = s
               s = ""
            End If
         End If
      End If
   Next
   Mid$(sTmp, iPos, 1) = s2
   ms_Text = sTmp
   txtFlex.Text = sTmp
   txtFlex.SelStart = iPos
Else
   InsertCharToLeft = InsertChar(KeyAscii)
   Exit Function
End If
InsertCharToLeft = True
End Function

Private Function InsertChar(ByVal KeyAscii As Integer) As Boolean
Dim s          As String
Dim i          As Integer
Dim i2         As Integer
Dim sTmp       As String
Dim s2         As String
Dim iPos       As Integer

iPos = GetValidPosRight(txtFlex.SelStart)

If iPos = txtFlex.SelStart Then Exit Function

s = ValidChar(KeyAscii, iPos)
If s = "" Then Exit Function

sTmp = ms_Text
Mid$(sTmp, iPos, 1) = s
For i = iPos To iMaxLen - 1 Step 1
   s = Mid$(ms_Text, i, 1)
   If s = ms_PromptChar Then
      Exit For
   End If
   For i2 = i + 1 To iMaxLen
      s2 = Mid$(ms_Text, i2, 1)
      If i2 = iMaxLen And ((mi_FieldType = NumericField And InStr("0" & ms_PromptChar, s2) = 0) Or (mi_FieldType <> NumericField And s2 <> ms_PromptChar)) Then
         Exit Function
      ElseIf WhatAction(i2) = NoValidPos Then
         i = i + 1
      Else
         s = ValidChar(Asc(s), i2)
         If s <> "" Then
            Mid$(sTmp, i2, 1) = s
         Else
            Exit Function
         End If
         Exit For
      End If
   Next
Next

ms_Text = sTmp
txtFlex.Text = ms_Text
i2 = GetValidPosRight(iPos)
txtFlex.SelStart = IIf(iPos = i2, iPos, i2 - 1)
InsertChar = True
End Function

Private Sub ToDecPoint(Optional ExitFocus As Boolean)
Call DelNul(ExitFocus)
txtFlex.SelStart = iDecPoint
End Sub

Private Function DelLeftFromCursor() As Boolean
Dim i          As Integer
Dim i2         As Integer
Dim i3         As Integer
Dim iPos       As Integer
Dim s          As String

For i3 = 1 To iMaxLen
   If WhatAction(i3) <> NoValidPos Then Exit For
Next

iPos = txtFlex.SelStart

If iPos < i3 Then Exit Function

For i = iPos To i3 Step -1
   If WhatAction(i) <> NoValidPos Then
      For i2 = i To 2 Step -1
         If WhatAction(i2 - 1) <> NoValidPos Then
            s = ValidChar(Asc(Mid$(ms_Text, i2 - 1, 1)), i)
            If s <> "" Then
               Mid$(ms_Text, i, 1) = s
            Else
               ms_Text = txtFlex
               Exit Function
            End If
            Exit For
         End If
      Next
   End If
Next
Mid$(ms_Text, i3, 1) = ms_PromptChar
txtFlex.Text = ms_Text
txtFlex.SelStart = iPos
DelLeftFromCursor = True
End Function

Private Function Delete() As Boolean
Dim i          As Integer
Dim i2         As Integer
Dim i3         As Integer
Dim iPos       As Integer
Dim s          As String

For i3 = iMaxLen To 1 Step -1
   If WhatAction(i3) <> NoValidPos Then Exit For
Next

iPos = txtFlex.SelStart + 1

If iPos = 0 Or iPos > i3 Then Exit Function

For i = iPos To i3 - 1 Step 1
   If WhatAction(i) <> NoValidPos Then
      For i2 = i To i3 - 1 Step 1
         If WhatAction(i2 + 1) <> NoValidPos Then
            s = ValidChar(Asc(Mid$(ms_Text, i2 + 1, 1)), i)
            If s <> "" Then
               Mid$(ms_Text, i, 1) = s
            Else
               ms_Text = txtFlex
               Exit Function
            End If
            Exit For
         End If
      Next
   End If
Next
Mid$(ms_Text, i3) = ms_PromptChar
txtFlex.Text = ms_Text
txtFlex.SelStart = iPos - 1
Delete = True
End Function

Private Function BackSpace() As Boolean
Dim i          As Integer
Dim i2         As Integer
Dim i3         As Integer
Dim iPos       As Integer
Dim s          As String

iPos = txtFlex.SelStart
For i = iPos To 1 Step -1
   If WhatAction(i) <> NoValidPos Then
      iPos = i
      Exit For
   End If
Next
If i <> iPos Then Exit Function

For i3 = iMaxLen To 1 Step -1
   If WhatAction(i3) <> NoValidPos Then Exit For
Next

If iPos = 0 Or i3 = 0 Then Exit Function

For i = iPos To i3 - 1 Step 1
   If WhatAction(i) <> NoValidPos Then
      For i2 = i To i3 - 1 Step 1
         If WhatAction(i2 + 1) <> NoValidPos Then
            s = ValidChar(Asc(Mid$(ms_Text, i2 + 1, 1)), i)
            If s <> "" Then
               Mid$(ms_Text, i, 1) = s
            Else
               ms_Text = txtFlex
               Exit Function
            End If
            Exit For
         End If
      Next
   End If
Next
Mid$(ms_Text, i3) = ms_PromptChar
txtFlex.Text = ms_Text
txtFlex.SelStart = iPos - 1
BackSpace = True
End Function

Public Function IsEmptyText() As Boolean
Dim i As Integer
Dim s As String

For i = 1 To Len(ms_Text)
   If WhatAction(i) <> NoValidPos Then
      If Mid$(ms_Text, i, 1) <> ms_PromptChar Then
         s = s & Mid$(ms_Text, i, 1)
      End If
   End If
Next

If mi_FieldType = NumericField Then
   IsEmptyText = Val(s) <> 0
Else
   IsEmptyText = (s = "")
End If
End Function

Private Sub txtFlex_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
Dim iSelStart As Integer
bEnterFocus = False
If bMaskUsed Then
   If Shift = vbShiftMask Then
      If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
         Exit Sub
      End If
   ElseIf Shift = vbCtrlMask Then
      If KeyCode = vbKeyC Then
         mnuCopy_click
      ElseIf KeyCode = vbKeyV Then
         mnuPaste_click
      End If
   End If
   If KeyCode = vbKeyInsert Then
      bInsertOn = Not bInsertOn
      Call CreateCaret(txtFlex.hwnd, 0, IIf(bInsertOn, txtFlex.FontSize, 1), txtFlex.FontSize * 2)
      ShowCaret txtFlex.hwnd
   ElseIf KeyCode = vbKeyBack Then
      If txtFlex.Locked Then
         If m_BeepOnError Then Beep
      ElseIf Shift = vbShiftMask Then  'shiftkey pressed
         If Not ShiftBackSpace And m_BeepOnError Then Beep
      Else
         If Not BackSpace And m_BeepOnError Then Beep
      End If
   ElseIf KeyCode = vbKeyDelete Then
      If txtFlex.Locked Then
         If m_BeepOnError Then Beep
      ElseIf txtFlex.SelText <> "" Then
         DeleteSelTxt
      ElseIf mi_FieldType = NumericField And Not Shift = vbCtrlMask Then
         If iDecPoint > 0 And txtFlex.SelStart < iDecPoint Then
            bDelete = True
            If Not Shift = vbShiftMask Then
               txtFlex.SelStart = Min(txtFlex.SelStart + 1, iDecPoint - 1)
            End If
            If Not DelLeftFromCursor And m_BeepOnError Then Beep
            bDelete = False
         Else
            If Not Delete And m_BeepOnError Then Beep
         End If
      ElseIf Shift = vbShiftMask Then
         If Not DelLeftFromCursor And m_BeepOnError Then Beep
      ElseIf Shift = vbCtrlMask Then   'ctrl key pressed
         iSelStart = txtFlex.SelStart
         If iSelStart < iMaxLen Then
            Mid$(ms_Text, iSelStart + 1) = Mid$(ms_TxtCopy, iSelStart + 1)
            txtFlex = ms_Text
            txtFlex.SelStart = iSelStart
         ElseIf m_BeepOnError Then
            Beep
         End If
      Else
         If Not Delete And m_BeepOnError Then Beep
      End If
   ElseIf KeyCode = vbKeyLeft Then
      txtFlex.SelStart = GetValidPosLeft(txtFlex.SelStart + 1) - 1
   ElseIf KeyCode = vbKeyRight Then
      txtFlex.SelStart = GetValidPosRight(txtFlex.SelStart)
   ElseIf (KeyCode = cnstKEYDecPoint Or KeyCode = vbKeyDecimal) And (mi_FieldType = NumericField And iDecPoint > 0) Then ' . pressed
      ToDecPoint
   ElseIf KeyCode = vbKeyHome Then
      txtFlex.SelStart = 0
   ElseIf KeyCode = vbKeyEnd Then
      txtFlex.SelStart = iMaxLen
   ElseIf KeyCode = vbKeyDown And m_UpAndDownKeys2NextFlexMask Then
      FocusNextMask Down
   ElseIf KeyCode = vbKeyUp And m_UpAndDownKeys2NextFlexMask Then
      FocusNextMask Up
   End If
   KeyCode = 0
End If
End Sub

Private Function GetValidPosLeft(ByVal iOldPos As Integer) As Integer
Dim i As Integer
For i = iOldPos - 1 To 1 Step -1
   If WhatAction(i) <> NoValidPos Then
      GetValidPosLeft = i
      Exit Function
   End If
Next
GetValidPosLeft = iOldPos
End Function

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

If xObject Is Nothing Then Exit Sub
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
         If m_ExitOnEnter Or Not txtFlex.Enabled Then
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
'For Each xObject In Extender.Parent
'   If TypeOf xObject Is FlexMaskEditBox Then
'      If xObject.TabIndex = iNextTabIndex Then xObject.SetFocus
'      End If
'   End If
'Next

If ObjHwnds.Count > 0 Then
   RaiseEvent ExitOnArrowKeys(Cancel)
   If Not Cancel Then
      l = ObjHwnds.Item(CStr(iNextTabIndex))
      Set ObjHwnds = Nothing
      Set ObjTabIndex = Nothing
      Set xObject = Nothing
      SetFocusAPI l
   End If
End If

Set ObjHwnds = Nothing
Set ObjTabIndex = Nothing
Set xObject = Nothing
End Sub

Private Function IsAnyValidNextPos() As Boolean
Dim i As Integer
If iMaxLen > txtFlex.SelStart Then
   For i = txtFlex.SelStart + 1 To iMaxLen Step 1
      If WhatAction(i) <> NoValidPos Then
         IsAnyValidNextPos = True
         Exit Function
      End If
   Next
End If
End Function

Private Sub txtFlex_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
If KeyAscii = vbKeyReturn And (m_UpAndDownKeys2NextFlexMask Or m_ExitOnEnter) Then
   Call FocusNextMask(Down, True)
   KeyAscii = 0
ElseIf bMaskUsed Then
   If KeyAscii >= vbKeySpace Then
      If txtFlex.Locked Then
         If m_BeepOnError Then Beep
         KeyAscii = 0
         Exit Sub
      ElseIf txtFlex.SelLength > 0 Then
         DeleteSelTxt
      End If
      If KeyAscii = 45 And InStr(ms_Text, "-") > 0 And mi_FieldType = NumericField Then
         If m_BeepOnError Then Beep
      ElseIf KeyAscii > 47 And KeyAscii < 58 And mi_FieldType = NumericField And txtFlex.SelStart < InStr(ms_Text, "-") Then
         If m_BeepOnError Then Beep
      ElseIf (KeyAscii = m_iDecimalSeperator) And (mi_FieldType = NumericField And iDecPoint > 0) Then  ' . pressed
         ToDecPoint
       ElseIf Not (KeyAscii = cnstASCiiDecPoint And (mi_FieldType = NumericField And iDecPoint > 0)) Then
         If bInsertOn Then
            If Not OverWriteChar(KeyAscii) And m_BeepOnError Then Beep
         Else
            If mi_FieldType = NumericField And txtFlex.SelStart < iDecPoint Then
               If Not InsertCharToLeft(KeyAscii) And m_BeepOnError Then Beep
            Else
               If InsertChar(KeyAscii) Then
                  If m_AutoNextFlexInput Then
                     If Not IsAnyValidNextPos Then
                        Call FocusNextMask(Down, True)
                     End If
                  End If
               ElseIf m_BeepOnError Then
                  Beep
               End If
            End If
         End If
      End If
   End If
   KeyAscii = 0
End If
If iCharPos > 0 Then
   iCharPos = 0
   bCharTran2Upper = False
ElseIf bCharTran2Upper Then
   iCharPos = txtFlex.SelStart
End If
End Sub

Private Sub txtFlex_LinkNotify()
RaiseEvent LinkNotifyFlex
End Sub

Private Sub UserControl_EnterFocus()

If bRefresh Then Refresh

If Not txtFlex.Enabled Then
   Call FocusNextMask(Down)
End If

bCharTran2Upper = False
Set txtFlex.Font = UserControl.Font
bEnterFocus = True
bInsertOn = (GetKeyState(VK_INSERT) = 1)
txtFlex.BackColor = m_BackColor
txtFlex.ForeColor = m_ForeColor

   
If m_FormatString <> "" Then
   txtFlex.MaxLength = iMaxLen
   txtFlex.Alignment = m_Alignment
   txtFlex.Text = ms_Text
   Call DelNul(True)
   bMayRaiseChange = True
   If m_SelLength = 0 Then
      If mi_CursorPos = -1 Then
         GotoFirstPrompChar
      Else
         txtFlex.SelStart = mi_CursorPos
      End If
   End If
Else
   Call DelNul(True)
   If mb_AutoToLastPos And bMaskUsed Then
      GotoFirstPrompChar
   End If
End If
Call CreateCaret(txtFlex.hwnd, 0, IIf(bInsertOn, txtFlex.FontSize, 1), txtFlex.FontSize * 2)
ShowCaret txtFlex.hwnd
Border = m_BorderIfFocus
If m_SelLength > 0 Then
   txtFlex.SelStart = m_SelStart
   txtFlex.SelLength = m_SelLength
   m_SelStart = 0
   m_SelLength = 0
ElseIf m_AutoSelect And bAutoSelect Then
   txtFlex.SelStart = 0
   txtFlex.SelLength = iMaxLen
   bAutoSelect = False
End If
End Sub

Public Property Get Border() As Boolean
Attribute Border.VB_MemberFlags = "40"
Border = GetWindowLong(txtFlex.hwnd, GWL_STYLE) <> lInitStyle
End Property

Public Property Let Border(bNewBorder As Boolean)
Dim lNewStyle As Long
lNewStyle = GetWindowLong(txtFlex.hwnd, GWL_STYLE)
If bNewBorder Then
   If lNewStyle <> lNewStyle Or WS_THICKFRAME Then
      Call SetWindowLong(txtFlex.hwnd, GWL_STYLE, lNewStyle Or WS_THICKFRAME)
      Call SetWindowPos(txtFlex.hwnd, hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If
Else
   If lNewStyle <> lInitStyle Then
      Call SetWindowLong(txtFlex.hwnd, GWL_STYLE, lInitStyle)
      Call SetWindowPos(txtFlex.hwnd, hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If
End If
End Property

Private Sub UserControl_ExitFocus()
Dim sTmp As String
On Error Resume Next   'if the format string fails

If bNoEval Then
   bNoEval = False
   Exit Sub
End If

If bMaskUsed Then
   If mi_FieldType = NumericField Then
      Call DelNul(False)
      If iDecPoint Then
         txtFlex.SelStart = iDecPoint - 1
      End If
   End If
End If

Call FormatView
Border = m_BorderNoFocus
End Sub

Private Sub FormatView()
If m_ForeColorInActive <> 0 Then
   txtFlex.ForeColor = m_ForeColorInActive
End If
If m_BackColorInActive <> 0 Then
   txtFlex.BackColor = m_BackColorInActive
End If

If m_FormatString <> "" Then
   If Not m_FormatFont Is Nothing Then
      Set txtFlex.Font = m_FormatFont
   End If
   mi_CursorPos = txtFlex.SelStart
   bMayRaiseChange = False
   txtFlex.MaxLength = 0
   txtFlex.Alignment = m_FormatAlignment
   If mi_FieldType = NumericField Then
      txtFlex = Format$(NumText, m_FormatString)
   Else
      txtFlex = Format$(Text, m_FormatString)
   End If
   bMayRaiseChange = True
End If
End Sub

Public Function NumText() As String
Dim i As Integer
Dim s As String
s = ms_Text
For i = 1 To iMaxLen
   If WhatAction(i) = NoValidPos And i <> iDecPoint Then
      Mid$(s, i, 1) = ms_PromptChar
   End If
Next
s = Replace(s, ms_PromptChar, "")
s = Replace(s, m_DecimalSeperator, LocalDecimalSeperator)
If IsNumeric(s) Then
   NumText = s
End If
End Function

Private Sub UserControl_Initialize()
Dim i As Integer
Dim l As Long
sDateFormats(1) = LocalDate()
If sDateFormats(2) = "" Then
   For i = 0 To 13 Step 1
      sDateFormats(i + 2) = Token("ddmmyyyy^ddyyyymm^yyyyddmm^yyyymmdd^mmddyyyy^mmyyyydd^ddmmyy^ddyymm^yyddmm^yymmdd^mmddyy^mmyydd^mmdd^ddmm", 0, , l)
   Next
End If
bMayRaiseChange = True
mi_CursorPos = -1
End Sub

Private Sub UserControl_Resize()
txtFlex.Move 0, 0, ScaleWidth, ScaleHeight
RaiseEvent Resize
End Sub

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "Characteristics"
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "103c"
Dim sTmpStr    As String
Dim i          As Integer
Dim s          As String
Dim bIsDate    As Boolean

If Not Ambient.UserMode Then   'design mode
   Text = ms_Text
   Exit Property
End If

s = ms_Text
If bMaskUsed Then
   If mi_FieldType = DateField Then
      If ValidDate Then
         s = TransFormDate(s, m_DateFormat, m_InOutDateFormat, , m_InOutDateSeperator)
         bIsDate = True
      End If
   End If
   If Not m_MaskCharInclude Then
      If mi_FieldType = DateField And bIsDate Then
         sTmpStr = Replace(s, m_InOutDateSeperator, "")
      Else
         For i = 1 To iMaxLen
            If Not WhatAction(i) = NoValidPos Then
               sTmpStr = sTmpStr & Mid$(s, i, 1)
            End If
         Next
      End If
      If Not m_PromptInclude Then
         sTmpStr = Replace(sTmpStr, ms_PromptChar, " ")
      End If
      Text = sTmpStr
   ElseIf Not m_PromptInclude Then
      Text = Replace(s, ms_PromptChar, " ")
   Else
      Text = s
   End If
Else
   Text = txtFlex
End If
End Property

Public Property Let Text(ByVal New_Text As String)
Dim i                As Integer
Dim i2               As Integer
Dim sTmp             As String
Dim sTmp2            As String
Dim sTmpNewText      As String

If bRefresh Then Refresh

If ms_Text <> "" Then
   ms_Text = ms_TxtCopy
Else
   BuildMask
End If

If Not Ambient.UserMode Then 'design mode
   ms_Text = New_Text
   Exit Property
End If

If mi_FieldType = NumericField Then
   New_Text = Replace(New_Text, ".", m_DecimalSeperator)
   New_Text = Replace(New_Text, LocalDecimalSeperator, m_DecimalSeperator)
   sTmpNewText = Replace(New_Text, ms_PromptChar, "")
   If IsNumeric("0" & sTmpNewText) Then
      If iDecPoint = 0 And InStr(sTmpNewText, m_DecimalSeperator) > 0 Then
         sTmpNewText = Trim$(Mid$(sTmpNewText, 1, InStr(sTmpNewText, m_DecimalSeperator)))
      End If
      If Val(sTmpNewText) = 0 And Len(sTmpNewText) > 0 Then sTmpNewText = "0"
   End If
ElseIf mi_FieldType = DateField Then
   'if no seperator in text f.i. "030458"
   If m_InOutDateFormat = LocalSetting Then
      New_Text = Format(New_Text, sDateFormats(1))
   End If
   If InStr(New_Text, m_InOutDateSeperator) = 0 And Len(New_Text) > 0 Then
      sTmp = sDateFormats(m_InOutDateFormat + 1)
      If m_InOutDateFormat = LocalSetting Then sTmp = Replace(sTmp, LocalDateSeperator, "")
      sTmp2 = Left$(New_Text, 1)
      For i = 2 To Len(New_Text)
         If i <= Len(sTmp) Then
            If Mid$(sTmp, i, 1) <> Mid$(sTmp, i - 1, 1) Then
               sTmp2 = sTmp2 & m_InOutDateSeperator
            End If
         End If
         sTmp2 = sTmp2 & Mid$(New_Text, i, 1)
      Next
      New_Text = sTmp2
   End If
   sTmpNewText = TransFormDate(New_Text, m_InOutDateFormat, m_DateFormat, m_InOutDateSeperator, m_DateSeperator)
Else
   sTmpNewText = New_Text
End If

If CanPropertyChange("Text") Then
   If bMaskUsed Then
      If mi_FieldType = DateField Then
         If ValidDate(sTmpNewText) Then
            For i = 1 To Len(sTmpNewText)
               sTmp = Mid$(sTmpNewText, i, 1)
               If WhatAction(i) <> NoValidPos And IsNumeric(sTmp) Then
                  Mid$(ms_Text, i, 1) = sTmp
               End If
            Next
            If InStr(UCase$(sDateFormats(m_DateFormat + 1)), "YYYY") Then
               If ms_Text <> sTmpNewText Then
                  ms_Text = ms_TxtCopy
               End If
            End If
         End If
      Else
         FitInMask sTmpNewText
      End If
      txtFlex.Text = ms_Text
   Else
      txtFlex.Text = sTmpNewText
   End If
   sOldTxt = txtFlex
   If mb_AutoToLastPos And bMaskUsed Then
      GotoFirstPrompChar
   Else
      txtFlex.SelStart = 0
   End If
   m_DataHasChanged = False
   PropertyChanged "Text"
End If
If Not bFormatView Then Call FormatView
bAutoSelect = True
End Property

Private Sub GotoFirstPrompChar()
Dim i As Integer
If mi_FieldType = NumericField And iDecPoint > 0 Then
   txtFlex.SelStart = iDecPoint - 1
Else
   i = InStr(ms_Text, ms_PromptChar)
   If i = 0 Then
      i = InStr(ms_TxtCopy, ms_PromptChar)
   End If
   txtFlex.SelStart = Max(i - 1, 0)
End If
End Sub

Public Property Get Mask() As String
Mask = m_mask
End Property

Public Property Let Mask(ByVal New_Mask As String)
On Error GoTo ExitOnError
Dim sYear      As String
Dim i          As Integer
Dim sTmp       As String
Dim sTmp2      As String
Dim sSep       As String
If mi_FieldType = DateField Then
   New_Mask = ""
   sTmp2 = InsSepInDateformat(m_DateFormat)
   sSep = IIf(m_DateFormat = LocalSetting, LocalDateSeperator, m_DateSeperator)
   If m_CenturyON Then
      sYear = IIf(Val(m_Century) <> 0, Format$(m_Century, "##00"), Left$(CStr(Year(Date)), 2))
      For i = 0 To 2
         sTmp = UCase$(Token(sTmp2, i, sSep))
         If sTmp = "YYYY" Then
            New_Mask = New_Mask & IIf(i, sSep, "") & "\" & Left$(sYear, 1) & "\" & Right$(sYear, 1) & "##"
         ElseIf sTmp <> "" Then
            New_Mask = New_Mask & IIf(i, sSep, "") & "##"
         End If
      Next
   Else
      For i = 0 To 2
         sTmp = Token(sTmp2, i, sSep)
         If sTmp <> "" Then
            New_Mask = New_Mask & IIf(i, sSep, "") & IIf(Len(sTmp) = 4, "####", "##")
         End If
      Next
   End If
   m_mask = New_Mask
   Call BuildMask
   bRefresh = False
   If Not Ambient.UserMode Then 'design mode
      txtFlex.Text = m_mask
   End If
ElseIf New_Mask = "" Then
   m_mask = ""
   bMaskUsed = False
Else
   m_mask = New_Mask
   Call BuildMask
   bRefresh = False
End If
PropertyChanged "Mask"
Exit Property
ExitOnError:
m_mask = ""
bMaskUsed = False
End Property

Public Property Let CenturyON(ByVal New_CenturyON As Boolean)
Attribute CenturyON.VB_Description = "Eeuw Aanduiding In Datum"
m_CenturyON = New_CenturyON
bRefresh = True
PropertyChanged "CenturyON"
End Property

Private Function GetValidPosRight(ByVal iPos As Integer) As Integer
Dim i As Integer
For i = iPos + 1 To iMaxLen
   If WhatAction(i) <> NoValidPos Then
      GetValidPosRight = i
      Exit Function
   End If
Next
GetValidPosRight = iPos
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
txtFlex.Enabled() = New_Enabled
UserControl.Enabled = New_Enabled
PropertyChanged "Enabled"
End Property

Public Property Get BorderStyle() As flxBorderStyle
BorderStyle = txtFlex.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As flxBorderStyle)
txtFlex.BorderStyle() = New_BorderStyle
PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
bRefresh = False
InOutDateSeperator = m_InOutDateSeperator
Mask = m_mask
txtFlex.ForeColor = m_ForeColor
txtFlex.BackColor = m_BackColor
txtFlex.Refresh
End Sub

Private Sub txtFlex_Click()
RaiseEvent Click
End Sub

Private Sub txtFlex_DblClick()
RaiseEvent DblClick
End Sub

Public Property Get FieldType() As flxFieldType
Attribute FieldType.VB_Description = "Invoer Type Aanduiding (0) AlfaNumeriek (2) numeriek (1) Datum"
Attribute FieldType.VB_ProcData.VB_Invoke_Property = "Eigenschappen"
FieldType = mi_FieldType
End Property

Public Property Let FieldType(ByVal New_FieldType As flxFieldType)
mi_FieldType = New_FieldType
bRefresh = True
If New_FieldType = DateField Then
   ms_Text = ""
End If
PropertyChanged "FieldType"
End Property

Public Property Get PromptChar() As String
Attribute PromptChar.VB_Description = "Opvul Karakter (VERPLICHT)"
Attribute PromptChar.VB_ProcData.VB_Invoke_Property = "Characteristics"
PromptChar = ms_PromptChar
End Property

Public Property Let PromptChar(ByVal New_PromptChar As String)
ms_PromptChar = IIf(Len(New_PromptChar) <> 1, "_", New_PromptChar)
mi_Asc_ms_PromptChar = Asc(ms_PromptChar)
bRefresh = True
PropertyChanged "PromptChar"
End Property

Public Property Get AutoToLastPos() As Boolean
Attribute AutoToLastPos.VB_Description = "Auto to the last char in text box."
Attribute AutoToLastPos.VB_ProcData.VB_Invoke_Property = "Characteristics"
AutoToLastPos = mb_AutoToLastPos
End Property

Public Property Let AutoToLastPos(ByVal New_AutoToLastPos As Boolean)
mb_AutoToLastPos = New_AutoToLastPos
PropertyChanged "AutoToLastPos"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

mi_FieldType = m_def_FieldType
ms_PromptChar = m_def_PromptChar
mb_AutoToLastPos = m_def_AutoToLastPos
m_mask = m_def_Mask
m_PromptInclude = m_def_PromptInclude
m_BeepOnError = m_def_BeepOnError
m_MaskCharInclude = m_def_MaskCharInclude
m_CenturyON = m_def_CenturyON
m_ExitOnEnter = m_def_ExitOnEnter
m_Century = m_def_Century
m_SpecialChars = m_def_SpecialChars
m_DataHasChanged = m_def_DataHasChanged
m_FormatString = m_def_FormatString
m_FormatAlignment = m_def_FormatAlignment
m_Alignment = m_def_Alignment
m_HideSelection = m_def_HideSelection
m_CalculatorInMenu = m_def_CalculatorInMenu
m_UpAndDownKeys2NextFlexMask = m_def_UpAndDownKeys2NextFlexMask
m_ForeColor = m_def_ForeColor
m_ForeColorInActive = m_def_ForeColorInActive
m_BackColorInActive = m_def_BackColorInActive
m_BackColor = m_def_BackColor
m_AutoNextFlexInput = m_def_AutoNextFlexInput
Set m_FormatFont = Ambient.Font
m_InsertZerosInNumField = m_def_InsertZerosInNumField
m_DateFormat = m_def_DateFormat
m_DateSeperator = m_def_DateSeperator
m_InOutDateFormat = m_def_InOutDateFormat
m_InOutDateSeperator = m_def_InOutDateSeperator
m_Char2ActivateTranNextChar2Upper = m_def_Char2ActivateTranNextChar2Upper
m_BorderIfFocus = m_def_BorderIfFocus
m_BorderNoFocus = m_def_BorderNoFocus
m_AutoSelect = m_def_AutoSelect
m_DecimalSeperator = LocalDecimalSeperator
End Sub

Private Sub txtFlex_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Function aUbound(aArray As Variant) As Long
aUbound = -1
On Error GoTo EndOnError
aUbound = UBound(aArray)
EndOnError:
End Function

Private Sub txtFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemsArray()  As Variant
Dim bTextMenu     As Boolean
Dim i             As Integer

RaiseEvent MouseDown(Button, Shift, X, Y)

If Button = vbLeftButton And mb_AutoToLastPos And bEnterFocus Then
   If bMaskUsed Then
      i = InStr(ms_Text, ms_PromptChar)
      If i = 0 Then
         i = InStr(ms_TxtCopy, ms_PromptChar)
      End If
      txtFlex.SelStart = GetValidPosRight(IIf(txtFlex.SelStart > i, i - 1, txtFlex.SelStart)) - 1
   End If
ElseIf Button = vbRightButton Then
   LockWindowUpdate txtFlex.hwnd
   txtFlex.Enabled = False

   bTextMenu = True
   RaiseEvent PopUpItems(ItemsArray, bTextMenu)

   If aUbound(ItemsArray) <> -1 Then
      mnuPopUpOwnDef(0).Visible = True
      For i = 1 To mnuPopUpOwnDef.Count - 1 Step 1
         Unload mnuPopUpOwnDef(i)
      Next
      For i = 0 To aUbound(ItemsArray) Step 1
         If i > 0 Then Load mnuPopUpOwnDef(i)
         mnuPopUpOwnDef(i).Caption = ItemsArray(i)
      Next
   Else
      bTextMenu = True
      mnuPopUpOwnDef(0).Visible = False
   End If

   mnuCopy.Visible = bTextMenu
   mnuCopy.Enabled = txtFlex.SelText <> ""

   mnuPaste.Visible = bTextMenu
   mnuPaste.Enabled = Clipboard.GetFormat(vbCFText)

   mnuDelete.Visible = bTextMenu
   mnuDelete.Enabled = txtFlex.SelText <> ""

   mnuUndo.Visible = bTextMenu
   mnuUndo.Enabled = ms_Text <> sOldTxt

   mnuLine.Visible = m_CalculatorInMenu And bTextMenu

   mnuCalculator.Visible = m_CalculatorInMenu

   PopupMenu mnuPopUp

   txtFlex.Enabled = True
   LockWindowUpdate 0&
   txtFlex.SetFocus
End If
bEnterFocus = False
End Sub

Private Sub mnuPopUpOwnDef_Click(Index As Integer)
RaiseEvent PopUpItemsClick(Index)
End Sub

Private Sub mnuUndo_click()
bFormatView = True
Text = sOldTxt
bFormatView = False
End Sub

Private Sub mnuCalculator_click()
Dim sTmp As String

txtFlex.Enabled = True
LockWindowUpdate 0&
sTmp = Calculator()

If IsNumeric(sTmp) And Not txtFlex.Locked Then Text = sTmp

End Sub

Private Sub mnuCopy_click()
Clipboard.SetText IIf(m_PromptInclude, txtFlex.SelText, Replace(txtFlex.SelText, ms_PromptChar, "")), vbCFText
End Sub

Private Sub mnuPaste_click()
Dim s As String
Dim i As Integer
If txtFlex.Locked Then Exit Sub
s = Clipboard.GetText
Call DeleteSelTxt
For i = 1 To Len(s)
   Call InsertChar(Asc(Mid$(s, i, 1)))
Next
End Sub

Private Sub mnuDelete_click()
DeleteSelTxt
End Sub

Private Sub txtFlex_LinkError(LinkErr As Integer)
RaiseEvent LinkErrorFlex(LinkErr)
End Sub

Private Sub txtFlex_LinkOpen(Cancel As Integer)
RaiseEvent LinkOpenFlex(Cancel)
End Sub

Private Sub txtFlex_LinkClose()
RaiseEvent LinkCloseFlex
End Sub

Private Sub txtFlex_Change()
If bMayRaiseChange Then RaiseEvent Change
End Sub

Private Sub txtFlex_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txtFlex_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txtFlex_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub txtFlex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtFlex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub LinkExecute(ByVal Command As String)
txtFlex.LinkExecute Command
End Sub

Public Property Get LinkItem() As String
LinkItem = txtFlex.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
txtFlex.LinkItem() = New_LinkItem
PropertyChanged "LinkItem"
End Property

Public Property Get LinkMode() As flxLinkMode
LinkMode = txtFlex.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As flxLinkMode)
txtFlex.LinkMode() = New_LinkMode
PropertyChanged "LinkMode"
End Property

Public Sub LinkPoke()
txtFlex.LinkPoke
End Sub

Public Sub LinkRequest()
txtFlex.LinkRequest
End Sub

Public Sub LinkSend()
txtFlex.LinkSend
End Sub

Public Property Get LinkTimeout() As Integer
LinkTimeout = txtFlex.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
txtFlex.LinkTimeout() = New_LinkTimeout
PropertyChanged "LinkTimeout"
End Property

Public Property Get LinkTopic() As String
LinkTopic = txtFlex.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
txtFlex.LinkTopic() = New_LinkTopic
PropertyChanged "LinkTopic"
End Property

Public Property Get Locked() As Boolean
Locked = txtFlex.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
txtFlex.Locked() = New_Locked
PropertyChanged "Locked"
End Property

Public Property Get MouseIcon() As Picture
Set MouseIcon = txtFlex.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
Set txtFlex.MouseIcon = New_MouseIcon
PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
MousePointer = txtFlex.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
txtFlex.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property

Private Sub txtFlex_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
txtFlex.OLEDrag
End Sub

Private Sub txtFlex_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Public Property Get OLEDragMode() As flxOleDragMode
OLEDragMode = txtFlex.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As flxOleDragMode)
txtFlex.OLEDragMode() = New_OLEDragMode
PropertyChanged "OLEDragMode"
End Property

Private Sub txtFlex_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Public Property Get OLEDropMode() As flxOleDropMode
OLEDropMode = txtFlex.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As flxOleDropMode)
txtFlex.OLEDropMode() = New_OLEDropMode
PropertyChanged "OLEDropMode"
End Property

Public Property Get SelLength() As Long
SelLength = m_SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
txtFlex.SelLength() = Abs(New_SelLength)
m_SelLength = txtFlex.SelLength
PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
txtFlex.SelStart() = Abs(New_SelStart)
m_SelStart = txtFlex.SelStart
PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
SelText = txtFlex.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
Dim i                As Integer
Dim iSelStart        As Integer
Dim iSelLength       As Integer
If txtFlex.Locked Then Exit Property

iSelStart = txtFlex.SelStart
iSelLength = txtFlex.SelLength
Call DeleteSelTxt
txtFlex.SelStart = iSelStart

For i = 1 To Len(New_SelText) Step 1
   Call InsertChar(Asc(Mid$(New_SelText, i, 1)))
Next
m_SelStart = 0
m_SelLength = 0
PropertyChanged "SelText"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
   m_DecimalSeperator = .ReadProperty("DecimalSeperator", m_def_DecimalSeperator)
   m_iDecimalSeperator = Asc(m_DecimalSeperator)
   txtFlex.Enabled = .ReadProperty("Enabled", True)
   Set txtFlex.Font = .ReadProperty("Font", Ambient.Font)
   Set UserControl.Font = txtFlex.Font
   m_CenturyON = .ReadProperty("CenturyON", m_def_CenturyON)
   txtFlex.BorderStyle = .ReadProperty("BorderStyle", 1)
   txtFlex.LinkItem = .ReadProperty("LinkItem", "")
   txtFlex.LinkMode = .ReadProperty("LinkMode", 0)
   txtFlex.LinkTimeout = .ReadProperty("LinkTimeout", 50)
   txtFlex.LinkTopic = .ReadProperty("LinkTopic", "")
   txtFlex.Locked = .ReadProperty("Locked", False)
   txtFlex.MaxLength = .ReadProperty("MaxLength", 0)
   m_mask = .ReadProperty("Mask", m_def_Mask)
   Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
   txtFlex.MousePointer = .ReadProperty("MousePointer", 0)
   txtFlex.OLEDragMode = .ReadProperty("OLEDragMode", 0)
   txtFlex.OLEDropMode = .ReadProperty("OLEDropMode", 0)
   m_SelLength = .ReadProperty("SelLength", 0)
   m_SelStart = .ReadProperty("SelStart", 0)
   txtFlex.SelText = .ReadProperty("SelText", "")
   txtFlex.ToolTipText = .ReadProperty("ToolTipText", "")
   mi_FieldType = .ReadProperty("FieldType", m_def_FieldType)
   ms_PromptChar = .ReadProperty("PromptChar", m_def_PromptChar)
   mi_Asc_ms_PromptChar = Asc(ms_PromptChar)
   mb_AutoToLastPos = .ReadProperty("AutoToLastPos", m_def_AutoToLastPos)
   m_PromptInclude = .ReadProperty("PromptInclude", m_def_PromptInclude)
   m_BeepOnError = .ReadProperty("BeepOnError", m_def_BeepOnError)
   m_MaskCharInclude = .ReadProperty("MaskCharInclude", m_def_MaskCharInclude)
   m_ExitOnEnter = .ReadProperty("ExitOnEnter", m_def_ExitOnEnter)
   m_Century = .ReadProperty("Century", m_def_Century)
   m_SpecialChars = .ReadProperty("SpecialChars", m_def_SpecialChars)
   m_DataHasChanged = .ReadProperty("DataHasChanged", m_def_DataHasChanged)
   m_FormatString = .ReadProperty("FormatString", m_def_FormatString)
   m_FormatAlignment = .ReadProperty("FormatAlignment", m_def_FormatAlignment)
   m_Alignment = .ReadProperty("Alignment", m_def_Alignment)
   txtFlex.Alignment = m_Alignment
   m_HideSelection = .ReadProperty("HideSelection", m_def_HideSelection)
   txtFlex.FontBold = .ReadProperty("FontBold", 0)
   txtFlex.FontItalic = .ReadProperty("FontItalic", 0)
   txtFlex.FontSize = .ReadProperty("FontSize", 8)
   txtFlex.FontStrikethru = .ReadProperty("FontStrikethru", 0)
   txtFlex.FontUnderline = .ReadProperty("FontUnderline", 0)
   m_CalculatorInMenu = .ReadProperty("CalculatorInMenu", m_def_CalculatorInMenu)
   m_UpAndDownKeys2NextFlexMask = .ReadProperty("UpAndDownKeys2NextFlexMask", m_def_UpAndDownKeys2NextFlexMask)
   m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
   m_ForeColorInActive = .ReadProperty("ForeColorInActive", m_def_ForeColorInActive)
   m_BackColorInActive = .ReadProperty("BackColorInActive", m_def_BackColorInActive)
   m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
   m_AutoNextFlexInput = .ReadProperty("AutoNextFlexInput", m_def_AutoNextFlexInput)
   Set m_FormatFont = .ReadProperty("FormatFont", Ambient.Font)
   m_InsertZerosInNumField = .ReadProperty("InsertZerosInNumField", m_def_InsertZerosInNumField)
   m_DateFormat = .ReadProperty("DateFormat", m_def_DateFormat)
   m_DateSeperator = .ReadProperty("DateSeperator", m_def_DateSeperator)
   m_InOutDateFormat = .ReadProperty("InOutDateFormat", m_def_InOutDateFormat)
   m_InOutDateSeperator = .ReadProperty("InOutDateSeperator", m_def_InOutDateSeperator)
   
   If mi_FieldType = DateField Then
      bRefresh = True
   End If
   
   Text = .ReadProperty("Text", "")
   m_Char2ActivateTranNextChar2Upper = .ReadProperty("Char2ActivateTranNextChar2Upper", m_def_Char2ActivateTranNextChar2Upper)
   m_BorderIfFocus = .ReadProperty("BorderIfFocus", m_def_BorderIfFocus)
   m_BorderNoFocus = .ReadProperty("BorderNoFocus", m_def_BorderNoFocus)
   m_AutoSelect = .ReadProperty("AutoSelect", m_def_AutoSelect)
   txtFlex.Tag = .ReadProperty("Tag", "")
End With

End Sub


Private Sub UserControl_Show()
lInitStyle = GetWindowLong(txtFlex.hwnd, GWL_STYLE)

If bRefresh Then Refresh

If txtFlex.ToolTipText <> Extender.ToolTipText Then
   ToolTipText = Extender.ToolTipText
End If
If Extender.Tag <> txtFlex.Tag Then
   txtFlex.Tag = Extender.Tag
End If

If m_ForeColorInActive <> 0 Then
   txtFlex.ForeColor = m_ForeColorInActive
   txtFlex.BackColor = m_BackColorInActive
Else
   txtFlex.ForeColor = m_ForeColor
   txtFlex.BackColor = m_BackColor
End If


If Ambient.UserMode Then
   If iDecPoint Then
      Call ToDecPoint
      txtFlex.SelStart = txtFlex.SelStart - 1
   Else
      If mb_AutoToLastPos And bMaskUsed Then GotoFirstPrompChar
   End If
   Call FormatView
End If
Border = m_BorderNoFocus
End Sub

Private Sub UserControl_Terminate()
Set m_FormatFont = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
   Call .WriteProperty("Enabled", txtFlex.Enabled, True)
   Call .WriteProperty("Font", txtFlex.Font, Ambient.Font)
   Call .WriteProperty("BorderStyle", txtFlex.BorderStyle, 1)
   Call .WriteProperty("LinkItem", txtFlex.LinkItem, "")
   Call .WriteProperty("LinkMode", txtFlex.LinkMode, 0)
   Call .WriteProperty("LinkTimeout", txtFlex.LinkTimeout, 50)
   Call .WriteProperty("LinkTopic", txtFlex.LinkTopic, "")
   Call .WriteProperty("Locked", txtFlex.Locked, False)
   Call .WriteProperty("MaxLength", txtFlex.MaxLength, 0)
   Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call .WriteProperty("MousePointer", txtFlex.MousePointer, 0)
   Call .WriteProperty("OLEDragMode", txtFlex.OLEDragMode, 0)
   Call .WriteProperty("OLEDropMode", txtFlex.OLEDropMode, 0)
   Call .WriteProperty("SelLength", m_SelLength, 0)
   Call .WriteProperty("SelStart", m_SelStart, 0)
   Call .WriteProperty("SelText", txtFlex.SelText, "") '
   Call .WriteProperty("Text", ms_Text, "")
   Call .WriteProperty("ToolTipText", txtFlex.ToolTipText, "")
   Call .WriteProperty("FieldType", mi_FieldType, m_def_FieldType)
   Call .WriteProperty("PromptChar", ms_PromptChar, m_def_PromptChar)
   Call .WriteProperty("AutoToLastPos", mb_AutoToLastPos, m_def_AutoToLastPos)
   Call .WriteProperty("PromptInclude", m_PromptInclude, m_def_PromptInclude)
   Call .WriteProperty("BeepOnError", m_BeepOnError, m_def_BeepOnError)
   Call .WriteProperty("MaskCharInclude", m_MaskCharInclude, m_def_MaskCharInclude)
   Call .WriteProperty("CenturyON", m_CenturyON, m_def_CenturyON)
   Call .WriteProperty("ExitOnEnter", m_ExitOnEnter, m_def_ExitOnEnter)
   Call .WriteProperty("Century", m_Century, m_def_Century)
   Call .WriteProperty("SpecialChars", m_SpecialChars, m_def_SpecialChars)
   Call .WriteProperty("DataHasChanged", m_DataHasChanged, m_def_DataHasChanged)
   Call .WriteProperty("FormatString", m_FormatString, m_def_FormatString)
   Call .WriteProperty("FormatAlignment", m_FormatAlignment, m_def_FormatAlignment)
   Call .WriteProperty("Alignment", m_Alignment, m_def_Alignment)
   Call .WriteProperty("HideSelection", m_HideSelection, m_def_HideSelection)
   Call .WriteProperty("FontBold", txtFlex.FontBold, 0)
   Call .WriteProperty("FontItalic", txtFlex.FontItalic, 0)
   Call .WriteProperty("FontSize", txtFlex.FontSize, 8)
   Call .WriteProperty("FontStrikethru", txtFlex.FontStrikethru, 0)
   Call .WriteProperty("FontUnderline", txtFlex.FontUnderline, 0)
   Call .WriteProperty("CalculatorInMenu", m_CalculatorInMenu, m_def_CalculatorInMenu)
   Call .WriteProperty("UpAndDownKeys2NextFlexMask", m_UpAndDownKeys2NextFlexMask, m_def_UpAndDownKeys2NextFlexMask)
   Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
   Call .WriteProperty("ForeColorInActive", m_ForeColorInActive, m_def_ForeColorInActive)
   Call .WriteProperty("BackColorInActive", m_BackColorInActive, m_def_BackColorInActive)
   Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
   Call .WriteProperty("AutoNextFlexInput", m_AutoNextFlexInput, m_def_AutoNextFlexInput)
   Call .WriteProperty("FormatFont", m_FormatFont, Ambient.Font)
   Call .WriteProperty("InsertZerosInNumField", m_InsertZerosInNumField, m_def_InsertZerosInNumField)
   Call .WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
   Call .WriteProperty("DateSeperator", m_DateSeperator, m_def_DateSeperator)
   Call .WriteProperty("Mask", m_mask, m_def_Mask)
   Call .WriteProperty("InOutDateFormat", m_InOutDateFormat, m_def_InOutDateFormat)
   Call .WriteProperty("InOutDateSeperator", m_InOutDateSeperator, m_def_InOutDateSeperator)
   Call .WriteProperty("Char2ActivateTranNextChar2Upper", m_Char2ActivateTranNextChar2Upper, m_def_Char2ActivateTranNextChar2Upper)
   Call .WriteProperty("BorderIfFocus", m_BorderIfFocus, m_def_BorderIfFocus)
   Call .WriteProperty("BorderNoFocus", m_BorderNoFocus, m_def_BorderNoFocus)
   Call .WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
   Call .WriteProperty("Tag", txtFlex.Tag, "")
   Call .WriteProperty("DecimalSeperator", m_DecimalSeperator, m_def_DecimalSeperator)
End With
End Sub

Public Property Get PromptInclude() As Boolean
Attribute PromptInclude.VB_Description = "PromptChar Included in Text property"
PromptInclude = m_PromptInclude
End Property

Public Property Let PromptInclude(ByVal New_PromptInclude As Boolean)
m_PromptInclude = New_PromptInclude
PropertyChanged "PromptInclude"
End Property

Public Property Get BeepOnError() As Boolean
Attribute BeepOnError.VB_Description = "Geluid Bij Foute Invoer "
Attribute BeepOnError.VB_ProcData.VB_Invoke_Property = "Characteristics"
BeepOnError = m_BeepOnError
End Property

Public Property Let BeepOnError(ByVal New_BeepOnError As Boolean)
m_BeepOnError = New_BeepOnError
PropertyChanged "BeepOnError"
End Property

Public Property Get MaskCharInclude() As Boolean
Attribute MaskCharInclude.VB_Description = "Mask Karakters In Uitvoer Tekst"
MaskCharInclude = m_MaskCharInclude
End Property

Public Property Let MaskCharInclude(ByVal New_MaskCharInclude As Boolean)
m_MaskCharInclude = New_MaskCharInclude
PropertyChanged "MaskCharInclude"
End Property

Public Property Get CenturyON() As Boolean
CenturyON = m_CenturyON
End Property

Public Property Get ExitOnEnter() As Boolean
Attribute ExitOnEnter.VB_ProcData.VB_Invoke_Property = "Characteristics"
ExitOnEnter = m_ExitOnEnter
End Property

Public Property Let ExitOnEnter(ByVal New_ExitOnEnter As Boolean)
m_ExitOnEnter = New_ExitOnEnter
PropertyChanged "ExitOnEnter"
End Property

Public Property Get Century() As Integer
Century = m_Century
End Property

Public Property Let Century(ByVal New_Century As Integer)
m_Century = Min(Max(0, New_Century), 35)
bRefresh = True
PropertyChanged "Century"
End Property

Public Property Get SpecialChars() As String
Attribute SpecialChars.VB_Description = "Karakters Welke Door Het MaskFilter Genegeerd Worden (bv * voor Find ..*)"
Attribute SpecialChars.VB_ProcData.VB_Invoke_Property = "Characteristics"
SpecialChars = m_SpecialChars
End Property

Public Property Let SpecialChars(ByVal New_SpecialChars As String)
m_SpecialChars = New_SpecialChars
PropertyChanged "SpecialChars"
End Property

Public Property Get DataHasChanged() As Boolean
Attribute DataHasChanged.VB_MemberFlags = "400"
DataHasChanged = (sOldTxt <> ms_Text) Or m_DataHasChanged
End Property

Public Property Let DataHasChanged(ByVal New_DataHasChanged As Boolean)
If Ambient.UserMode = False Then Err.Raise 387
m_DataHasChanged = New_DataHasChanged
PropertyChanged "DataHasChanged"
End Property

Public Property Get FormatString() As String
Attribute FormatString.VB_ProcData.VB_Invoke_Property = "Characteristics"
FormatString = m_FormatString
End Property

Public Property Let FormatString(ByVal New_Format As String)
m_FormatString = New_Format
PropertyChanged "FormatString"
End Property

Public Property Get FormatAlignment() As flxAlignMent
FormatAlignment = m_FormatAlignment
End Property

Public Property Let FormatAlignment(ByVal New_FormatAlignment As flxAlignMent)
m_FormatAlignment = New_FormatAlignment
PropertyChanged "FormatAlignment"
End Property

Public Property Get Alignment() As flxAlignMent
Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As flxAlignMent)
m_Alignment = New_Alignment
txtFlex.Alignment = m_Alignment
PropertyChanged "Alignment"
End Property

Public Property Get MaxLength() As Long
MaxLength = txtFlex.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
If mi_FieldType = DateField Then
   txtFlex.MaxLength = 0
Else
   txtFlex.MaxLength() = Min(Max(0, New_MaxLength), 999)
End If
bRefresh = True
PropertyChanged "MaxLength"
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
hwnd = UserControl.hwnd
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Specifies whether the selection in a Masked edit control is hidden when the control loses focus."
HideSelection = m_HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
m_HideSelection = New_HideSelection
PropertyChanged "HideSelection"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
FontBold = txtFlex.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
txtFlex.FontBold() = New_FontBold
PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
FontItalic = txtFlex.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
txtFlex.FontItalic() = New_FontItalic
PropertyChanged "FontItalic"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
FontSize = txtFlex.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
txtFlex.FontSize() = New_FontSize
PropertyChanged "FontSize"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
FontStrikethru = txtFlex.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
txtFlex.FontStrikethru() = New_FontStrikethru
PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
FontUnderline = txtFlex.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
txtFlex.FontUnderline() = New_FontUnderline
PropertyChanged "FontUnderline"
End Property

Public Property Get Calculator() As String
Dim s    As String
Dim s1   As String * 1
Dim s2   As String
Dim i    As Integer
bNoEval = True
s = Replace(IIf(txtFlex.SelText <> "", txtFlex.SelText, ms_Text), ms_PromptChar, "")
For i = 1 To Len(s)
   s1 = Mid$(s, i, 1)
   If IsNumeric(s1) Or s1 = "-" Then
      s2 = s2 & s1
   ElseIf s1 = m_DecimalSeperator Then
      s2 = s2 & m_DecimalSeperator
   End If
Next
Calculator = frmCalculator.Calc(s2)
End Property

Private Sub DeleteSelTxt()
Dim i                   As Integer
Dim iTmpSelStart        As Integer
If ms_Text = "" Or txtFlex.Locked Then Exit Sub
iTmpSelStart = txtFlex.SelStart
For i = txtFlex.SelStart + 1 To txtFlex.SelStart + txtFlex.SelLength
   If WhatAction(i) <> NoValidPos Then
      Mid$(ms_Text, i, 1) = ms_PromptChar
   End If
Next
txtFlex = ms_Text
txtFlex.SelStart = iTmpSelStart
End Sub

Public Property Get CalculatorInMenu() As Boolean
Attribute CalculatorInMenu.VB_ProcData.VB_Invoke_Property = "Characteristics"
CalculatorInMenu = m_CalculatorInMenu
End Property

Public Property Let CalculatorInMenu(ByVal New_CalculatorInMenu As Boolean)
m_CalculatorInMenu = New_CalculatorInMenu
PropertyChanged "CalculatorInMenu"
End Property

Public Property Get Version()
Version = AppVersion()
End Property

Public Property Get UpAndDownKeys2NextFlexMask() As Boolean
Attribute UpAndDownKeys2NextFlexMask.VB_ProcData.VB_Invoke_Property = "Characteristics"
UpAndDownKeys2NextFlexMask = m_UpAndDownKeys2NextFlexMask
End Property

Public Property Let UpAndDownKeys2NextFlexMask(ByVal New_UpAndDownKeys2NextFlexMask As Boolean)
m_UpAndDownKeys2NextFlexMask = New_UpAndDownKeys2NextFlexMask
PropertyChanged "UpAndDownKeys2NextFlexMask"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorInActive() As OLE_COLOR
ForeColorInActive = m_ForeColorInActive
End Property

Public Property Let ForeColorInActive(ByVal New_ForeColorInActive As OLE_COLOR)
m_ForeColorInActive = New_ForeColorInActive
PropertyChanged "ForeColorInActive"
End Property

Public Property Get BackColorInActive() As OLE_COLOR
BackColorInActive = m_BackColorInActive
End Property

Public Property Let BackColorInActive(ByVal New_BackColorInActive As OLE_COLOR)
m_BackColorInActive = New_BackColorInActive
PropertyChanged "BackColorInActive"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
m_BackColor = New_BackColor
PropertyChanged "BackColor"
End Property

Public Property Get AutoNextFlexInput() As Boolean
Attribute AutoNextFlexInput.VB_ProcData.VB_Invoke_Property = "Characteristics"
AutoNextFlexInput = m_AutoNextFlexInput
End Property

Public Property Let AutoNextFlexInput(ByVal New_AutoNextFlexInput As Boolean)
m_AutoNextFlexInput = New_AutoNextFlexInput
PropertyChanged "AutoNextFlexInput"
End Property

Public Property Get FormatFont() As Font
Set FormatFont = m_FormatFont
End Property

Public Property Set FormatFont(ByVal New_FormatFont As Font)
Set m_FormatFont = New_FormatFont
PropertyChanged "FormatFont"
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
Set Font = txtFlex.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set txtFlex.Font = New_Font
Set UserControl.Font = txtFlex.Font
PropertyChanged "Font"
End Property

Public Property Get ToolTipText() As String
ToolTipText = txtFlex.ToolTipText
End Property

Public Property Let ToolTipText(sToolTipText As String)
txtFlex.ToolTipText = sToolTipText
Extender.ToolTipText = sToolTipText
PropertyChanged "ToolTipText"
End Property

Public Property Get InsertZerosInNumField() As Boolean
InsertZerosInNumField = m_InsertZerosInNumField
End Property

Public Property Let InsertZerosInNumField(ByVal New_InsertZerosInNumField As Boolean)
m_InsertZerosInNumField = New_InsertZerosInNumField
PropertyChanged "InsertZerosInNumField"
End Property

Public Property Get Dateformat() As flxDateFormat
Dateformat = m_DateFormat
End Property

Public Property Let Dateformat(ByVal New_DateFormat As flxDateFormat)
m_DateFormat = New_DateFormat
If m_DateFormat = LocalSetting Then
   m_DateSeperator = LocalDateSeperator
End If
bRefresh = True
PropertyChanged "DateFormat"
End Property

Public Property Get DateSeperator() As String
DateSeperator = m_DateSeperator
End Property

Public Function DateFormatByText() As String
DateFormatByText = InsSepInDateformat(m_DateFormat)
End Function

Public Property Let DateSeperator(ByVal New_DateSeperator As String)
m_DateSeperator = IIf(New_DateSeperator = "", "-", New_DateSeperator)
If m_DateFormat = LocalSetting Then
   m_DateSeperator = LocalDateSeperator
End If
bRefresh = True
PropertyChanged "DateSeperator"
End Property

Public Function ValidDate(Optional sDate As String, Optional iFormat As flxDateFormat = -1, Optional sSeperator As String) As Boolean
'Valid formats are :
'ddmmyyyy ddyyyymm yyyyddmm yyyymmdd mmddyyyy mmyyyydd
'ddmmyy   ddyymm   yyddmm   yymmdd   mmddyy   mmyydd
'mmdd     ddmm
'ANY DELIMITER is ALLOWED for example
'dd-mm-yyyy or dd@mm@YY or dd/mm/yy

'maybe an ridiculous input like ValidDate("BLABLA", "TSJA")
On Error GoTo ExitOnError
Dim sTmp          As String
Dim iDay          As Integer
Dim iMonth        As Integer
Dim iYear         As Integer
Dim i             As Integer
Dim sFormat       As String
Dim sSep          As String

If sDate = "" Then sDate = ms_Text
If iFormat = -1 Then iFormat = m_DateFormat
sFormat = InsSepInDateformat(iFormat)
sSep = IIf(iFormat = LocalSetting, LocalDateSeperator(), IIf(sSeperator <> "", sSeperator, m_DateSeperator))
iYear = -1
For i = 0 To 2
   sTmp = Token(sDate, i, sSep)
   Select Case Left$(UCase$(Token(sFormat, i, sSep)), 1)
      Case "D"
         If IsNumeric(sTmp) Then iDay = Val(sTmp)
      Case "M"
         If IsNumeric(sTmp) Then iMonth = Val(sTmp)
      Case "Y"
         If IsNumeric(sTmp) And Val(sTmp) < 9999 Then
            iYear = Year(Format$("1", "MMMM") & " 1," & CStr(Val(sTmp)))
         Else
            Exit Function
         End If
   End Select
Next

If iDay < 1 Or iMonth < 1 Or iMonth > 12 Then Exit Function

If iDay > Choose(iMonth, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31) Then
   'only one exception (leap year) but then the year is required
   If iMonth = 2 And iDay = 29 And iYear > -1 Then
      If iYear Mod 100 <> 0 And iYear Mod 4 = 0 Or iYear Mod 400 = 0 Then
         ValidDate = True
      End If
   End If
Else
   ValidDate = True
End If
ExitOnError:
End Function


Public Function TransFormDate(sDate As String, Optional InDateFormat As flxDateFormat = LocalSetting, Optional OutDateFormat As flxDateFormat = LocalSetting, Optional sInSep As String, Optional sOutSep As String) As String
On Error GoTo ExitOnError
Dim s1(0 To 2)          As String
Dim s2(0 To 2)          As String
Dim sSep                As String
Dim sFormat             As String
Dim i                   As Integer
Dim i2                  As Integer
Dim sTmp                As String
Dim sTmp2               As String

If sDate = "" Then Exit Function
sSep = IIf(InDateFormat = LocalSetting, LocalDateSeperator(), IIf(sInSep <> "", sInSep, m_DateSeperator))
sFormat = InsSepInDateformat(InDateFormat, sSep)

For i = 0 To 2
   sTmp = Token(sDate, i, sSep)
   If IsNumeric(sTmp) Then
      s1(i) = CStr(Val(sTmp))
   End If
   s2(i) = UCase$(Token(sFormat, i, sSep))
Next

sSep = IIf(OutDateFormat = LocalSetting, LocalDateSeperator(), IIf(sOutSep <> "", sOutSep, m_DateSeperator))

sFormat = InsSepInDateformat(OutDateFormat, sSep)

For i = 0 To 2
   sTmp = UCase$(Token(sFormat, i, sSep))
   If sTmp <> "" Then
      For i2 = 0 To 2
         If InStr(s2(i2), Left$(sTmp, 1)) > 0 Then
            If s1(i2) = "" Then
               Exit Function
            ElseIf Len(sTmp) < 3 Then
               If Len(s1(i2)) > 2 Then
                  s1(i2) = Right$(s1(i2), 2)
               End If
            Else
               If Val(s1(i2)) < 9999 Then
                  s1(i2) = CStr(Year(Format$("1", "MMMM") & " 1," & Abs(s1(i2))))
               End If
            End If
            sTmp2 = sTmp2 & IIf(sTmp2 <> "", sSep, "") & Right$("0000" & s1(i2), Len(sTmp))
            Exit For
         End If
      Next
   End If
Next
TransFormDate = sTmp2
Exit Function
ExitOnError:
End Function

Private Function InsSepInDateformat(Optional eFormat As flxDateFormat = LocalSetting, Optional sSeperator As String) As String
Dim i As Integer
Dim sTmpF As String
Dim sSep As String

If eFormat = LocalSetting Then
   InsSepInDateformat = LocalDate()
   Exit Function
End If
sSep = IIf(sSeperator = "", m_DateSeperator, sSeperator)
sTmpF = sDateFormats(eFormat + 1)
InsSepInDateformat = Left$(sTmpF, 1)
For i = 2 To Len(sTmpF)
   InsSepInDateformat = InsSepInDateformat & IIf(Mid$(sTmpF, i, 1) <> Mid$(sTmpF, i - 1, 1), IIf(eFormat = LocalSetting, "", sSep), "") & Mid$(sTmpF, i, 1)
Next
End Function

Public Property Get InOutDateformat() As flxDateFormat
InOutDateformat = m_InOutDateFormat
End Property

Public Property Let InOutDateformat(ByVal New_InOutDateFormat As flxDateFormat)
m_InOutDateFormat = New_InOutDateFormat
If m_InOutDateFormat = LocalSetting Then
   m_InOutDateSeperator = LocalDateSeperator
End If
PropertyChanged "InOutDateFormat"
End Property

Public Property Get InOutDateSeperator() As String
InOutDateSeperator = m_InOutDateSeperator
End Property

Public Property Let InOutDateSeperator(ByVal New_InOutDateSeperator As String)
If m_InOutDateFormat = LocalSetting Then
   m_InOutDateSeperator = LocalDateSeperator()
Else
    m_InOutDateSeperator = IIf(New_InOutDateSeperator = "", LocalDateSeperator(), New_InOutDateSeperator)
End If
PropertyChanged "InOutDateSeperator"
End Property

Public Property Get Char2ActivateTranNextChar2Upper() As String
Char2ActivateTranNextChar2Upper = m_Char2ActivateTranNextChar2Upper
End Property

Public Property Let Char2ActivateTranNextChar2Upper(ByVal New_Char2ActivateTranNextChar2Upper As String)
m_Char2ActivateTranNextChar2Upper = New_Char2ActivateTranNextChar2Upper
PropertyChanged "Char2ActivateTranNextChar2Upper"
End Property

Public Property Get BorderIfFocus() As Boolean
BorderIfFocus = m_BorderIfFocus
End Property

Public Property Let BorderIfFocus(ByVal New_BorderIfFocus As Boolean)
m_BorderIfFocus = New_BorderIfFocus
PropertyChanged "BorderIfFocus"
End Property

Public Property Get BorderNoFocus() As Boolean
BorderNoFocus = m_BorderNoFocus
End Property

Public Property Let BorderNoFocus(ByVal New_BorderNoFocus As Boolean)
m_BorderNoFocus = New_BorderNoFocus
PropertyChanged "BorderNoFocus"
End Property

Public Property Get AutoSelect() As Boolean
Attribute AutoSelect.VB_Description = "Select All The New Text Automatic; SelLength = Len(Text)"
AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
m_AutoSelect = New_AutoSelect
PropertyChanged "AutoSelect"
End Property

Public Property Let Visible(ByVal xVisible As Boolean)
Extender.Visible = xVisible
End Property

Public Property Get Visible() As Boolean
Visible = Extender.Visible
End Property

Public Property Get Tag() As String
Tag = txtFlex.Tag
End Property

Public Property Let Tag(ByVal New_Tag As String)
txtFlex.Tag = New_Tag
Extender.Tag = New_Tag
PropertyChanged "Tag"
End Property

Public Property Get TextAsDisplayed() As String
TextAsDisplayed = ms_Text
End Property

Public Property Get DecimalSeperator() As String
Attribute DecimalSeperator.VB_ProcData.VB_Invoke_Property = "Characteristics"
DecimalSeperator = m_DecimalSeperator
End Property

Public Property Let DecimalSeperator(ByVal New_DecimalSeperator As String)
If Len(New_DecimalSeperator) <> 1 Then Err.Raise 122, , "MUST be 1 char long"
m_DecimalSeperator = Mid$(New_DecimalSeperator, 1, 1)
m_iDecimalSeperator = Asc(m_DecimalSeperator)
PropertyChanged "DecimalSeperator"
End Property

