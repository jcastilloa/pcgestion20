VERSION 5.00
Begin VB.Form frmFIB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FlexInputBox"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin PCGestion.FlexMaskEditBox FlexMaskEditBox 
      Height          =   780
      Index           =   0
      Left            =   1935
      TabIndex        =   3
      Top             =   225
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   1376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8,25
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
   Begin VB.VScrollBar VScroll1 
      CausesValidation=   0   'False
      Height          =   1635
      LargeChange     =   200
      Left            =   540
      Max             =   0
      SmallChange     =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "cmdButtons"
      Height          =   375
      Index           =   0
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblQuestion 
      Caption         =   "lblQuestion"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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


Private lOldPos               As Long
Private bCancel               As Boolean
Private iTopFirstQuestion     As Integer
Private iMinimumWidthButton   As Integer

Public ButtonClicked          As Integer
Public UC                     As FlexMaskInputBox


Private Function TranMask2CalcLen(sMask As String) As String
Dim i       As Integer
Dim s2      As String
Dim s       As String * 1

For i = 1 To Len(sMask) Step 1
   s = Mid$(sMask, i, 1)
   If InStr("9#", s) Then
      s2 = s2 & "9"
   ElseIf InStr("aA&Cc?", s) Then
      s2 = s2 & "W"
   ElseIf InStr("{}<>\", s) = 0 Then
      s2 = s2 & s
   End If
Next
TranMask2CalcLen = s2 & "1"
End Function

Public Function CalculateBox()
Dim i As Integer, iX1 As Integer, iX2 As Integer
Dim iS2 As Long, iS3 As Long
Dim l2 As Long, i3 As Long
Dim s As String, h As Single
Dim y As Single, z As Single
Dim lblW As Long

bCancel = False
lOldPos = 0

VScroll1.Max = 0

With UC.InputForm.Setup
   iS2 = .FormWidth * 0.5
   iTopFirstQuestion = .TopFirstQuestion
   iMinimumWidthButton = .MinButtonWidth
End With

With Me
   For i = 1 To FlexMaskEditBox.Count - 1 Step 1
      With FlexMaskEditBox(i)
         Set Me.Font = .Font
         .Height = Me.TextHeight("X") + IIf(.BorderIfFocus Or .BorderNoFocus, 200, 110)
         If iS2 = 0 Then
            .Width = Me.TextWidth(TranMask2CalcLen(.Mask)) + 10
         Else
            .Width = Min(Me.TextWidth(TranMask2CalcLen(.Mask)), (iS2)) + 10
         End If
         iX1 = Max(iX1, .Width)
      End With
   Next
   
   For i = 1 To lblQuestion.Count - 1 Step 1
      Set .Font = lblQuestion(i).Font
      lblW = Max(lblW, .TextWidth(lblQuestion(i).Caption))
   Next
   
   If iS2 = 0 Then
      iS2 = Min(Screen.Width * 0.75, lblW + iX1 + 400)
   End If
   
   iS3 = iS2 + Max(iS2 - iX1, 0)
   
   For i = 1 To lblQuestion.Count - 1 Step 1
      Set .Font = lblQuestion(i).Font
      i3 = .TextWidth(lblQuestion(i).Caption)
      lblQuestion(i).Height = (.TextHeight("X") + 10) * (i3 \ iS3 + 1) + 10
      lblQuestion(i).Width = Min(iS3, lblW)
      iX2 = Max(iX2, lblQuestion(i).Width)
   Next
   .Width = Max(iX2 + iX1 + 400, UC.InputForm.Setup.FormWidth)
End With

For i = 1 To lblQuestion.Count - 1 Step 1
   With FlexMaskEditBox(i)
      i3 = Max(lblQuestion(i).Height, .Height)
      l2 = lblQuestion(Max(i - 1, 0)).Top + lblQuestion(Max(i - 1, 0)).Height + 70
      lblQuestion(i).Move 30, i3 + (IIf(i > 1, l2, iTopFirstQuestion) - lblQuestion(i).Height)
      .Move iX2 + 90, IIf(i > 1, l2, iTopFirstQuestion) + (i3 - .Height)
      lblQuestion(i).TabIndex = i * 2
      .TabIndex = i * 2 + 1
   End With
Next

With Me
   For i = 1 To cmdButtons.Count - 1 Step 1
      Set .Font = cmdButtons(i).Font
      cmdButtons(i).Width = Max(iMinimumWidthButton, TextWidth(cmdButtons(i).Caption) + 200)
      h = Max(h, TextHeight("X") + 150 + IIf(cmdButtons(i).Picture = 0, 0, 400))
   Next
   
   y = .Width
   
   With lblQuestion(lblQuestion.Count - 1)
      If UC.InputForm.Setup.MaxFormHight = 0 And (.Top + .Height + h + 700) < Screen.Height - 800 Then
         Me.Height = (.Top + .Height + h + 700)
      ElseIf UC.InputForm.Setup.MaxFormHight > (.Top + .Height + h + 400) Then
         Me.Height = UC.InputForm.Setup.MaxFormHight
      Else
         If UC.InputForm.Setup.MaxFormHight Then
            Me.Height = Min(Screen.Height - 800, UC.InputForm.Setup.MaxFormHight)
            VScroll1.Visible = True
            VScroll1.Max = ((.Top + .Height + h + 400) - Me.ScaleHeight) * 0.1
         Else
            Me.Height = Min(Screen.Height - 800, .Top + .Height + h + 400)
            VScroll1.Visible = True
            VScroll1.Max = ((.Top + .Height + h + 400) - Me.ScaleHeight) * 0.1
         End If
      End If
   End With

   z = .Height - (h + 400)
   
   For i = cmdButtons.Count - 1 To 1 Step -1
      With cmdButtons(i)
         .Move y - (.Width + 150), z, .Width, h
         y = cmdButtons(i).Left
      End With
   Next
   
   If VScroll1.Max > 0 Then
      .Width = .Width + 375
      VScroll1.Move .Width - 375, 0, 275, .ScaleHeight
      Call VScroll1_Change
   Else
      For i = 1 To lblQuestion.Count - 1 Step 1
         lblQuestion(i).Visible = True
         FlexMaskEditBox(i).Visible = True
      Next
   End If

End With

Call PaintPictureOnForm

End Function

Private Sub cmdButtons_Click(Index As Integer)
Dim Cancel As Boolean
ButtonClicked = Index
Call UC.Buttons(Index, Cancel)
If Not Cancel Then
   bCancel = True
   Me.Hide
Else
   ButtonClicked = 0
End If
End Sub

Private Sub FlexMaskEditBox_ExitOnArrowKeys(Index As Integer, Cancel As Boolean)
Call UC.ValidateMaskBox(Index, Cancel)
End Sub

Private Sub FlexMaskEditBox_PopUpItems(Index As Integer, MenuItemsArray() As Variant, TextMenu As Boolean)
Call UC.PopUpItem(Index, MenuItemsArray(), TextMenu)
End Sub

Private Sub FlexMaskEditBox_PopUpItemsClick(Index As Integer, MenuIndex As Integer)
Call UC.ItemsClick(Index, MenuIndex)
End Sub

Private Sub FlexMaskEditBox_Validate(Index As Integer, Cancel As Boolean)
Call UC.ValidateMaskBox(Index, Cancel)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyPageDown Then
   VScroll1.Value = Min(VScroll1.Max, VScroll1.Value + (Me.ScaleHeight * 0.1))
ElseIf KeyCode = vbKeyPageUp Then
   VScroll1.Value = Max(0, VScroll1.Value - (Me.ScaleHeight * 0.1))
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

If Not bCancel Then
   Cancel = vbCancel
   Exit Sub
End If
For i = 1 To Me.FlexMaskEditBox.Count - 1 Step 1
   Unload Me.FlexMaskEditBox(i)
   Unload Me.lblQuestion(i)
Next
For i = 1 To Me.cmdButtons.Count - 1 Step 1
   Unload Me.cmdButtons(i)
Next
Set UC = Nothing
End Sub

Private Sub Form_Activate()
DispOrigAuthor Me
SetFocusAPI FlexMaskEditBox(1).hwnd
End Sub

Private Sub PaintPictureOnForm()
With UC.InputForm.Picture
   If Not .Picture Is Nothing Then
      If .Height = 0 Then .Height = .Picture.Height
      If .Width = 0 Then .Width = .Picture.Width
      .posX = Min(.posX, Me.Width - (.Width + 100 + IIf(VScroll1.Max, 300, 0)))
      PaintPicture .Picture, .posX, .posY, .Width, .Height
   End If
End With
End Sub

Private Sub VScroll1_Change()
Dim i             As Integer
Dim l             As Long
Dim lCmdTop       As Long
Dim bVisible      As Boolean
Dim bOnce         As Boolean

l = VScroll1.Value

lOldPos = lOldPos - l * 10

lCmdTop = cmdButtons.Item(1).Top - 50

For i = 1 To lblQuestion.Count - 1 Step 1
   With lblQuestion.Item(i)
      .Visible = False
      .Top = .Top + lOldPos
      If .Top >= iTopFirstQuestion Then
         .Visible = (.Top + .Height) < lCmdTop
      End If
      bVisible = .Visible
   End With
   With FlexMaskEditBox(i)
      .Visible = False
      .Top = .Top + lOldPos
      .Visible = bVisible
      If Not bOnce Then
         If Me.Visible And bVisible Then
            .SetFocus
            bOnce = True
         End If
      End If
   End With
Next
lOldPos = l * 10

End Sub



