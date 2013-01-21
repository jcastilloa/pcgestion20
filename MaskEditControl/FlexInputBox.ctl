VERSION 5.00
Begin VB.UserControl FlexMaskInputBox 
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FlexInputBox.ctx":0000
   ScaleHeight     =   1125
   ScaleWidth      =   4095
   ToolboxBitmap   =   "FlexInputBox.ctx":058A
End
Attribute VB_Name = "FlexMaskInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Event ButtonClick(Index As Integer, Cancel As Boolean)
Event Validation(Index As Integer, Cancel As Boolean)
Event PopUpItems(Index As Integer, MenuItemsArray() As Variant, TextMenu As Boolean)
Event PopUpItemsClick(Index As Integer, MenuIndex As Integer)

Public InputForm                 As InputFormProps
Public Questions                 As QuestionCollection
Public cmdButtons                As ButtonCollection
Private UserForm                 As frmFIB
Attribute UserForm.VB_VarHelpID = -1

Private Sub LoadControls()
Dim i As Integer
For i = UserForm.lblQuestion.Count To Questions.Count Step 1
   Load UserForm.lblQuestion(i)
   Set UserForm.lblQuestion(i).Font = Questions(i).Font
   Load UserForm.FlexMaskEditBox(i)
Next
For i = UserForm.cmdButtons.Count To cmdButtons.Count Step 1
   Load UserForm.cmdButtons(i)
   Set UserForm.cmdButtons(i).Font = cmdButtons(i).Font
Next
End Sub

Public Property Get ButtonClicked() As Integer
ButtonClicked = UserForm.ButtonClicked
End Property

Friend Sub ItemsClick(Index As Integer, MenuIndex As Integer)
RaiseEvent PopUpItemsClick(Index, MenuIndex)
End Sub

Friend Sub Buttons(Index As Integer, Cancel As Boolean)
RaiseEvent ButtonClick(Index, Cancel)
End Sub

Friend Sub PopUpItem(Index As Integer, MenuItemsArray() As Variant, TextMenu As Boolean)
RaiseEvent PopUpItems(Index, MenuItemsArray(), TextMenu)
End Sub

Friend Sub ValidateMaskBox(Index As Integer, Cancel As Boolean)
RaiseEvent Validation(Index, Cancel)
End Sub

Public Sub TerminateInputBox()
Unload UserForm
Call UserControl_Terminate
End Sub

Private Sub UserControl_Terminate()
Set InputForm = Nothing
Set Questions = Nothing
Set cmdButtons = Nothing
Set UserForm = Nothing
End Sub

Public Sub Initialize()

Call UserControl_Terminate

Set InputForm = New InputFormProps
Set Questions = New QuestionCollection
Set cmdButtons = New ButtonCollection
Set UserForm = New frmFIB
Set UserForm.UC = Extender

End Sub

Public Property Get FlexMaskEditBox(Number) As FlexMaskEditBox
If Number < 1 Or Number > Questions.Count Then
   Err.Raise 200, , "Number " & CStr(Number) & " FlexMaskEditBox Do Not Exist! Number MUST be between 1 and " & CStr(Questions.Count)
End If
Call LoadControls
Set FlexMaskEditBox = UserForm.FlexMaskEditBox(Number)
End Property

Public Property Get FlexMaskEditBoxByObj(Number) As Object
If Number < 1 Or Number > Questions.Count Then
   Err.Raise 200, , "Number " & CStr(Number) & " FlexMaskEditBox Do Not Exist! Number MUST be between 1 and " & CStr(Questions.Count)
End If
Call LoadControls
Set FlexMaskEditBoxByObj = UserForm.FlexMaskEditBox(Number)
End Property

Public Sub ActivateInputBox()
If cmdButtons.Count = 0 Then
   cmdButtons.Add "Oke"
   cmdButtons.Add "Cancel"
End If

Call Refresh

UserForm.Show vbModal
End Sub

Public Sub Refresh()
Dim i As Integer

Call LoadControls

For i = 1 To Questions.Count
   With Questions.Item(i)
      If .BackColor = 0 Then
         .BackColor = InputForm.BackColor
      End If
   End With
Next

With UserForm
   Call ReadlblProps
   Call ReadcmdProps
   Call ReadFormProps
   Call .CalculateBox
End With
End Sub

Private Sub ReadFormProps()
With UserForm
   If InputForm.BackColor Then
      .BackColor = InputForm.BackColor
   Else
      .BackColor = UserControl.BackColor
   End If
   If InputForm.Caption <> "" Then
      .Caption = InputForm.Caption
   End If
   If Not InputForm.Icon Is Nothing Then
      Set .Icon = InputForm.Icon
   End If
   .Visible = InputForm.Visible
End With
End Sub

Private Sub ReadcmdProps()
Dim i As Integer
With UserForm
   For i = 1 To cmdButtons.Count
      With .cmdButtons(i)
        If cmdButtons(i).BackColor Then
            .BackColor = cmdButtons(i).BackColor
         End If
         Set .Font = cmdButtons(i).Font
         Set .Picture = cmdButtons(i).Picture
         .Caption = cmdButtons(i).Caption
         .Enabled = cmdButtons(i).Enabled
         .Tag = cmdButtons(i).Tag
         .ToolTipText = cmdButtons(i).ToolTipText
         .Visible = cmdButtons(i).Visible
      End With
   Next
End With
End Sub

Private Sub ReadlblProps()
Dim i As Integer
With UserForm
   For i = 1 To Questions.Count
      With .lblQuestion(i)
         If Questions.Item(i).BackColor Then
            .BackColor = Questions.Item(i).BackColor
         Else
            .BackColor = UserForm.BackColor
         End If
         
         If Questions.Item(i).ForeColor Then
            .ForeColor = Questions.Item(i).ForeColor
         Else
            .ForeColor = UserForm.ForeColor
         End If
         
         Set .Font = Questions.Item(i).Font
         
         .Caption = Questions.Item(i).Caption
         .Tag = Questions.Item(i).Tag
         .Visible = Questions.Item(i).Visible
         .Alignment = Questions.Item(i).Alignment
         
      End With
   Next
End With
End Sub

