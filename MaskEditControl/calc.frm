VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4140
   ClientLeft      =   4695
   ClientTop       =   1530
   ClientWidth     =   2775
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   2775
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdButton 
      Caption         =   "Round UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   96
      TabIndex        =   24
      Top             =   3792
      Width           =   1596
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Round  DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   96
      TabIndex        =   23
      Top             =   3456
      Width           =   1596
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1824
      TabIndex        =   22
      Top             =   3456
      Width           =   876
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Oke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1824
      TabIndex        =   21
      Top             =   3792
      Width           =   876
   End
   Begin VB.CommandButton Percent 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      Height          =   372
      Left            =   1800
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   372
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "="
      Height          =   360
      Index           =   4
      Left            =   2304
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   372
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "/"
      Height          =   372
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "*"
      Height          =   372
      Index           =   2
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      Height          =   372
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      Height          =   372
      Index           =   1
      Left            =   1824
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C&E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2280
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&C"
      Height          =   360
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1572
      UseMaskColor    =   -1  'True
      Width           =   372
   End
   Begin VB.ListBox lstListBox 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   168
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   2484
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      Height          =   372
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   372
      Index           =   8
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   372
      Index           =   9
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   372
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   372
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   372
      Index           =   6
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   372
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   372
      Index           =   2
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   372
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   372
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   852
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   372
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   372
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   384
      Left            =   120
      TabIndex        =   12
      Top             =   1104
      Width           =   2520
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Adjusted by Ted Schopenhouer   ted.schopenhouer@12Move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                           and  Douwe Konings dkonings@xs4all.nl

'This sources may be used freely without the intention of commercial distribution.
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

'In other words, when you are developing a program for yourself or for
'a company without selling this product to thirt party's it's allowed to
'use this source code. When you, or the company you work for, sells the
'program then permission is needed!!!!!


Const OKE = 0
Const CANCELACTION = 1
Const ROUNDDOWN = 2
Const ROUNDUP = 3

Enum enumInput
      NONE = 0
      OPS = 1
      CE = 2
      NUMS = 3
      NEG = 4
End Enum

Dim ReturnValue      As Variant
Dim Value1           As Double
Dim Value2           As Double
Dim DecimalFlag      As Boolean
Dim NumOps           As Integer
Dim LastInput        As enumInput
Dim OpFlag           As String
Dim TmpOpFlag        As String
Dim TempReadout      As Double
Dim TmpPercent       As Double
Dim TmpReadOut2      As Variant

Private Sub cmdButton_Click(Index As Integer)
Operator_Click 4
Select Case Index
   Case OKE
      ReturnValue = Readout
   Case CANCELACTION
      ReturnValue = "CANCEL"
   Case ROUNDDOWN
      ReturnValue = Int(Readout)
   Case ROUNDUP
      ReturnValue = IIf(CDbl(Readout) - Int(Readout) > 0, Int(Readout) + 1, Int(Readout))
End Select
Unload Me
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, Separacion_MDIForm
DecimalFlag = False
NumOps = 0
LastInput = IIf(ReturnValue > 0, NUMS, NONE)
OpFlag = " "
If InStr(CStr(ReturnValue), Format$(0, ".")) Then
   Readout = ReturnValue
   DecimalFlag = True
Else
   Readout = Format$(ReturnValue, "0.")
End If
End Sub

Private Sub CANCELACTION_Click()
Operator(4).SetFocus
Readout = Format$(0, "0.")
Value1 = 0
Value2 = 0
Form_Load
lstListBox.Clear
lstListBox.Refresh
End Sub

Private Sub CancelEntry_Click()
Operator(4).SetFocus
Readout = Format$(0, "0.")
DecimalFlag = False
LastInput = CE
End Sub

Private Sub Decimal_Click()
Operator(4).SetFocus
If LastInput = NEG Then
   Readout = Format$(0, "-0")
ElseIf LastInput <> NUMS Then
   Readout = Format$(0, "0.")
End If
DecimalFlag = True
LastInput = NUMS
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 48 To 57
      Number_Click (KeyAscii - 48)
   Case 67, 99
      CANCELACTION_Click
   Case 69, 101
      CancelEntry_Click
   Case 46
      Decimal_Click
   Case 37, 92
      Percent_Click
   Case 47
      Operator_Click 0
   Case 43
      Operator_Click 1
   Case 42
      Operator_Click 2
   Case 45
      Operator_Click 3
   Case 61
      Operator_Click 4
   Case vbKeyEscape
      Unload Me
End Select
End Sub

Private Sub Number_Click(Index As Integer)
Operator(4).SetFocus
If LastInput <> NUMS Then
    Readout = Format$(0, ".")
    DecimalFlag = False
End If
If DecimalFlag Then
    Readout = Readout + Number(Index).Caption
Else
   Readout = Left$(Readout, InStr(Readout, Format$(0, ".")) - 1) + Number(Index).Caption + Format$(0, ".")
End If
If LastInput = NEG Then Readout = "-" & Readout

LastInput = NUMS
End Sub

Private Sub Operator_Click(Index As Integer)
Dim Tmp As Variant

Operator(4).SetFocus

If lstListBox.ListCount = 0 Then
   lstListBox.AddItem "     " & IIf(Mid$(Readout, Len(Readout)) = "," Or Mid$(Readout, Len(Readout)) = ".", Left$(Readout, Len(Readout) - 1), Readout)
End If

TempReadout = Readout

If LastInput = NUMS Then
    NumOps = NumOps + 1
End If

Select Case NumOps
    Case 0
       If Operator(Index).Caption = "-" And LastInput <> NEG Then
           Readout = "-" & Readout
           LastInput = NEG
       End If
    Case 1
      Value1 = Readout
      If Operator(Index).Caption = "-" And LastInput <> NUMS And OpFlag <> "=" Then
         Readout = "-"
         LastInput = NEG
      End If
    Case 2
      Value2 = TempReadout
      Select Case OpFlag
          Case "+"
              Value1 = CDbl(Value1) + CDbl(Value2)
          Case "-"
              Value1 = CDbl(Value1) - CDbl(Value2)
          Case "*"
              Value1 = CDbl(Value1) * CDbl(Value2)
          Case "/"
              If Value2 <> 0 Then Value1 = CDbl(Value1) / CDbl(Value2)
          Case "="
              Value1 = CDbl(Value2)
          Case "%"
               TmpPercent = CDbl(Value2)
               If TmpOpFlag = "+" Then
                  Value1 = CDbl(Value1) + CDbl(Value2)
               ElseIf TmpOpFlag = "-" Then
                  Value1 = CDbl(Value1) - CDbl(Value2)
               ElseIf TmpOpFlag = "/" Then
                  Value1 = CDbl(Value1) / CDbl(Value2)
               Else
                  Value1 = CDbl(Value1) * CDbl(Value2)
               End If
      End Select
      Readout = Value1
      NumOps = 1
      Tmp = IIf(OpFlag = "%", TmpReadOut2, Value2)
      Tmp = IIf(Mid$(Tmp, Len(Tmp)) = "," Or Mid$(Tmp, Len(Tmp)) = ".", Left$(Tmp, Len(Tmp) - 1), Tmp)
      lstListBox.AddItem IIf(OpFlag = "%", "", OpFlag) & " " & TmpOpFlag & "  " & Tmp & IIf(OpFlag = "%", OpFlag & " = " & TmpPercent, "")
      lstListBox.AddItem "      " & Readout.Caption
      lstListBox.ListIndex = lstListBox.ListCount - 1
      lstListBox.ListIndex = -1
      TmpOpFlag = " "
      TmpPercent = 0
End Select
If LastInput <> NEG Then
    LastInput = OPS
    OpFlag = Operator(Index).Caption
End If
End Sub

Private Sub Percent_Click()
Operator(4).SetFocus
TmpOpFlag = OpFlag
TmpReadOut2 = Readout
Readout = (CDbl(Readout) * 0.01) * CDbl(Value1)
LastInput = OPS
OpFlag = "%"
NumOps = NumOps + 1
DecimalFlag = True
End Sub

Public Function Calc(ByVal sValue As String) As String
ReturnValue = Val(sValue)
frmCalculator.Show vbModal
Calc = ReturnValue
End Function
