VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BWord -  Find "
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFindCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblMatchCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   885
      Width           =   975
   End
   Begin VB.Label lblFind 
      Caption         =   "Find What ?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMatchCase_Click()
If chkMatchCase.Value = 1 Then
    MatchCase = True
  Else
    MatchCase = False
End If
End Sub
Private Sub cmdFindCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
Dim Resultado As Long
Dim op As String
    op = ""
    ' op=0 means find will not be case sensative
    ' op=4 means find will be case sensative
    If chkMatchCase.Value = 1 Then op = 4
    With ChildForms(frm).rtfText
        If .SelLength > 0 Then
            ' The Find function will find the specified
            ' text and return the position in the rtfText Box
            Resultado = .Find(txtFind.Text, .SelStart, .SelStart + .SelLength, Val(op))
            ' If Text Found then highlight it and unload frmFind
            If Resultado >= 0 Then
                ChildForms(frm).SetFocus
                Unload Me
                Exit Sub
            Else
            ' If text not found then repeat find from beginning
            ' depending on user's choice
                If MsgBox("BWord searched in the actual selection." & Chr(10) & "No Mtches." & Chr(10) & "Do you wish to continue?", vbYesNo + vbQuestion, "BWord") = vbNo Then
                    Exit Sub
                Else
                    GoTo BFind
                End If
            End If
        End If
        
        If .SelStart < 1 Then GoTo BFind
        
        Resultado = .Find(txtFind.Text, .SelStart, , Val(op))
        If Resultado >= 0 Then
            ChildForms(frm).SetFocus
            Unload Me
            Exit Sub
        ElseIf .SelStart > 1 Then
            If MsgBox("Do you want BWord to search from the beginning of " & ChildForms(frm).Caption & "?", vbYesNo + vbQuestion, "Search endend") = vbNo Then
                Exit Sub
            Else
                GoTo BFind
            End If
        End If
BFind:
        Resultado = .Find(txtFind.Text, 0, Len(.Text), Val(op))
        If Resultado >= 0 Then
            ChildForms(frm).SetFocus
            Unload Me
            Exit Sub
        Else
            MsgBox "BWord could not Find.", vbOKOnly + vbExclamation, "Sorry !!"
            Exit Sub
        End If
    End With

End Sub
Private Sub Form_Load()
    Pos = 0
    If MatchCase = True Then
        chkMatchCase.Value = 1
    End If
End Sub
Private Sub txtFind_Change()
    SearchStr = txtFind.Text
End Sub
