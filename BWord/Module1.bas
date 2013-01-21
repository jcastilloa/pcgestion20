Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302
Public Const WM_COPY = &H301
Global DocumentClosed As Boolean
Global frmCount As Integer
'Global ChildForms(1 To 30) As frmDocument
Global UnAvail(1 To 30) As Boolean
Global Pos As Integer
Global SearchStr As String
Global MatchCase As Boolean
Global Opened As Boolean
Global DocTemp As Integer, NeedSaved(30) As Boolean
Global file(1 To 30) As String
Global strSaveClipBoard As String


Public miCampo As ADODB.field

Function frm() As Integer
    On Error GoTo CreateNew
    ' frm will return the Active form number
    ' very useful to find out the active form
    ' when we are using an array of frmDocument
    'frm = Val(frmMain.ActiveForm.Tag)
    Exit Function
CreateNew:
    ' New document created
    Dim ret As Integer
    DocTemp = FirstAvail
    If DocTemp <> -1 Then
     '   Set ChildForms(DocTemp) = New frmDocument
   '     ChildForms(DocTemp).Caption = "Document " & DocTemp
       ' ChildForms(DocTemp).Tag = DocTemp
    Else
        MsgBox "You are only allowed 30 Documents opened at one time."
    End If
    
    'frm = Val(frmMain.ActiveForm.Tag)
End Function
Function FirstAvail() As Integer
    For x = 1 To 30
        If UnAvail(x) = False Then
            UnAvail(x) = True
            FirstAvail = x
            Exit Function
        End If
    Next
    ' -1 means maximum of 30 documents have been opened
    FirstAvail = -1
End Function
Function GetBinary(Number As Integer) As String
    Dim binstr As String
    binstr = ""
    Number = Number + 1
    For x = 7 To 0 Step -1
        If Number > 2 ^ x Then
            Number = Number - 2 ^ x
            binstr = binstr & "1"
        Else
            binstr = binstr & "0"
        End If
    Next
    GetBinary = binstr
End Function
Function BintoDec(binstr As String) As Integer
    Dim Number As Integer
    For x = 0 To 7
        If Mid$(binstr, x + 1, 1) = "1" Then
            Number = Number + (2 ^ (7 - x))
        End If
    Next
    BintoDec = Number
End Function
Public Sub DisplayLineNumber()

Dim LineNum As Integer

With frmMain.rtfText
    If .SelStart > 0 Then
        LineNum = .GetLineFromChar(.SelStart)
    End If
    frmMain.sbStatusBar.Panels("Text").Text = "Numero de línea: " & CStr(LineNum + 1)
End With
End Sub

Public Sub BPrint()
On Error Resume Next
    ' Don't Print if Document is Blank
    If ActiveForm Is Nothing Then Exit Sub
    
    With dlgCommonDialog
        .DialogTitle = "Print"
        ' Cancel Error will Raise an error
        ' on Cancel
        .CancelError = True
        
        ' cdlPDReturnDC returns device context
        ' to hDC ( we will use in printing)
        ' cdlPDNoPageNums will hide the print
        ' selection ie Print pages from to
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        
        ' On selection by user
        ' only the selected text will be printed
        If ActiveForm.rtfText.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        
        .ShowPrinter
        
        ' If printing is not canceled
        ' Document will be printed
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hdc
        End If
    
    End With
End Sub
