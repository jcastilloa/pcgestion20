VERSION 5.00
Begin VB.Form frmSendImage 
   Caption         =   "BWord - Insertar Imagen"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7530
   FontTransparent =   0   'False
   Icon            =   "frmSendImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.ComboBox PatternCombo 
      Height          =   315
      ItemData        =   "frmSendImage.frx":030A
      Left            =   0
      List            =   "frmSendImage.frx":030C
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.FileListBox FileList 
      Height          =   2625
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.DirListBox DirList 
      Height          =   990
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4455
      Left            =   2280
      ScaleHeight     =   4395
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmSendImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Private Sub cmdSend_Click()
On Error GoTo SendError
' Copy the picture into the clipboard.
    If Clipboard.GetFormat(vbCFText) = True Then
        strSaveClipBoard = Clipboard.GetText
        flag = True
    End If
    
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    ' Paste the picture into the RichTextBox.
    SendMessage frmMain.rtfText.hwnd, WM_PASTE, 0, 0
    If flag = True Then
        Clipboard.Clear
        Clipboard.SetText strSaveClipBoard
    Else
        Clipboard.Clear
    End If
    
    frmMain.setMenu
    frmSendImage.Hide
Exit Sub
SendError:
End Sub

Private Sub FileList_DblClick()
    cmdSend_Click
End Sub

Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
    On Error GoTo DriveError
    DirList.Path = DriveList.Drive
    Exit Sub

DriveError:
    DriveList.Drive = DirList.Path
    Exit Sub
End Sub


Private Sub FileList_Click()
Dim FName As String
    On Error GoTo LoadPictureError
    FName = FileList.Path + "\" + FileList.FileName
    Caption = "BWord Insert Picture" & " [" & FName & "]"
    MousePointer = vbHourglass
    DoEvents
    Picture1.Picture = LoadPicture(FName)
    MousePointer = vbDefault
    Exit Sub
LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "BWord Insert Picture" & " [Invalid picture]"
    Exit Sub
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    PatternCombo.AddItem "Bitmaps (*.bmp)"
    PatternCombo.AddItem "GIF (*.gif)"
    PatternCombo.AddItem "JPEG (*.jpg)"
    PatternCombo.AddItem "Matafiles (*.wmf)"
    PatternCombo.AddItem "TIF (*.tif)"
    PatternCombo.AddItem "Todos (*.gif;*.jpg;*.bmp;*.wmf;*.tif)"
    PatternCombo.ListIndex = 0
End Sub

Private Sub Form_Resize()
Const GAP = 20
On Error Resume Next
Dim wid As Integer
Dim hgt As Integer

    If WindowState = vbMinimized Then Exit Sub
    wid = DriveList.Width
    DriveList.Move GAP, GAP, wid
    PatternCombo.Move GAP, ScaleHeight - PatternCombo.Height - 750, wid
   
    hgt = (PatternCombo.Top - DriveList.Top - DriveList.Height) / 2
    If hgt < 100 Then hgt = 100
    DirList.Move GAP, DriveList.Top + DriveList.Height + GAP, wid, hgt
    FileList.Move GAP, DirList.Top + DirList.Height + GAP, wid, hgt
    cmdSend.Move ScaleWidth / 2 - cmdSend.Width / 2, PatternCombo.Top + PatternCombo.Height + 4 * GAP, wid
    Picture1.Move Picture1.Left, Picture1.Top, ScaleWidth - cmdSend.Width - 7 * GAP, 2 * hgt + 2 * PatternCombo.Height
End Sub

Private Sub PatternCombo_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer
    pat = PatternCombo.List(PatternCombo.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    FileList.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub

Private Sub Picture1_DblClick()
    cmdSend_Click
End Sub

Private Sub Picture1_Resize()
    Form_Resize
End Sub
