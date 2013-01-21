Attribute VB_Name = "Utils"
Option Explicit

' Variables for Format Painter
Global fPaint As Integer
Global fBold As Boolean
Global fItalics As Boolean
Global fUnderline As Boolean
Global fStrikeThru As Boolean
Global fSize As Integer
Global fFont As String
Global fColor As String
'Variables for Most recently used files
Global BMruFile(0 To 3) As String

' Win32 Declarations for Print sub
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long       ' Actual DC to draw on
    hdcTarget As Long ' Target DC for determining text formatting
    rc As Rect        ' Region of the DC to draw to (in twips)
    rcPage As Rect    ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type
Public Const WM_USER = &H400
Const EM_FORMATRANGE As Long = WM_USER + 57
Const PHYSICALOFFSETX As Long = 112
Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Sub SelectIt(TxtBx As TextBox)
    'Select all of the text in a TextBox control
    TxtBx.SelStart = 0
    TxtBx.SelLength = Len(TxtBx.Text)
End Sub
Public Sub CenterForm(f As Form)
    f.Left = (Screen.Width - f.Width) / 2
    f.Top = (Screen.Height - f.Height) / 2
End Sub
Public Sub Menu_FormatBullet(MainForm As Form)
    With frmMain.rtfText
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            ' selection is mixed or not bulleted
            ' so set it.
            .SelBullet = True
            frmMain.tbToolBar.Buttons("Bullets").Value = tbrPressed
        ElseIf .SelBullet = True Then
            ' selection is bold, toggle it
            .SelBullet = False
            .SelHangingIndent = False
            frmMain.tbToolBar.Buttons("Bullets").Value = tbrUnpressed
        End If
    End With
End Sub

Public Sub FPainter()
    Dim I As Integer
    On Error Resume Next
    
    ' fPaint is 1 When user clicke the
    ' Format Painter tool bar
    If fPaint = 1 Then
        ' Mouse Pointer changes to up arrow
        ' for all the current open documents
       ' For I = 1 To frmCount
            frmMain.rtfText.MousePointer = 99
      '  Next
        
        ' Saves the Font attributes of the
        ' selected text to these variables
        fBold = frmMain.rtfText.SelBold
        fItalics = frmMain.rtfText.SelItalic
        fUnderline = frmMain.rtfText.SelUnderline
        fStrikeThru = frmMain.rtfText.SelStrikeThru
        fFont = frmMain.rtfText.SelFontName
        fSize = frmMain.rtfText.SelFontSize
        fColor = frmMain.rtfText.SelColor
        
        ' Makes FPaint 2 ie Ready to Paint the format
        fPaint = 2
        frmMain.rtfText.SelLength = 0
    
    ElseIf fPaint = 2 Then
        
        ' Copies the format to the selected text
        If fBold = True Then
            frmMain.rtfText.SelBold = True
        Else
            frmMain.rtfText.SelBold = False
        End If
        
        If fItalics = True Then
            frmMain.rtfText.SelItalic = True
        Else
            frmMain.rtfText.SelItalic = False
        End If
        If fUnderline = True Then
            frmMain.rtfText.SelUnderline = True
        Else
            frmMain.rtfText.SelUnderline = False
        End If
        If fStrikeThru = True Then
            frmMain.rtfText.SelStrikeThru = True
        Else
            frmMain.rtfText.SelStrikeThru = False
        End If
        
        frmMain.rtfText.SelFontName = fFont
        frmMain.rtfText.SelFontSize = fSize
        frmMain.rtfText.SelColor = fColor
        
        ' fPaint 3 means Job Done
        fPaint = 3
        
        ' Make the button reusable
        frmMain.tbToolBar.Buttons("FPt").Value = tbrUnpressed
        
        ' Change the Mouse pointer to default
       
            frmMain.rtfText.MousePointer = rtfDefault


    
    End If
    
End Sub


Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
    '** Description:
    '** Print the active document
    On Error GoTo PrintError
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long

    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips

    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin

    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text

    ' Get length of text in RTF
    TextLength = Len(RTF.Text)

    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        Printer.NewPage                  ' Move on to next page
        Printer.Print Space(1) ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
    Loop

    ' Commit the print job
    Printer.EndDoc

    ' Allow the RTF to free up memory
    r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    Exit Sub
PrintError:
    MsgBox "Unexpected Error in printing"
End Sub


Private Sub chkMru(x As Integer)
     If BMruFile(x) <> "" Then
        frmMain.mnuFileMRU(x).Caption = BMruFile(x)
        frmMain.mnuFileMRU(x).Visible = True
     End If
End Sub

Public Sub BMruInit()
    Dim I As Integer
    BMruFile(0) = GetSetting(App.Title, "Settings", "File0", "")
    BMruFile(1) = GetSetting(App.Title, "Settings", "File1", "")
    BMruFile(2) = GetSetting(App.Title, "Settings", "File2", "")
    BMruFile(3) = GetSetting(App.Title, "Settings", "File3", "")
    For I = 0 To 3
        chkMru I
    Next
End Sub

Public Sub storeMRU(f As String)
    Dim I As Integer
    Dim ifSame As Boolean
    ifSame = False
    For I = 0 To 3
        If BMruFile(I) = f Then
            ifSame = True
        End If
    Next
    For I = 0 To 3
        If BMruFile(I) = "" Then
            BMruFile(I) = f
            Exit Sub
        End If
    Next
    If ifSame = False Then
        For I = 3 To 1 Step -1
            If BMruFile(I) = f Then Exit For
            BMruFile(I) = BMruFile(I - 1)
        Next
        BMruFile(0) = f
    End If
End Sub

Public Sub storeMRUinReg()
    SaveSetting App.Title, "Settings", "File0", BMruFile(0)
    SaveSetting App.Title, "Settings", "File1", BMruFile(1)
    SaveSetting App.Title, "Settings", "File2", BMruFile(2)
    SaveSetting App.Title, "Settings", "File3", BMruFile(3)
End Sub
