VERSION 5.00
Begin VB.Form frmDocPreview 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previsualización"
   ClientHeight    =   6270
   ClientLeft      =   1050
   ClientTop       =   1500
   ClientWidth     =   7935
   Icon            =   "DocPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Height          =   405
      Left            =   840
      Picture         =   "DocPreview.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Print"
      Top             =   120
      Width           =   405
   End
   Begin VB.CommandButton cmdZoomIn 
      Height          =   405
      Left            =   1320
      Picture         =   "DocPreview.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Zoom in"
      Top             =   120
      Width           =   405
   End
   Begin VB.CommandButton cmdZoomOut 
      Height          =   405
      Left            =   1770
      Picture         =   "DocPreview.frx":0EE6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Zoom out"
      Top             =   120
      Width           =   405
   End
   Begin VB.ComboBox cboScale 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   135
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevPage 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Prev page"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      TabIndex        =   13
      ToolTipText     =   "Next page"
      Top             =   120
      Width           =   315
   End
   Begin VB.ComboBox cboPageNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "DocPreview.frx":0FD0
      Left            =   4920
      List            =   "DocPreview.frx":0FD2
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   825
   End
   Begin VB.TextBox txtTotalPages 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "txtTotalPages"
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      Picture         =   "DocPreview.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   45
      Width           =   585
   End
   Begin VB.PictureBox PicZ 
      BackColor       =   &H8000000D&
      Height          =   5355
      Left            =   0
      ScaleHeight     =   5295
      ScaleWidth      =   7605
      TabIndex        =   2
      Top             =   600
      Width           =   7665
      Begin VB.PictureBox Pic5 
         BackColor       =   &H80000009&
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   2595
         TabIndex        =   9
         Top             =   120
         Width           =   2655
      End
      Begin VB.PictureBox Pic4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   2715
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   3015
         TabIndex        =   8
         Top             =   120
         Width           =   3075
      End
      Begin VB.PictureBox Pic3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   3285
         Left            =   120
         ScaleHeight     =   3225
         ScaleWidth      =   3765
         TabIndex        =   7
         Top             =   0
         Width           =   3825
      End
      Begin VB.PictureBox Pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   3795
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   4515
         TabIndex        =   6
         Top             =   60
         Width           =   4575
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   4215
         Left            =   60
         ScaleHeight     =   4155
         ScaleWidth      =   5325
         TabIndex        =   5
         Top             =   30
         Width           =   5385
      End
      Begin VB.PictureBox PicX 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   4695
         Left            =   30
         ScaleHeight     =   4635
         ScaleWidth      =   6015
         TabIndex        =   4
         Top             =   0
         Width           =   6075
      End
      Begin VB.PictureBox picP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         Height          =   5310
         Left            =   0
         ScaleHeight     =   5250
         ScaleWidth      =   6885
         TabIndex        =   3
         Top             =   -30
         Width           =   6945
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5415
      LargeChange     =   10
      Left            =   7680
      Max             =   200
      TabIndex        =   0
      Top             =   600
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      LargeChange     =   10
      Left            =   0
      Max             =   200
      TabIndex        =   1
      Top             =   6000
      Width           =   7665
   End
End
Attribute VB_Name = "frmDocPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Const SRCCOPY = &HCC0020


'-------------------------------------------------------------------------------------------------------------------
' By using the following messages in VB, it is possible to make a RichTextBox support WYSIWYG display and output:
' EM_SETTARGETDEVICE message is used to tell a RichTextBox to base its display on a target device.
' EM_FORMATRANGE message sends a page at a time to an output device using the specified coordinates.

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    firstChar As Long         ' First character of range (0 for start of doc)
    lastChar As Long          ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long               ' Actual DC to draw on
    hdcTarget As Long         ' Target DC for determining text formatting
    rectRegion As Rect        ' Region of the DC to draw to (in twips)
    rectPage As Rect          ' Page size of the entire DC (in twips)
    mCharRange As CharRange   ' Range of text to draw (see above user type)
End Type


Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, Ip As Any) As Long
     
Dim mFormatRange As FormatRange
Dim rectDrawTo As Rect
Dim rectPage As Rect
Dim TextLength As Long
Dim newStartPos As Long
Dim dumpaway As Long
     
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
     (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
     ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
'-------------------------------------------------------------------------------------------------------------------

Dim mNotShow As Boolean
Dim mSizeNo As Integer
Dim mTotalPages As Integer



Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   
   gprint = False
   
   
     ' we don't want the sizes to change after they have been appropriately sized
   PicZ.AutoSize = False             ' Base, always visible
   picP.AutoSize = False             ' For print intermediary, always invisible
   PicX.AutoSize = False             ' For display intermediary, always invisible
   Pic1.AutoSize = False             ' As 150%
   Pic2.AutoSize = False             ' As 100%
   Pic3.AutoSize = False             ' As 75%
   Pic4.AutoSize = False             ' As 50%
   Pic5.AutoSize = False             ' As 25%
   
   
       ' By default VB prints in twips. If a Picturebox is using pixels, we have to
       ' convert twips to pixels.  Therefore we fix the size of Pictureboxes before
       ' setting its ScaleMode to pixel (Eash pixel is about 15 twips, depending on
       ' the resolution of device)
      
   Dim mNormalWidth, mNormalHeight
   Dim mAdjFactor
   Dim mRect, mNewRect, mfactor
   Dim mpage As Integer
   
      ' Render document size in line with that of the printer (but note that doc is
      ' shown on screen without print margins)
   DocWYSIWYG frmMain.ActiveControl ' frmMain.ActiveForm.ActiveControl
   
      ' Obtain size of the printer
   mNormalWidth = Printer.ScaleWidth
   mNormalHeight = Printer.ScaleHeight
   
      ' Due to diff of resolution between screen and printer, we may use an adjustment
      ' factor, here we don't have any adjustment
   mAdjFactor = 100 / 100
   
   mNormalWidth = mNormalWidth * mAdjFactor
   mNormalHeight = mNormalHeight * mAdjFactor
   
      ' Mark down rectangle area, see remarks later
   mRect = mNormalWidth * mNormalHeight
   
      ' Make the invisible PicX of the same size as printer
   PicX.Width = mNormalWidth
   PicX.Height = mNormalHeight
   
  
     ' Percentage may be expressed in terms of original area (in that case, we have
     ' to derive the width and height from the computed area), or in terms of width
     ' and height themselves.  Here, to stress the point, we apply the percentage
     ' in terms of the area for sizes over 100%, but apply the percentage in terms
     ' of the width and height themselves for sizes are below 100%.
   
       ' Set 150%
   mNewRect = mRect * (150 / 100)
     ' By what percentage (factor) the width and the height should be reduced in order
     ' to arrive at an area for the new rectangle?
     ' (mNormalWidth * mfactor) * (mNormalHeight * mfactor) = mNewRect (mfactor Square)
     ' * (mNormalWidth * mNormalHeight) = mNewRect
   mfactor = Sqr(mNewRect / (mNormalWidth * mNormalHeight))
   Pic1.Width = CInt(mNormalWidth * mfactor)
   Pic1.Height = CInt(mNormalHeight * mfactor)
   
       ' Set 100%
   Pic2.Width = PicX.Width
   Pic2.Height = PicX.Height
       
      ' Re remarks earlier, we choose not to derive width and height from area for
      ' sizes below 100%.
       ' Set 75%
   Pic3.Width = CInt(mNormalWidth * 75 / 100)
   Pic3.Height = CInt(mNormalHeight * 75 / 100)
   
       ' Set 50%
   Pic4.Width = CInt(mNormalWidth * 50 / 100)
   Pic4.Height = CInt(mNormalHeight * 50 / 100)
   
       ' Set 25%
   Pic5.Width = CInt(mNormalWidth * 25 / 100)
   Pic5.Height = CInt(mNormalHeight * 25 / 100)
   
     ' Set ScaleMode to pixels.
   frmDocPreview.ScaleMode = vbPixels
   PicZ.ScaleMode = vbPixels
   PicX.ScaleMode = vbPixels
   Pic1.ScaleMode = vbPixels
   Pic2.ScaleMode = vbPixels
   Pic3.ScaleMode = vbPixels
   Pic4.ScaleMode = vbPixels
   Pic5.ScaleMode = vbPixels
   
     ' Set AutoRedraw to True
   PicZ.AutoRedraw = True
   picP.AutoRedraw = True
   PicX.AutoRedraw = True
   Pic1.AutoRedraw = True
   Pic2.AutoRedraw = True
   Pic3.AutoRedraw = True
   Pic4.AutoRedraw = True
   Pic5.AutoRedraw = True
   
    ' Set BorderStyle to Fixed Single
   PicZ.BorderStyle = 1
   PicX.BorderStyle = 1
   Pic1.BorderStyle = 1
   Pic2.BorderStyle = 1
   Pic3.BorderStyle = 1
   Pic4.BorderStyle = 1
   Pic5.BorderStyle = 1
   
    ' Set Fillstyle to Transparent
   PicZ.FillStyle = 1
   picP.FillStyle = 1
   PicX.FillStyle = 1
   Pic1.FillStyle = 1
   Pic2.FillStyle = 1
   Pic3.FillStyle = 1
   Pic4.FillStyle = 1
   Pic5.FillStyle = 1
   

   ' Backcolor of PicZ is blue (&H8000000D), the rest are white (&H80000009)
   PicZ.BackColor = &H8000000D
   picP.BackColor = &H80000009
   PicX.BackColor = &H80000009
   Pic1.BackColor = &H80000009
   Pic2.BackColor = &H80000009
   Pic3.BackColor = &H80000009
   Pic4.BackColor = &H80000009
   Pic5.BackColor = &H80000009
   

    ' Before showing first page, test how many pages are there in total in RTB.
   mTotalPages = PageCtnProc(frmDocPreview.PicX)
    ' Display the No. of total pages available
   txtTotalPages.Text = "Total " & CStr(mTotalPages) & " páginas"
    ' Enable/disable page movement buttons
   setPageButtons
   
   Dim I As Integer
   cboPageNo.Clear
   For I = 1 To mTotalPages
       cboPageNo.AddItem I
   Next I
   cboPageNo.Text = cboPageNo.List(0)
   
   
      ' Set max of scroll bars
   VScroll1.Max = 1000
   HScroll1.Max = 1000
    
      ' For ComboBox list
    cboScale.AddItem "150"
    cboScale.AddItem "100"
    cboScale.AddItem "75"
    cboScale.AddItem "50"
    cboScale.AddItem "25"
    cboScale.Text = cboScale.List(4)      ' i.e. 25%
    
    
      ' Instead Selprint whole document content such as:
      '   frmmain.ActiveForm.ActiveControl.SelPrint frmDocPreview.picX.Hdc
      ' we only print a single page at a time.  Initially we show page 1.
      '
      ' Whatever page, we will print it to PicX first (then project to other
      ' pictureboxes according to the sizes they play)
   mpage = 1
   FormPreviewPage frmDocPreview.PicX, mpage
   
    
     ' Now stretchblt to wanted sizes.
    For I = 1 To 5
        DoEvents
        If MakeSizes(I) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Next
    Screen.MousePointer = vbDefault
     
     ' Start display of preview screen.
     ' Note picZ is always visible, picX always not.
    PicZ.Visible = True
    picP.Visible = False
    PicX.Visible = False
    
    mNotShow = False        ' Show appropriate picture on screen
    mSizeNo = 5             ' i.e. cboScale.List=4, 25%
    ChangePreview
    
End Sub




Private Sub cboPageNo_click()
    Dim mpage As Integer
    mpage = cboPageNo.ListIndex + 1
    setPageButtons
    
    Screen.MousePointer = vbHourglass
    
     ' Print a new page to PicX
    FormPreviewPage frmDocPreview.PicX, mpage
     ' Again have to stretchblt to various sizes.
    Dim I
    For I = 1 To 5
        DoEvents
        If MakeSizes(I) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Next
    
     ' Have to change size (and then change back) to refresh display of new screen
     ' During the change, not to show any picture, hence mNotShow is temporarily
     ' set to True
    If mSizeNo = 1 Then
        mSizeNo = 2
        mNotShow = True
        ChangePreview
        mNotShow = False
        mSizeNo = 1
        ChangePreview
    Else
        mSizeNo = mSizeNo - 1
        mNotShow = True
        ChangePreview
        mNotShow = False
        mSizeNo = mSizeNo + 1
        ChangePreview
    End If
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdPrevPage_Click()
    If mTotalPages = 1 Then
        Exit Sub
    Else
        If Val(cboPageNo.Text) > 1 Then
            cboPageNo.Text = cboPageNo.List(cboPageNo.ListIndex - 1)
            cboPageNo_click
        End If
    End If
End Sub



Private Sub cmdNextPage_Click()
    If mTotalPages = 1 Then
        Exit Sub
    Else
        If Val(cboPageNo.Text) < mTotalPages Then
             cboPageNo.Text = cboPageNo.List(cboPageNo.ListIndex + 1)
             cboPageNo_click
        End If
    End If
End Sub



Private Sub setPageButtons()
    If mTotalPages = 1 Then
        cmdPrevPage.Enabled = False
        cmdNextPage.Enabled = False
    Else
        If Val(cboPageNo.Text) = 1 Then
             cmdPrevPage.Enabled = False
             cmdNextPage.Enabled = True
        ElseIf Val(cboPageNo.Text) = mTotalPages Then
             cmdPrevPage.Enabled = True
             cmdNextPage.Enabled = False
        Else
             cmdPrevPage.Enabled = True
             cmdNextPage.Enabled = True
        End If
    End If
End Sub
Private Sub HScroll1_Change()
   Select Case mSizeNo
      Case 1
          Pic1.Left = -HScroll1.Value
      Case 2
          Pic2.Left = -HScroll1.Value
      Case 3
          Pic3.Left = -HScroll1.Value
      Case 4
          Pic4.Left = -HScroll1.Value
      Case 5
          Pic5.Left = -HScroll1.Value
   End Select
End Sub
Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub


Private Sub VScroll1_Change()
   Select Case mSizeNo
      Case 1
          Pic1.Top = -VScroll1.Value
      Case 2
          Pic2.Top = -VScroll1.Value
      Case 3
          Pic3.Top = -VScroll1.Value
      Case 4
          Pic4.Top = -VScroll1.Value
      Case 5
          Pic5.Top = -VScroll1.Value
   End Select
End Sub



Private Sub ChangePreview()
   Select Case mSizeNo
      Case 1
          If mNotShow = False Then
               Pic1.Visible = True
          Else
               Pic1.Visible = False
          End If
          Pic2.Visible = False
          Pic3.Visible = False
          Pic4.Visible = False
          Pic5.Visible = False
      Case 2
          Pic1.Visible = False
          If mNotShow = False Then
               Pic2.Visible = True
          Else
               Pic2.Visible = False
          End If
          Pic2.Visible = True
          Pic3.Visible = False
          Pic4.Visible = False
          Pic5.Visible = False
      Case 3
          Pic1.Visible = False
          Pic2.Visible = False
          If mNotShow = False Then
               Pic3.Visible = True
          Else
               Pic3.Visible = False
          End If
          Pic4.Visible = False
          Pic5.Visible = False
      Case 4
          Pic1.Visible = False
          Pic2.Visible = False
          Pic3.Visible = False
          If mNotShow = False Then
               Pic4.Visible = True
          Else
               Pic4.Visible = False
          End If
          Pic5.Visible = False
      Case 5
          Pic1.Visible = False
          Pic2.Visible = False
          Pic3.Visible = False
          Pic4.Visible = False
          If mNotShow = False Then
               Pic5.Visible = True
          Else
               Pic5.Visible = False
          End If
   End Select
End Sub



' Combo does not honour "Change", we use "Click" instead
Private Sub cboScale_Click()
    Select Case cboScale.Text
        Case "150"
            mSizeNo = 1
            cmdZoomIn.Enabled = False
            cmdZoomOut.Enabled = True
        Case "100"
            mSizeNo = 2
        Case "75"
            mSizeNo = 3
        Case "50"
            mSizeNo = 4
        Case "25"
            mSizeNo = 5
            cmdZoomIn.Enabled = True
            cmdZoomOut.Enabled = False
    End Select
    If mSizeNo > 1 And mSizeNo < 5 Then
         cmdZoomIn.Enabled = True
         cmdZoomOut.Enabled = True
    End If
    ChangePreview
End Sub


Private Sub cmdPrint_Click()
     gprint = True
     Unload Me
End Sub



Private Sub cmdZoomin_click()
     If mSizeNo = 1 Then
          Exit Sub
     End If
     Select Case mSizeNo
          Case 5
               mSizeNo = 4
               cboScale.Text = cboScale.List(3)
               cmdZoomOut.Enabled = True
          Case 4
               mSizeNo = 3
               cboScale.Text = cboScale.List(2)
          Case 3
               mSizeNo = 2
               cboScale.Text = cboScale.List(1)
          Case 2
               mSizeNo = 1
               cboScale.Text = cboScale.List(0)
               cmdZoomIn.Enabled = False
     End Select
     If mSizeNo > 1 And mSizeNo < 5 Then
              cmdZoomIn.Enabled = True
              cmdZoomOut.Enabled = True
     End If
     ChangePreview
End Sub



Private Sub cmdzoomout_click()
    If mSizeNo = 5 Then
         Exit Sub
    End If
    Select Case mSizeNo
         Case 1
              cmdZoomIn.Enabled = True
              mSizeNo = 2
              cboScale.Text = cboScale.List(1)
         Case 2
              mSizeNo = 3
              cboScale.Text = cboScale.List(2)
         Case 3
              mSizeNo = 4
              cboScale.Text = cboScale.List(3)
         Case 4
              mSizeNo = 5
              cboScale.Text = cboScale.List(4)
              cmdZoomOut.Enabled = False
              cmdZoomIn.Enabled = True
     End Select
     If mSizeNo > 1 And mSizeNo < 5 Then
              cmdZoomIn.Enabled = True
              cmdZoomOut.Enabled = True
     End If
     ChangePreview
End Sub



Private Function MakeSizes(ByVal mofSize As Integer) As Boolean
    Dim SrcX As Long, SrcY As Long
    Dim DestX As Long, DestY As Long
    Dim SrcWidth As Long, SrcHeight As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim SrcHDC As Long, DestHDC As Long
    Dim mresult
      
    SrcX = 0: SrcY = 0: DestX = 0: DestY = 0
      
    SrcWidth = PicX.ScaleWidth
    SrcHeight = PicX.ScaleHeight
    SrcHDC = PicX.hdc
   
   Select Case mofSize
       Case 1
          DestWidth = Pic1.ScaleWidth
          DestHeight = Pic1.ScaleHeight
          DestHDC = Pic1.hdc
          
      Case 2
          DestWidth = Pic2.ScaleWidth
          DestHeight = Pic2.ScaleHeight
          DestHDC = Pic2.hdc
       
      Case 3
          DestWidth = Pic3.ScaleWidth
          DestHeight = Pic3.ScaleHeight
          DestHDC = Pic3.hdc
          
      Case 4
          DestWidth = Pic4.ScaleWidth
          DestHeight = Pic4.ScaleHeight
          DestHDC = Pic4.hdc
      Case 5
          DestWidth = Pic5.ScaleWidth
          DestHeight = Pic5.ScaleHeight
          DestHDC = Pic5.hdc
   End Select

   mresult = StretchBlt(DestHDC, DestX, DestY, DestWidth, DestHeight, SrcHDC, _
      SrcX, SrcY, SrcWidth, SrcHeight, SRCCOPY)

   If mresult = 0 Then
       MsgBox "Error al cambiar el tamaño de las imágenes. Imposible continuar"
       MakeSizes = False
   Else
       MakeSizes = True
   End If
End Function




Private Sub cmdClose_Click()
    Unload Me
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' To display the same as it would print on the selected printer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function DocWYSIWYG(RTB As Control) As Long
     Dim LeftMargin As Long, RightMargin As Long
     Dim linewidth As Long
     Dim PrinterhDC As Long
     Dim r As Long
     Printer.ScaleMode = vbTwips

     LeftMargin = gLeftMargin * 1440
     RightMargin = Printer.Width - gRightMargin * 1440

     linewidth = RightMargin - LeftMargin

     DocWYSIWYG = linewidth
End Function




Sub FormPreviewPage(inControl As Control, InPage As Integer)
    Dim PageCtn
    
      ' Clear picture box control
    Set inControl.Picture = LoadPicture

      ' Set printable area rect.
      ' Note in frmDocPreview, scaleModes are all in vbPixels,
      ' have to compute the vbtwips equivalent
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = inControl.Width * Screen.TwipsPerPixelX
    rectPage.Bottom = inControl.Height * Screen.TwipsPerPixelY
 
      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = gLeftMargin * 1440
    rectDrawTo.Top = gTopMargin * 1440
    rectDrawTo.Right = inControl.Width * Screen.TwipsPerPixelX _
         - gRightMargin * 1440
    rectDrawTo.Bottom = inControl.Height * Screen.TwipsPerPixelY _
         - gBottomMargin * 1440
 
    mFormatRange.hdc = inControl.hdc           ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = inControl.hdc     ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo       ' Area on page to draw to
    mFormatRange.rectPage = rectPage           ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0      ' Start of text
    mFormatRange.mCharRange.lastChar = -1      ' End of the text

    'TextLength = Len(frmMain.ActiveForm.ActiveControl.Text)
    TextLength = Len(frmMain.ActiveControl.Text)

    PageCtn = 1
    Do
        newStartPos = SendMessage(frmMain.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
        'newStartPos = SendMessage(frmMain.ActiveForm.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
      
        If newStartPos >= TextLength Then
            Exit Do
        End If
        If PageCtn = InPage Then
            Exit Do
        End If
        
        ' Clear picture box control
        Set inControl.Picture = LoadPicture
       
        mFormatRange.mCharRange.firstChar = newStartPos       ' Starting position for next page
        
        mFormatRange.hdc = inControl.hdc
        mFormatRange.hdcTarget = inControl.hdc
        
        PageCtn = PageCtn + 1
        DoEvents
    Loop

    dumpaway = SendMessage(inControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub



' Test how many pages are there in total
Function PageCtnProc(inControl As Control) As Integer
    Dim mPageCtn As Integer
    
      ' Set printable area rect.
      ' Note in frmDocPreview, scaleModes are all in vbPixels;
      ' convert them to vbtwips.
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = inControl.Width * Screen.TwipsPerPixelX
    rectPage.Bottom = inControl.Height * Screen.TwipsPerPixelY
 
      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = gLeftMargin * 1440
    rectDrawTo.Top = gTopMargin * 1440
    rectDrawTo.Right = inControl.Width * Screen.TwipsPerPixelX _
         - gRightMargin * 1440
    rectDrawTo.Bottom = inControl.Height * Screen.TwipsPerPixelY _
         - gBottomMargin * 1440
 
      ' Set up the print instructions
    mFormatRange.hdc = inControl.hdc            ' Use the same DC for measuring and rendering
    mFormatRange.hdcTarget = inControl.hdc      ' Point at hDC
    mFormatRange.rectRegion = rectDrawTo        ' Area on page to draw to
    mFormatRange.rectPage = rectPage            ' Entire size of page
    mFormatRange.mCharRange.firstChar = 0       ' Start of text
    mFormatRange.mCharRange.lastChar = -1       ' End of the text

TextLength = Len(frmMain.ActiveControl.Text)
    TextLength = Len(frmMain.ActiveControl.Text)

    mPageCtn = 1
    Do
          ' Print the page by sending EM_FORMATRANGE message
        newStartPos = SendMessage(frmMain.rtfText.hwnd, EM_FORMATRANGE, True, mFormatRange)
         '  newStartPos = SendMessage(frmMain.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
     
        If newStartPos >= TextLength Then
            Exit Do
        End If
        mFormatRange.mCharRange.firstChar = newStartPos       ' Starting position for next page
        mFormatRange.hdc = inControl.hdc
        mFormatRange.hdcTarget = inControl.hdc
        
        mPageCtn = mPageCtn + 1
        DoEvents
    Loop
    
     ' Clear picture box control
    Set inControl.Picture = LoadPicture

    dumpaway = SendMessage(inControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    
    PageCtnProc = mPageCtn
End Function




Sub DocPrintProc()
    On Error Resume Next
    DoEvents
    
      ' Clear picture box control
    Set frmDocPreview.picP.Picture = LoadPicture
    
    Dim mydialog1 As Object
    Dim mFromPage As Integer, mToPage As Integer, mpage As Integer
    
    Set mydialog1 = frmMain.dlgCommonDialog
    mydialog1.DialogTitle = "Print"
    mydialog1.CancelError = True

       ' Allow user select page range
    mydialog1.flags = cdlPDReturnDC + cdlPDPageNums
       ' But default to one of these
    'If frmMain.ActiveForm.rtfText.SelLength = 0 Then
    '    mydialog1.flags = mydialog1.flags + cdlPDAllPages
    'Else
    '    mydialog1.flags = mydialog1.flags + cdlPDSelection
    'End If

    mydialog1.ShowPrinter
    
    If Err = MSComDlg.cdlCancel Then
         Exit Sub
    End If
    
    
    mFromPage = mydialog1.FromPage
    mToPage = mydialog1.ToPage

  '  If frmMain.ActiveForm.WindowState <> 1 Then
     '   DocWYSIWYG frmMain.ActiveForm.ActiveControl
    '    frmMain.ActiveForm.Move 0, 0
   ' Else
   '     MsgBox "Cannot proceed with minimized screen"
  '      Exit Sub
  '  End If
    
    'If MsgBox("Proceed to print", vbYesNo + vbQuestion) = vbNo Then
    '    Exit Sub
    'End If
    
    Printer.Print ""
    Printer.ScaleMode = vbTwips
    
      ' Set printable rect area
    rectPage.Left = 0
    rectPage.Top = 0
    rectPage.Right = Printer.ScaleWidth
    rectPage.Bottom = Printer.ScaleHeight

      ' Set rect in which to print (relative to printable area)
    rectDrawTo.Left = gLeftMargin * 1440
    rectDrawTo.Top = gTopMargin * 1440
    rectDrawTo.Right = Printer.ScaleWidth - gRightMargin * 1440
    rectDrawTo.Bottom = Printer.ScaleHeight - gBottomMargin * 1440

     ' Dump earlier pages if any to PicP before reaching first wanted page
    mFormatRange.hdc = frmDocPreview.picP.hdc
    mFormatRange.hdcTarget = frmDocPreview.picP.hdc
    
    newStartPos = 0                                   ' Next char to start
    mFormatRange.rectRegion = rectDrawTo              ' Area on page to draw to
    mFormatRange.rectPage = rectPage                  ' Entire size of page
    mFormatRange.mCharRange.firstChar = newStartPos   ' Start of text
    mFormatRange.mCharRange.lastChar = -1             ' End of the text

    TextLength = Len(frmMain.ActiveControl.Text)
    'TextLength = Len(frmMain.ActiveForm.ActiveControl.Text)

      ' Dumping if any
    mpage = 1
    Do
        If mpage = mFromPage Then
            Exit Do
        End If
        
        ' Don't clear picture box control here, unless you want to print
        ' from first page always.
        
          ' Print the page by sending EM_FORMATRANGE message
           newStartPos = SendMessage(frmMain.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
     
        'newStartPos = SendMessage(frmMain.ActiveForm.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
        
        If newStartPos >= TextLength Then
            Exit Do
        End If
        
        mFormatRange.mCharRange.firstChar = newStartPos             ' Starting position for next page
        
        mFormatRange.hdc = frmDocPreview.picP.hdc
        mFormatRange.hdcTarget = frmDocPreview.picP.hdc
        
        mpage = mpage + 1
        DoEvents
    Loop

       ' Must cleanse memory here before print, otherwise font will not be right
    dumpaway = SendMessage(Screen.ActiveForm.ActiveControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
    
    If newStartPos >= TextLength Then
        Exit Sub
    End If
        
    
       ' Have to reinitialize printer here
    Printer.Print ""
    Printer.ScaleMode = vbTwips
    
    
       ' Actual print to printer, starting from the user-selected Page No.
    mFormatRange.hdc = Printer.hdc
    mFormatRange.hdcTarget = Printer.hdc
    
      ' Update char range
    mFormatRange.mCharRange.firstChar = newStartPos
    
    Do
          ' Print the page by sending EM_FORMATRANGE message
        
           'newStartPos = SendMessage(frmMain.ActiveForm.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
     
        newStartPos = SendMessage(frmMain.ActiveControl.hwnd, EM_FORMATRANGE, True, mFormatRange)
        
        If newStartPos >= TextLength Then
            Exit Do
        End If
        If mpage = mToPage Then
            Exit Do
        End If
        
        mFormatRange.mCharRange.firstChar = newStartPos              ' Starting position for next page
        
        Printer.NewPage                  ' Move on to next page
        Printer.Print ""                 ' Re-initialize hDC
        mFormatRange.hdc = Printer.hdc
        mFormatRange.hdcTarget = Printer.hdc
        
        mpage = mpage + 1
        DoEvents
    Loop

      ' Commit the print job
    Printer.EndDoc

      ' Free up memory
    dumpaway = SendMessage(Screen.ActiveForm.ActiveControl.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub


Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
