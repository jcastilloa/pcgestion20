VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1950
      Left            =   525
      ScaleHeight     =   1890
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   330
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'In the code 128 barcode symbology, a character is represented
'an 11 digit binary number. This approach draws a black line
'for a '1' and a while line for a '0'. It is therefore extremely
'easy to draw a barcode for any  character in this symbology.
'There's also some housekeeping required and a modulus 103 checksum
'is required to validate the barcode. Try it out with a scanner -
'it works great!
'http://www.barcodeisland.com/code128.phtml has a good explanation
'of the code 128 smbology.
'Please vote if you find this code useful :-)

Option Explicit

Private varAllCodes     As Integer
Private varCodeIn()     As String
Private varTmp          As String

Private Sub DrawBarcode()

  Dim i As Integer
  Const cBarWidth As Integer = 1 'set the width of bars
  Const cBarHeight As Integer = 25 'set the height of the bars
  Const cQuietSpace As Integer = 10 'set the quiet space

  Form1.Cls
  Form1.DrawWidth = cBarWidth
  For i = 1 To Len(varTmp) 'let's create the actual code
    If Mid$(varTmp, i, 1) = 1 Then
      Form1.Line (i * cBarWidth + cQuietSpace, 2)-(i * cBarWidth + cQuietSpace, cBarHeight + 2), RGB(0, 0, 0)
    End If
    If Mid$(varTmp, i, 1) = 0 Then
      Form1.Line (i * cBarWidth + cQuietSpace, 2)-(i * cBarWidth + cQuietSpace, cBarHeight + 2), RGB(255, 255, 255)
    End If
  Next i
  
  Set Picture1.Picture = Form1.Picture
  DoEvents
  
  'Form1.Refresh
 ' DoEvents
 'Form1.Width = (cQuietSpace + Len(varTmp) + cQuietSpace) * cBarWidth 'adjust width to include current code
  'Picture1.Left = Form1.ScaleWidth / 2 - Picture1.ScaleWidth / 2 'centre the barcode on the form

End Sub


Public Sub actualiza(codigo As String)

Call Text1_Chnge(codigo)

End Sub
Private Sub Form_Load()

  Dim i As Integer
  
  i = -1
  Open App.Path & "\code128.txt" For Input As #1
  Do Until EOF(1)
    i = i + 1
    ReDim Preserve varCodeIn(i + 1)
    Line Input #1, varCodeIn(i)
  Loop
  varAllCodes = i
  Close #1
  'Text1_Chnge ("0000")

End Sub

Private Sub Text1_Chnge(codigo As String)

  Dim varChecksum     As String
  Dim varChk          As Integer
  Dim varData         As String
  Dim varCurrentChar  As Integer
  Dim i               As Integer
  Dim j               As Integer
  Const cCharacterSet As String = "11010010000" 'represents an alphanumeric characterset
  Const cStopChar     As String = "11000111010" 'code 128 stop character
  Const cTerminator   As String = "11"          'code 128 termination character

  If LenB(codigo) = 0 Then Exit Sub
  varChk = 104 'number of characterset we're using - must add to checksum
  For i = 1 To Len(codigo)
    varCurrentChar = Asc(Mid$(codigo, i, 1))
    For j = 0 To varAllCodes 'look at all codes
      If Val(Split(varCodeIn(j), ",")(1)) = varCurrentChar Then 'to find the one that corresponds to the ascii value of the current character
        varData = varData & Split(varCodeIn(j), ",")(2) 'get the code value from the array
        varChk = varChk + Val(Split(varCodeIn(j), ",")(0) * i) 'add values to create checksum
        Exit For 'we've found the code
      End If
    Next j
  Next i
  varChk = varChk Mod 103 'modulus 103 calculation to create checksum
  varChecksum = Split(varCodeIn(varChk), ",")(2) 'get the appropriate code
  varTmp = cCharacterSet & varData & varChecksum & cStopChar & cTerminator 'put it all together
  Call DrawBarcode

End Sub
