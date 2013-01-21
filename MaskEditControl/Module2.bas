Attribute VB_Name = "Module2"
Option Explicit

'Downloaded from the internet.
'Developed By Herman Liu, adjusted By Ted Schopenhouer

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                           and  Douwe Konings dkonings@xs4all.nl

'This sources may be used freely without the intention of commercial distribution.
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

'In other words, when you are developing a program for yourself or for
'a company without selling this product to thirt party's it's allowed to
'use this source code. When you, or the company you work for, sells the
'program then permission is needed!!!!!


Declare Function CreateFontIndirect Lib "gdi32" Alias _
       "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 33
End Type

Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' In order for Windows NT to work
Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Const GM_ADVANCED = 2

Sub DispOrigAuthor(inObj As Object)
    On Error Resume Next
    Dim l As LOGFONT
    Dim mFont As Long
    Dim mPrevFont As Long
    Dim i As Integer
    Dim origMode As Integer
    Dim x As Single, y As Single
    Dim tmpX As Single, tmpY As Single
    Dim mresult
     ' For Windows NT to work
    mresult = SetGraphicsMode(inObj.hdc, GM_ADVANCED)
    origMode = inObj.ScaleMode
    inObj.ScaleMode = vbPixels
    
    l.lfFaceName = "Areial" & Chr$(0)
    l.lfEscapement = 0
    l.lfHeight = 6.5 * -20 / Screen.TwipsPerPixelY
       
    mFont = CreateFontIndirect(l)
    mPrevFont = SelectObject(inObj.hdc, mFont)
    
    x = inObj.ScaleLeft + 6
    y = inObj.ScaleHeight - 13
    inObj.CurrentX = x
    inObj.CurrentY = y
    tmpX = x
    tmpY = y
    
    inObj.ForeColor = &H808080
    tmpX = tmpX + 1: tmpY = tmpY + 1
    inObj.CurrentX = tmpX
    inObj.CurrentY = tmpY
    inObj.Print cOrgName
            
    inObj.CurrentX = x
    inObj.CurrentY = y
    tmpX = x
    tmpY = y
    inObj.ForeColor = vbWhite
    tmpX = tmpX - 1: tmpY = tmpY - 1
    inObj.CurrentX = tmpX
    inObj.CurrentY = tmpY
    inObj.Print cOrgName
            
    inObj.CurrentX = x
    inObj.CurrentY = y
    inObj.ForeColor = vbBlack
    inObj.Print cOrgName
            
    mresult = SelectObject(inObj.hdc, mPrevFont)
    mresult = DeleteObject(mFont)
    inObj.ScaleMode = origMode
End Sub




