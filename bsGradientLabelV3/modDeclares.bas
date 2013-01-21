Attribute VB_Name = "Module1"
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
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Sub FontStuff()
  On Error GoTo GetOut
  Me.Cls
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
  Dim FONTSIZE As Integer
  FONTSIZE = Val(txtSize.Text)

  F.lfEscapement = 10 * Val(txtDegree.Text) 'rotation angle, in tenths
  FontName = "Arial Black" + Chr$(0) 'null terminated
  F.lfFacename = FontName
  F.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(Me.hdc, hFont)
  CurrentX = 3930
  CurrentY = 3860
  Print "SParq"
  
'  Clean up, restore original font
  hFont = SelectObject(Me.hdc, hPrevFont)
  DeleteObject hFont
  
  Exit Sub
GetOut:
  Exit Sub

End Sub

Private Sub Command1_Click()
  FontStuff
End Sub


Private Sub txtDegree_Change()
   If Val(txtDegree) < 1 Then txtDegree = 1: Exit Sub
   If Val(txtDegree) > 360 Then txtDegree = 360: Exit Sub
   Command1_Click
End Sub

Private Sub txtsize_Change()
  If Not IsNumeric(txtSize.Text) Then txtSize.Text = "18"
End Sub

