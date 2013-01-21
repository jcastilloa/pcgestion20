Attribute VB_Name = "mdlAPI"
Option Explicit

'Private Declare Function serGDIPlus_GradRect& _
'                Lib "D:\C++\SerAPI\SerAPI_GDI+\Debug\SerAPI.dll" ( _
'                ByVal hWnd&, _
'                ByVal R1&, _
'                ByVal G1&, _
'                ByVal B1&, _
'                ByVal R2&, _
'                ByVal G2&, _
'                ByVal B2&, _
'                ByVal lAngle&, _
'                ByVal lPerc&)


Const BITSPIXEL = 12
Private Declare Function GetDeviceCaps& _
                Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal nIndex As Long)
'''''''''''''''''''''''''''''''''''''''''''''
Public Enum genumSysColors
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum
Private Declare Function GetSysColor& _
                Lib "user32" ( _
                ByVal nIndex As Long)
'''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Type BITMAPINFOHEADER   '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Const WS_CLIPSIBLINGS = &H4000000
'Private Declare Function SendMessage& _
'                Lib "user32" _
'                Alias "SendMessageA" ( _
'                ByVal hwnd As Long, _
'                ByVal wMsg As Long, _
'                ByVal wParam As Long, _
'                lParam As Any)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Const BLACKNESS& = &H42
Public Const DSTINVERT& = &H550009
Public Const MERGECOPY& = &HC000CA
Public Const MERGEPAINT& = &HBB0226
Public Const NOTSRCCOPY& = &H330008
Public Const NOTSRCERASE& = &H1100A6
Public Const PATCOPY& = &HF00021
Public Const PATINVERT& = &H5A0049
Public Const PATPAINT& = &HFB0A09
Public Const SRCCOPY& = &HCC0020
Public Const SRCPAINT& = &HEE0086
Public Const SRCINVERT& = &H660046
Public Const SRCERASE& = &H440328
Public Const SRCAND& = &H8800C6
Public Const WHITENESS& = &HFF0062

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const IMAGE_BITMAP = 0&

Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup

Public mProgress As Long         '% filter progress
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function fMakeEmboss(ByVal picImg&, ByVal Factor&, pProgress&, _
                            Optional EmbossLevel As geDisabledEmbossLevel = delEmbossNormal) _
                            As Boolean
    Dim hdcNew&
    Dim oldHand&
    Dim ret&
    Dim BytesPerScanLine&
    Dim PadBytesPerScanLine&
    
    On Error GoTo ErrOccurred
    
    Call GetObject(picImg, Len(PicInfo), PicInfo)
    hdcNew = CreateCompatibleDC(0&)
    oldHand = SelectObject(hdcNew, picImg)
    With DIBInfo.bmiHeader
        .biSize = 40
        .biWidth = PicInfo.bmWidth
        .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        PadBytesPerScanLine = BytesPerScanLine - ((.biWidth * .biBitCount) + 7) \ 8
        .biSizeImage = BytesPerScanLine * Abs(.biHeight)
    End With
    
    '   redimension  (BGR+pad,x,y)
    ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
    ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
    
    'get bytes
    ret = GetDIBits(hdcNew, picImg, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    ret = GetDIBits(hdcNew, picImg, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    
    If EmbossLevel = delEmbossNormal Then
        Call Emboss(pProgress, Factor)
    Else
        Call EmbossMore(pProgress, Factor)
    End If
    
    'copy bytes to device
    ret = SetDIBits(hdcNew, picImg, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
    Call SelectObject(hdcNew, oldHand)
    Call DeleteDC(hdcNew)
    
    ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
    ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
    
    fMakeEmboss = True
Exit_:
    Exit Function
ErrOccurred:
    fMakeEmboss = False
    Err.Clear
    Resume Exit_
End Function

Private Sub Emboss(ByRef pProgress As Long, ByVal BackCol As Long)
    Dim X As Long, Y As Long
    Dim R As Long, G As Long, B As Long
    Dim cB As Long, cG As Long, cR As Long
    
    On Error Resume Next
    mProgress = 0
    Call GetRGB(BackCol, cR, cG, cB)
    For Y = 1 To PicInfo.bmHeight - 1
        For X = 1 To PicInfo.bmWidth - 1
            B = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + cB)
            G = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + cG)
            R = Abs(CLng(iDATA(3, X, Y)) - CLng(iDATA(3, X + 1, Y + 1)) + cR)
            If R > 255 Then R = 255
            If R < 0 Then R = 0
            If G > 255 Then G = 255
            If G < 0 Then G = 0
            If B > 255 Then B = 255
            If B < 0 Then B = 0
            iDATA(1, X, Y) = B
            iDATA(2, X, Y) = G
            iDATA(3, X, Y) = R
        Next X
        mProgress = (Y * 100) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next Y
    pProgress = 100
    DoEvents
End Sub

Private Sub EmbossMore(ByRef pProgress As Long, ByVal BackCol As Long)
    Dim X As Long, Y As Long
    Dim R As Long, G As Long, B As Long
    Dim cB As Long, cG As Long, cR As Long
  
    mProgress = 0
    Call GetRGB(BackCol, cR, cG, cB)
    For Y = 2 To PicInfo.bmHeight - 1
        For X = 2 To PicInfo.bmWidth - 1
            B = CLng(bDATA(1, X - 1, Y - 1)) - CLng(bDATA(1, X + 1, Y - 1)) + _
                CLng(bDATA(1, X - 1, Y)) - CLng(bDATA(1, X + 1, Y)) + _
                CLng(bDATA(1, X - 1, Y + 1)) - CLng(bDATA(1, X + 1, Y + 1)) + cB
            G = CLng(bDATA(2, X - 1, Y - 1)) - CLng(bDATA(2, X + 1, Y - 1)) + _
                CLng(bDATA(2, X - 1, Y)) - CLng(bDATA(2, X + 1, Y)) + _
                CLng(bDATA(2, X - 1, Y + 1)) - CLng(bDATA(2, X + 1, Y + 1)) + cG
            R = CLng(bDATA(3, X - 1, Y - 1)) - CLng(bDATA(3, X + 1, Y - 1)) + _
                CLng(bDATA(3, X - 1, Y)) - CLng(bDATA(3, X + 1, Y)) + _
                CLng(bDATA(3, X - 1, Y + 1)) - CLng(bDATA(3, X + 1, Y + 1)) + cR
            If R > 255 Then R = 255
            If R < 0 Then R = 0
            If G > 255 Then G = 255
            If G < 0 Then G = 0
            If B > 255 Then B = 255
            If B < 0 Then B = 0
            iDATA(1, X, Y) = B
            iDATA(2, X, Y) = G
            iDATA(3, X, Y) = R
        Next X
        mProgress = (Y * 100) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next Y
    pProgress = 100
    DoEvents
End Sub

'-------------------------------------------AUXILIARY
Private Sub GetRGB(ByVal Col As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
    R = Col Mod 256
    G = ((Col And &HFF00&) \ 256&) Mod 256&
    B = (Col And &HFF0000) \ 65536
End Sub

Public Function fGetBitsPerPixel&(hdc&)
    fGetBitsPerPixel& = GetDeviceCaps(hdc&, BITSPIXEL)
End Function

Public Sub fGetRGB(ByVal lColor&, ByRef lRed&, ByRef lGrn&, ByRef lBlu&)
    
    lRed = (lColor And &HFF&)
    lGrn = (lColor And &HFF00&) / &H100
    lBlu = (lColor And &HFF0000) / &H10000
    
    'fGetRGB& = RGB(lRed, lGrn, lBlu)
End Sub

Public Function fGetSysColor&(eColor As genumSysColors)
    fGetSysColor& = GetSysColor(eColor)
End Function

Public Function fHighWord%(DWord&)
    fHighWord% = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function fLowWord%(DWord&)
    If DWord And &H8000& Then ' &H8000& = &H00008000
        fLowWord% = DWord Or &HFFFF0000
    Else
        fLowWord% = DWord And &HFFFF&
    End If
End Function

Public Sub fTest(lColor&)
    Dim lR1&, lG1&, lB1&
    Call fGetRGB&(lColor, lR1&, lG1&, lB1&)
    'Debug.Print "red:"; lR1, "green:"; lG1, "blue:"; lG1
End Sub

Public Function fMixTwoColors256&(ByVal lColor1&, ByVal lColor2&)
    Dim lR1&, lG1&, lB1&
    Dim lR2&, lG2&, lB2&
    Dim lTemp&
    
    If CBool(fHighWord(lColor1)) Then
        lTemp& = fLowWord(lColor1)
        lColor1 = fGetSysColor(lTemp&)
    End If
    If CBool(fHighWord(lColor2)) Then
        lTemp& = fLowWord(lColor2)
        lColor2 = fGetSysColor(lTemp&)
    End If
    
    Call fGetRGB&(lColor1, lR1&, lG1&, lB1&)
    Call fGetRGB&(lColor2, lR2&, lG2&, lB2&)
    
    fMixTwoColors256& = RGB(fAVG(lR1&, lR2&), fAVG(lG1&, lG2&), fAVG(lB1&, lB2&))
End Function

Private Function fAVG&(ParamArray A())
    Dim X, lSum&, lCount&
    
    On Error GoTo ErrOccurred
    For Each X In A()
        lSum& = lSum& + X
        lCount& = lCount& + 1
    Next
    fAVG& = lSum / lCount
    
    Exit Function
ErrOccurred:
    fAVG& = 0
End Function

Public Sub fDrawDottedRect(P As PictureBox, lColor&, X1&, Y1&, X2&, Y2&)
    Dim i&, iTemp%
    Dim iStep%
    
    iStep = Screen.TwipsPerPixelX * 2
    
    'Const conStep = 30
    
    With P
        iTemp = .DrawWidth
        .DrawWidth = 1
        For i = X1 To X2 Step iStep 'conStep
            P.PSet (i, Y1), lColor
            P.PSet (i, Y2), lColor
        Next
        For i = Y1 To Y2 Step iStep 'conStep
            P.PSet (X1, i), lColor
            P.PSet (X2, i), lColor
        Next
        .DrawWidth = iTemp%
    End With
End Sub

Public Function fGetGradientColorRatio(ByVal lColor As OLE_COLOR, _
                                       dRatio#, _
                                       Optional Boost0 As Boolean = True) _
                                       As OLE_COLOR
    Dim lC1&, lC2&, lC3&, lCC1&, lCC2&, lCC3&, a1&, a2&, b1&, b2&, lReminder&
    Const con255& = 255&
    Const con0& = 0&
    'If CBool(fHighWord(lColor)) Then lColor = fGetSysColor(lColor)
    If lColor < 0 Then lColor = fGetSysColor(fLowWord(lColor))
    
    If dRatio = 1# Or lColor = 0 Then
        fGetGradientColorRatio = lColor
    ElseIf dRatio > 1# Then
        Call fGetRGB(lColor, lC1&, lC2&, lC3&)
        
        If (lC1 <> con0) And ((lC1 <= lC2 Or lC2 = con0) And (lC1 <= lC3 Or lC3 = con0)) Then
            lCC1 = CLng(lC1 * dRatio)
            If lCC1 >= con255 And Boost0 = False Then
                lCC1 = IIf(lC1 = con0, con0, con255)
                lCC2 = IIf(lC2 = con0, con0, con255)
                lCC3 = IIf(lC3 = con0, con0, con255)
                 
                GoTo ReturnResultColor
            ElseIf lCC1 >= con255 And Boost0 Then
                lReminder = lCC1 - con255
                If lReminder > con255 Then lReminder = con255
                
                lCC2 = IIf(lC2 = con0, lReminder, con255)
                lCC3 = IIf(lC3 = con0, lReminder, con255)
            Else
                a1 = lCC1 - lC1
                b1 = con255 - lC1
                
                '   color 2
                If lC2 = 0 And (Not Boost0) Then
                    lCC2 = 0
                Else
                    b2 = con255 - lC2
                    a2 = a1 * b2 / b1
                    lCC2 = lC2 + a2
                    If lCC2 > con255 Then lCC2 = con255
                End If
                
                '   color 3
                If lC3 = 0 Then
                    lCC3 = 0
                Else
                    b2 = con255 - lC3
                    a2 = a1 * b2 / b1
                    lCC3 = lC3 + a2
                    If lCC3 > con255 Then lCC3 = con255
                End If
            End If
        ElseIf (lC2 <> 0) And ((lC2 <= lC3 Or lC3 = 0) And (lC2 <= lC1 Or lC1 = 0)) Then
        'ElseIf (lC2 <= lC1 And lC2 <= lC3) Or (lC3 = 0) Then
            lCC2 = CLng(lC2 * dRatio)
            If lCC2 >= con255 And Boost0 = False Then
                lCC1 = IIf(lC1 = con0, con0, con255)
                lCC2 = IIf(lC2 = con0, con0, con255)
                lCC3 = IIf(lC3 = con0, con0, con255)
                 
                GoTo ReturnResultColor
            ElseIf lCC2 >= con255 And Boost0 Then
                lReminder = lCC1 - con255
                If lReminder > con255 Then lReminder = con255
                
                lCC1 = IIf(lC1 = con0, lReminder, con255)
                lCC3 = IIf(lC3 = con0, lReminder, con255)
            Else
                a1 = lCC2 - lC2
                b1 = con255 - lC2
                
                '   color 1
                If lC1 = 0 Then
                    lCC1 = 0
                Else
                    b2 = con255 - lC1
                    a2 = a1 * b2 / b1
                    lCC1 = lC1 + a2
                    If lCC1 > con255 Then lCC1 = con255
                End If
                
                '   color 3
                If lC3 = 0 Then
                    lCC3 = 0
                Else
                    b2 = con255 - lC3
                    a2 = a1 * b2 / b1
                    lCC3 = lC3 + a2
                    If lCC3 > con255 Then lCC3 = con255
                End If
            End If
        Else
            lCC3 = CLng(lC3 * dRatio)
            If lCC3 >= con255 And Boost0 = False Then
                lCC1 = IIf(lC1 = con0, con0, con255)
                lCC2 = IIf(lC2 = con0, con0, con255)
                lCC3 = IIf(lC3 = con0, con0, con255)
                 
                GoTo ReturnResultColor
            ElseIf lCC1 >= con255 And Boost0 Then
                lReminder = lCC1 - con255
                If lReminder > con255 Then lReminder = con255
                
                lCC1 = IIf(lC1 = con0, lReminder, con255)
                lCC2 = IIf(lC2 = con0, lReminder, con255)
            Else
                a1 = lCC3 - lC3
                b1 = con255 - lC3
                
                '   color 1
                If lC1 = 0 Then
                    lCC1 = 0
                Else
                    b2 = con255 - lC1
                    a2 = a1 * b2 / b1
                    lCC1 = lC1 + a2
                    If lCC1 > con255 Then lCC1 = con255
                End If
                
                '   color 2
                If lC2 = 0 Then
                    lCC2 = 0
                Else
                    b2 = con255 - lC2
                    a2 = a1 * b2 / b1
                    lCC2 = lC2 + a2
                    If lCC2 > con255 Then lCC2 = con255
                End If
            End If
        End If
ReturnResultColor:
        fGetGradientColorRatio = RGB(lCC1, lCC2, lCC3)
        Exit Function
    Else    '   dRatio < 1#
        Call fGetRGB(lColor, lC1&, lC2&, lC3&)
        If lC1 >= lC2 And lC1 >= lC3 Then
            lCC1 = CLng(lC1 * dRatio)
            If lCC1 = con0 Then
                lCC1 = con0: lCC2 = con0: lCC3 = con0
            Else
                a1 = lC1 - lCC1
                
                '   color 2
                a2 = a1 * lC2 / lC1
                lCC2 = lC2 - a2
                If lCC2 < con0 Then lCC2 = con0
                
                '   color 3
                a2 = a1 * lC3 / lC1
                lCC3 = lC3 - a2
                If lCC3 < con0 Then lCC3 = con0
            End If
        ElseIf lC2 >= lC1 And lC2 >= lC3 Then
            lCC2 = CLng(lC2 * dRatio)
            If lCC2 = con0 Then
                lCC1 = con0: lCC2 = con0: lCC3 = con0
            Else
                a1 = lC2 - lCC2
                
                '   color 1
                a2 = a1 * lC1 / lC2
                lCC1 = lC1 - a2
                If lCC1 < con0 Then lCC1 = con0
                
                '   color 3
                a2 = a1 * lC3 / lC2
                lCC3 = lC3 - a2
                If lCC3 < con0 Then lCC3 = con0
            End If
        Else
            lCC3 = CLng(lC3 * dRatio)
            If lCC3 = con0 Then
                lCC1 = con0: lCC2 = con0: lCC3 = con0
            Else
                a1 = lC3 - lCC3
                
                '   color 1
                a2 = a1 * lC1 / lC3
                lCC1 = lC1 - a2
                If lCC1 < con0 Then lCC1 = con0
                
                '   color 2
                a2 = a1 * lC2 / lC3
                lCC2 = lC2 - a2
                If lCC2 < con0 Then lCC2 = con0
            End If
        End If
        fGetGradientColorRatio = RGB(lCC1, lCC2, lCC3)
    End If
End Function


