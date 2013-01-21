Attribute VB_Name = "Module"
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12Move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                           and  Douwe Konings dkonings@xs4all.nl

'This sources may be used freely without the intention of commercial distribution.
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

'In other words, when you are developing a program for yourself or for
'a company without selling this product to thirt party's it's allowed to
'use this source code. When you, or the company you work for, sells the
'program then permission is needed!!!!!

Public Enum UnLctl
   flxUnloadAll = 0
   flxQuestions = 1
   flxButtons = 2
End Enum



Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function HideCaret& Lib "user32" (ByVal hwnd As Long)





Public sDateFormats(1 To 15)     As String

Public Const cOrgName = "By Ted Schopenhouer"
Public Const WM_SETTINGCHANGE = &H1A
Public Const HWND_BROADCAST = &HFFFF&

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000

Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_ICURRDIGITS& = &H19
Public Const LOCALE_ICURRENCY = &H1B
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_IDAYLZERO = &H26
Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_IDIGITS& = &H11
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_IMEASURE = &HD
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_INEGCURR = &H1C
Public Const LOCALE_INEGSEPBYSPACE = &H57
Public Const LOCALE_INEGSIGNPOSN = &H53
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITLZERO = &H25
Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_S1159 = &H28
Public Const LOCALE_S2359 = &H29
Public Const LOCALE_SABBREVCTRYNAME = &H7
Public Const LOCALE_SABBREVDAYNAME1 = &H31
Public Const LOCALE_SABBREVDAYNAME2 = &H32
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SDATE& = &H1D
Public Const LOCALE_SDAYNAME1 = &H2A
Public Const LOCALE_SDAYNAME2 = &H2B
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_SINTLSYMBOL = &H15
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONGROUPING = &H18
Public Const LOCALE_SMONTHNAME1 = &H38
Public Const LOCALE_SMONTHNAME10 = &H41
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SNATIVECTRYNAME = &H8
Public Const LOCALE_SNATIVEDIGITS = &H13
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LOCALE_SNEGATIVESIGN = &H51
Public Const LOCALE_SPOSITIVESIGN = &H50
Public Const LOCALE_SSHORTDATE& = &H1F
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_STIMEFORMAT = &H1003


Public Function LocalDecimalSeperator() As String
Dim sBuffer As String * 50
Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, sBuffer, 50)
LocalDecimalSeperator = UCase$(Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1))
End Function


'Public Sub Get_locale(x As Long) ' Retrieve the regional setting
'
'Dim Symbol As String
'Dim iRet1 As Long
'Dim iRet2 As Long
'Dim lpLCDataVar As String
'Dim Pos As Integer
'Dim Locale As Long
'
'Locale = GetUserDefaultLCID()
'
'
'iRet1 = GetLocaleInfo(Locale, x, _
'lpLCDataVar, 0)
'Symbol = String$(iRet1, 0)
'
'iRet2 = GetLocaleInfo(Locale, x, Symbol, iRet1)
'Pos = InStr(Symbol, Chr$(0))
'If Pos > 0 Then
'   Symbol = Left$(Symbol, Pos - 1)
'   MsgBox "Regional Setting = " + Symbol
'End If
'
'End Sub
'
'Public Sub Set_locale() 'Change the regional setting
'
'Dim Symbol As String
'Dim iRet As Long
'Dim Locale As Long
'
'Locale = GetUserDefaultLCID() 'Get user Locale ID
'Symbol = "-" 'New character for the locale
'iRet = SetLocaleInfo(Locale, LOCALE_SDATE, Symbol)
'
'End Sub

Public Function Max(ByVal X As Long, ByVal Y As Long) As Long
Max = IIf(X > Y, X, Y)
End Function

Public Function Min(ByVal X As Long, ByVal Y As Long) As Long
Min = IIf(X > Y, Y, X)
End Function

Public Function AppVersion()
AppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function LocalDate() As String
Dim sBuffer    As String * 100
Dim i          As Integer
Dim i2         As Integer
Dim s(0 To 2)  As String
Dim m          As String
Dim d          As String
Dim Y          As String
Dim sTmp       As String

Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE&, sBuffer, 99)
sTmp = UCase$(Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1))

For i = 1 To Len(sTmp)

   Select Case Mid$(sTmp, i, 1)

      Case "D"
         If d = "" Then
            s(i2) = "DD"
            d = "d"
            i2 = i2 + 1
         End If
      Case "M"
         If m = "" Then
            s(i2) = "MM"
            m = "m"
            i2 = i2 + 1
         End If
      Case "Y"
         If Y = "" Then
            s(i2) = IIf(InStr(sTmp, "YYYY"), "YYYY", "YY")
            Y = "y"
            i2 = i2 + 1
         End If
   End Select

Next
LocalDate = s(0) & LocalDateSeperator & s(1) & LocalDateSeperator & s(2)
   
End Function

Public Function LocalDateSeperator() As String
Dim sBuffer As String * 50
Call GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE&, sBuffer, 99)
LocalDateSeperator = UCase$(Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1))
End Function


'this is the most powerfull search function ever written in VB

Public Function Token(sSearch As String, iFirstDeliInStr As Integer, Optional sSeperator As String = "^", Optional lStartPos As Long) As String
Dim l As Long
Dim i As Integer
Dim X As Long

Do While i <> iFirstDeliInStr
   lStartPos = InStr(lStartPos + 1, sSearch, sSeperator)
   If lStartPos = 0 Then
      lStartPos = Len(sSearch) + 1
      Exit Function
   End If
   i = i + 1
Loop

If lStartPos Then
   lStartPos = lStartPos + Len(sSeperator)
   l = InStr(lStartPos, sSearch, sSeperator)
Else
   l = InStr(sSearch, sSeperator)
   If l < 2 Then
      If l = 0 Then Token = sSearch
      lStartPos = lStartPos + 1
      Exit Function
   Else
      lStartPos = 1
   End If
End If

If l Then
   Token = Mid$(sSearch, lStartPos, l - lStartPos)
   lStartPos = l
Else
   Token = Mid$(sSearch, lStartPos)
   lStartPos = Len(sSearch) + 1
End If
End Function


