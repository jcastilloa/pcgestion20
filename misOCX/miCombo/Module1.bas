Attribute VB_Name = "miComboMod"
Option Explicit


Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F

