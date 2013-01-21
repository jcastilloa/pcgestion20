Attribute VB_Name = "Module1"
Option Explicit


Public SysTime As SYSTEMTIME
Public SysTime2 As SYSTEMTIME

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
