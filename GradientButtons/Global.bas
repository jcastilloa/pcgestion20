Attribute VB_Name = "mdlGlobal"
Option Explicit

Public Type gtypeRect
    Width   As Long
    Height  As Long
    Left    As Long
    Top     As Long
End Type


Public Function fAbout$()
    fAbout = LoadResString(101) & App.Major & "." & App.Minor
End Function
