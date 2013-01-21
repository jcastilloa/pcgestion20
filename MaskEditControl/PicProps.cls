VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Private mPicture        As Picture
Private mposX           As Single
Private mposY           As Single
Private mHeight         As Integer
Private mWidth          As Integer

Public Property Let Width(ByVal vData As Integer)
mWidth = vData
End Property

Public Property Get Width() As Integer
Width = mWidth
End Property

Public Property Let Height(ByVal vData As Integer)
mHeight = vData
End Property

Public Property Get Height() As Integer
Height = mHeight
End Property

Public Property Let posY(ByVal vData As Single)
mposY = vData
End Property

Public Property Get posY() As Single
posY = mposY
End Property

Public Property Let PosX(ByVal vData As Single)
mposX = vData
End Property

Public Property Get PosX() As Single
PosX = mposX
End Property

Public Property Set Picture(ByVal vData As Picture)
Set mPicture = vData
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_UserMemId = 0
Set Picture = mPicture
End Property

Private Sub Class_Terminate()
Set mPicture = Nothing
End Sub
