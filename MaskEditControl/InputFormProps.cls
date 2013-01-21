VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputFormProps"
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

Private mSetup          As SetupForm
Private mVisible        As Boolean
Private mPicture        As PictureProps
Private mIcon           As Picture
Private mCaption        As String
Private mBackColor      As OLE_COLOR

Public Property Let Visible(ByVal xVisible As Boolean)
mVisible = xVisible
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Public Property Set Picture(ByVal xPicture As PictureProps)
Set mPicture = xPicture
End Property

Public Property Get Picture() As PictureProps
Set Picture = mPicture
End Property

Public Property Set Icon(ByVal xIcon As Picture)
Set mIcon = xIcon
End Property

Public Property Get Icon() As Picture
Set Icon = mIcon
End Property

Public Property Let Caption(ByVal xCaption As String)
mCaption = xCaption
End Property

Public Property Get Caption() As String
Caption = mCaption
End Property

Public Property Let BackColor(ByVal xColor As OLE_COLOR)
mBackColor = xColor
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = mBackColor
End Property

Public Property Let Setup(ByVal xSetup As SetupForm)
Set mSetup = xSetup
End Property

Public Property Get Setup() As SetupForm
Set Setup = mSetup
End Property

Private Sub Class_Initialize()
Set mSetup = New SetupForm
Set mPicture = New PictureProps
End Sub

Private Sub Class_Terminate()
Set mSetup = Nothing
Set mPicture = Nothing
Set mIcon = Nothing
End Sub


