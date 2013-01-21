VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetupForm"
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

Private mFormWidth            As Integer
Private mMinButtonWidth       As Integer
Private mTopFirstQuestion     As Integer
Private mMaxFormHight         As Integer

Public Property Let MaxFormHight(ByVal xMaxFormHight As Integer)
mMaxFormHight = xMaxFormHight
End Property

Public Property Get MaxFormHight() As Integer
MaxFormHight = mMaxFormHight
End Property

Public Property Let FormWidth(ByVal xFormWidth As Integer)
mFormWidth = xFormWidth
End Property

Public Property Get FormWidth() As Integer
FormWidth = mFormWidth
End Property

Public Property Let MinButtonWidth(ByVal xMinButtonWidth As Integer)
mMinButtonWidth = xMinButtonWidth
End Property

Public Property Get MinButtonWidth() As Integer
MinButtonWidth = mMinButtonWidth
End Property

Public Property Let TopFirstQuestion(ByVal xTopFirstQuestion As Integer)
mTopFirstQuestion = xTopFirstQuestion
End Property

Public Property Get TopFirstQuestion() As Integer
TopFirstQuestion = mTopFirstQuestion
End Property

Private Sub Class_Initialize()
mFormWidth = 0
mMinButtonWidth = 500
mTopFirstQuestion = 200
mMaxFormHight = 0
End Sub
