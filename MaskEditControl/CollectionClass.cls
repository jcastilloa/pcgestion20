VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuestionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
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

Private mColl As Collection

Public Function Add(Optional Caption As String, Optional BackColor As OLE_COLOR, Optional ForeColor As OLE_COLOR, Optional Alignment As flxAlignMent = LeftJustify, Optional Visible As Boolean = True, Optional Tag As String, Optional Font As StdFont) As lblQuestionsProps
Dim xQuestion As New lblQuestionsProps

With xQuestion
   .Caption = Caption
   .BackColor = BackColor
   .ForeColor = ForeColor
   .Alignment = Alignment
   .Tag = Tag
   If Not Font Is Nothing Then
      Set .Font = Font
   End If
End With

mColl.Add xQuestion
Set Add = xQuestion
Set xQuestion = Nothing
End Function

Public Property Get Item(Number As Integer) As lblQuestionsProps
Attribute Item.VB_UserMemId = 0
If Number < 1 Or Number > mColl.Count Then
   Err.Raise 200, , "Number " & CStr(Number) & " Question Do Not Exist! Number MUST be between 1 and " & CStr(mColl.Count)
End If
Set Item = mColl(Number)
End Property

Public Property Get Count() As Long
Count = mColl.Count
End Property

Private Sub Class_Initialize()
Set mColl = New Collection
End Sub

Private Sub Class_Terminate()
Set mColl = Nothing
End Sub

