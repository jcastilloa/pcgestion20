VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonCollection"
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


Private mColl As Collection

Public Function Add(Optional Caption As String, Optional BackColor As OLE_COLOR, Optional Visible As Boolean = True, Optional ToolTipText As String, Optional Enabled As Boolean = True, Optional Picture As Picture, Optional Tag As String, Optional Font As StdFont) As cmdButtonsProps
Dim ButtonColl As New cmdButtonsProps

With ButtonColl
   .Caption = Caption
   .Visible = Visible
   .Enabled = Enabled
   .ToolTipText = ToolTipText
   .Tag = Tag
   Set .Picture = Picture
   .BackColor = BackColor
   If Not Font Is Nothing Then
      Set .Font = Font
   End If
End With

mColl.Add ButtonColl
Set Add = ButtonColl
Set ButtonColl = Nothing
End Function

Public Property Get Item(Number As Integer) As cmdButtonsProps
Attribute Item.VB_UserMemId = 0
If Number < 1 Or Number > mColl.Count Then
   Err.Raise 200, , "Number " & CStr(Number) & " Button Do Not Exist! Number MUST be between 1 and " & CStr(mColl.Count)
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

