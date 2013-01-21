VERSION 5.00
Begin VB.Form frmCustomAngle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Select Custom Angle"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   Icon            =   "CustomAngle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   1980
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   390
      Left            =   1980
      TabIndex        =   3
      Top             =   495
      Width           =   975
   End
   Begin VB.HScrollBar hsbAngle 
      Height          =   210
      LargeChange     =   5
      Left            =   165
      Max             =   359
      TabIndex        =   1
      Top             =   2055
      Width           =   1620
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      Height          =   1530
      Left            =   180
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   0
      Top             =   465
      Width           =   1590
   End
   Begin VB.Label lblAngle 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   375
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmCustomAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCanceled  As Boolean
Private miAngle     As Integer
Private mGradient   As New clsGradient
Public Function Display(frmCaller As Form, fAngle As Single, ByVal lColor1 As Long, ByVal lColor2 As Long) As Boolean

    mbCanceled = False
    With mGradient
        .Angle = fAngle
        .Color1 = lColor1
        .Color2 = lColor2
        .Draw picDraw
        miAngle = CInt(.Angle)
    End With
    Call DrawAngle(picDraw, miAngle)
    lblAngle.Caption = CStr(miAngle) & Chr$(176)
    hsbAngle.Value = miAngle
    Me.Move (frmCaller.Width - Me.Width) / 2, (frmCaller.Height - Me.Height) / 2
    Me.Show vbModal, frmCaller
    If Not mbCanceled Then
        fAngle = CSng(miAngle)
        Display = True
    End If
    
End Function

Private Sub cmdCancel_Click()
    mbCanceled = True
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub hsbAngle_Change()

    miAngle = hsbAngle.Value
    lblAngle.Caption = CStr(miAngle) & Chr$(176)
    mGradient.Angle = CSng(miAngle)
    mGradient.Draw picDraw
    Call DrawAngle(picDraw, miAngle)
    picDraw.Refresh
    
End Sub


Private Sub hsbAngle_Scroll()
    hsbAngle_Change
End Sub


