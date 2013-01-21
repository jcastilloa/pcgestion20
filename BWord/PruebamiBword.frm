VERSION 5.00
Object = "*\AmiBwordOCX.vbp"
Begin VB.Form PruebamiBword 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin miBwordOCX.miBword miBword1 
      Height          =   750
      Left            =   1275
      TabIndex        =   0
      Top             =   855
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "PruebamiBword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Me.miBword1.Carga

End Sub
