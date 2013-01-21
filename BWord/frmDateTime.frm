VERSION 5.00
Begin VB.Form frmDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BWord - Fecha / Hora"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Insertar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraDateTime 
      Caption         =   "Formatos"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstDateTime 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    frmMain.rtfText.SelText = lstDateTime.Text
    Unload Me
End Sub
Private Sub Form_Load()
Call CenterForm(Me)
lstDateTime.AddItem Format(Now, "long Time")
lstDateTime.AddItem Format(Now, "short time")
lstDateTime.AddItem Format(Now, "medium time")
lstDateTime.AddItem Format(Now, "medium date")
lstDateTime.AddItem Format(Now, "long date")
lstDateTime.AddItem Format(Now, "short date")
lstDateTime.AddItem Format(Date, "dd - mm - yyyy")
lstDateTime.AddItem Format(Date, "dd/mm/yy")
lstDateTime.AddItem Format(Date, "dd/mm/yyyy")
lstDateTime.AddItem Format(Date, "dd/mm")
lstDateTime.AddItem Format(Date, "dd")
lstDateTime.AddItem Format(Time, "hh-mm-ss")
lstDateTime.AddItem Format(Time, "hh.mm.ss")
lstDateTime.AddItem Format(Time, "hh-mm")
End Sub
Private Sub lstDateTime_DblClick()
    cmdOK_Click
End Sub
