VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMenuTactil 
   BackColor       =   &H00AFAFAF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Táctil"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   630
   Visible         =   0   'False
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   1080
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Mantenimientos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":0000
      PICN            =   "frmMenuTactil.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton chameleonButton1 
      Height          =   1080
      Left            =   45
      TabIndex        =   1
      Top             =   1155
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Compras"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":0CF6
      PICN            =   "frmMenuTactil.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton chameleonButton2 
      Height          =   1080
      Left            =   45
      TabIndex        =   2
      Top             =   2265
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Transferencias"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":19EC
      PICN            =   "frmMenuTactil.frx":1A08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton chameleonButton3 
      Height          =   1080
      Left            =   45
      TabIndex        =   3
      Top             =   3360
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Ventas"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":26E2
      PICN            =   "frmMenuTactil.frx":26FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton chameleonButton4 
      Height          =   1080
      Left            =   45
      TabIndex        =   4
      Top             =   4470
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Configuración"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":33D8
      PICN            =   "frmMenuTactil.frx":33F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   6615
      Left            =   1740
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Mantenimientos"
      TabPicture(0)   =   "frmMenuTactil.frx":40CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chameleonButton9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chameleonButton8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chameleonButton7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "frmMenuTactil.frx":40EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Transferencias"
      TabPicture(2)   =   "frmMenuTactil.frx":4106
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Ventas"
      TabPicture(3)   =   "frmMenuTactil.frx":4122
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Configuración"
      TabPicture(4)   =   "frmMenuTactil.frx":413E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin PCGestion.chameleonButton chameleonButton7 
         Height          =   1080
         Left            =   330
         TabIndex        =   7
         Top             =   900
         Width           =   1665
         _ExtentX        =   1640
         _ExtentY        =   1402
         BTYPE           =   3
         TX              =   "&Artículos"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   11513775
         BCOLO           =   11513775
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenuTactil.frx":415A
         PICN            =   "frmMenuTactil.frx":4176
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton chameleonButton8 
         Height          =   1080
         Left            =   330
         TabIndex        =   8
         Top             =   2010
         Width           =   1665
         _ExtentX        =   1640
         _ExtentY        =   1402
         BTYPE           =   3
         TX              =   "&Proveedores"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   11513775
         BCOLO           =   11513775
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenuTactil.frx":4E50
         PICN            =   "frmMenuTactil.frx":4E6C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton chameleonButton9 
         Height          =   1080
         Left            =   330
         TabIndex        =   9
         Top             =   3120
         Width           =   1665
         _ExtentX        =   1640
         _ExtentY        =   1402
         BTYPE           =   3
         TX              =   "&Clientes"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   11513775
         BCOLO           =   11513775
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMenuTactil.frx":5B46
         PICN            =   "frmMenuTactil.frx":5B62
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin PCGestion.chameleonButton chameleonButton5 
      Height          =   1080
      Left            =   45
      TabIndex        =   6
      Top             =   5580
      Width           =   1665
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   3
      TX              =   "&Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMenuTactil.frx":683C
      PICN            =   "frmMenuTactil.frx":6858
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmMenuTactil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chameleonButton5_Click()

Call cerrar_conexiones

End Sub

Private Sub Form_Load()

'Unload Me

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

End Sub


