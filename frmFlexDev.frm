VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexDev 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluciones ..."
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   10695
      TabIndex        =   14
      Top             =   1275
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "&Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexDev.frx":0000
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   315
      Left            =   9435
      TabIndex        =   13
      Top             =   1275
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "&Consultar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexDev.frx":001C
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   300
      Left            =   5385
      Top             =   1275
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   529
      Caption         =   "-F4- Consultar -F5- Ir a Rejilla  -F8- Salir"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
      CaptionAlignment=   1
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   1545
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   2725
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmFlexDev.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbTEMPOR"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ioREF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ioMODELO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioCODBAR"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexDev.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbCAJAS"
      Tab(1).Control(1)=   "ioIMPORTE"
      Tab(1).Control(2)=   "ioMOTIVO"
      Tab(1).Control(3)=   "ioFECHAI"
      Tab(1).Control(4)=   "ioFECHAF"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(9)=   "Label6"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Hoja 3"
      TabPicture(2)   =   "frmFlexDev.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label21"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chameleonButton1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ioCODVAL"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cbESTADO"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cbImprimir"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "ioVER"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   -74220
         TabIndex        =   16
         Top             =   30
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo cbCODCOL 
         Height          =   465
         Left            =   -70905
         TabIndex        =   17
         Top             =   30
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo miCombo2 
         Height          =   495
         Left            =   -66480
         TabIndex        =   18
         Top             =   30
         Width           =   3045
         _ExtentX        =   5821
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo cbTIPOAB 
         Height          =   495
         Left            =   -74220
         TabIndex        =   19
         Top             =   480
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miText ioFECHAINI 
         Height          =   480
         Index           =   0
         Left            =   -68910
         TabIndex        =   20
         Top             =   495
         Width           =   1425
         _ExtentX        =   2381
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   480
         Index           =   0
         Left            =   -66465
         TabIndex        =   21
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -68910
         TabIndex        =   22
         Top             =   525
         Width           =   4155
         _ExtentX        =   5821
         _ExtentY        =   873
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
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   480
         Left            =   -73995
         TabIndex        =   23
         Top             =   45
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo cbSUBFAM 
         Height          =   480
         Left            =   -68910
         TabIndex        =   24
         Top             =   45
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miText ioCODBAR 
         Height          =   495
         Left            =   870
         TabIndex        =   0
         Top             =   105
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioMODELO 
         Height          =   510
         Left            =   6960
         TabIndex        =   2
         Top             =   105
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioREF 
         Height          =   495
         Left            =   3345
         TabIndex        =   1
         Top             =   105
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbTEMPOR 
         Height          =   480
         Left            =   3330
         TabIndex        =   3
         Top             =   615
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   495
         Left            =   -67815
         TabIndex        =   6
         Top             =   120
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miText ioIMPORTE 
         Height          =   495
         Left            =   -67800
         TabIndex        =   8
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioMOTIVO 
         Height          =   495
         Left            =   -73455
         TabIndex        =   7
         Top             =   645
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAI 
         Height          =   495
         Left            =   -73455
         TabIndex        =   4
         Top             =   105
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAF 
         Height          =   495
         Left            =   -70380
         TabIndex        =   5
         Top             =   120
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioVER 
         Height          =   480
         Left            =   -68355
         TabIndex        =   11
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton cbImprimir 
         Height          =   450
         Left            =   -66840
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   90
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   794
         BTYPE           =   9
         TX              =   "&Imprimir Vale"
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
         MICON           =   "frmFlexDev.frx":008C
         PICN            =   "frmFlexDev.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   495
         Left            =   -71820
         TabIndex        =   10
         Top             =   240
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miText ioCODVAL 
         Height          =   495
         Left            =   -74160
         TabIndex        =   9
         Top             =   225
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   450
         Left            =   -66840
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   615
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   794
         BTYPE           =   9
         TX              =   "&Imprimir Rejilla"
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
         MICON           =   "frmFlexDev.frx":0D82
         PICN            =   "frmFlexDev.frx":0D9E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALE"
         Height          =   330
         Left            =   -74835
         TabIndex        =   46
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ver"
         Height          =   330
         Left            =   -68865
         TabIndex        =   45
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -72600
         TabIndex        =   44
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -71835
         TabIndex        =   42
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPORADA"
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL"
         Height          =   330
         Left            =   -75090
         TabIndex        =   40
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   -68430
         TabIndex        =   39
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   -68760
         TabIndex        =   38
         Top             =   735
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MOTIVO"
         Height          =   285
         Left            =   -74355
         TabIndex        =   37
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODBAR"
         Height          =   285
         Left            =   15
         TabIndex        =   36
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         Height          =   285
         Left            =   6060
         TabIndex        =   35
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REF."
         Height          =   285
         Left            =   2820
         TabIndex        =   34
         Top             =   165
         Width           =   525
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.FINAL"
         Height          =   285
         Left            =   -67500
         TabIndex        =   33
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.INICIAL"
         Height          =   285
         Left            =   -69960
         TabIndex        =   32
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   31
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   285
         Left            =   -71610
         TabIndex        =   30
         Top             =   105
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   345
         Left            =   -67290
         TabIndex        =   29
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   300
         Left            =   -74670
         TabIndex        =   28
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   330
         Left            =   -70245
         TabIndex        =   27
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   315
         Left            =   -74985
         TabIndex        =   26
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAM."
         Height          =   315
         Left            =   -69855
         TabIndex        =   25
         Top             =   120
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5310
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   11460
      _cx             =   20214
      _cy             =   9366
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14331047
      ForeColor       =   -2147483640
      BackColorFixed  =   15120763
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14859077
      ForeColorSel    =   -2147483635
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15573900
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFlexDev.frx":1A78
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   3
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmFlexDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmFlexDev
' Fecha/Hora  : 04/02/2004 21:05
' Autor       : JCASTILLO
' Propósito   : Consultar las devoluciones que hay en la base de datos
'---------------------------------------------------------------------------------------
Option Explicit

Dim first As Boolean

'Dim tmpstrcombo As String

Public miosql As New clsSmartSQL
Public artsql As New clsSmartSQL

Public miRc As New ADODB.Recordset
Public seleccionado As Boolean

Public desde_mnt As Boolean

'---------------------------------------------------------------------------------
'para seleccionar desde la venta una devolución distinta a la ultima
'devolución.
'---------------------------------------------------------------------------------
Public desde_ventas As Boolean
Public ID_Devolucion As Long
Public Caja_Devolucion As Byte
'---------------------------------------------------------------------------------

Dim miCod As MiCodBar

'Dim nif As New clsNIF

Private Sub cbBorrar_click()

ioCODBAR.Text = ""
ioVER.Text = "50"
ioFECHAI.Text = Date
ioIMPORTE.Text = ""
cbCAJAS.Text = CajaActual
ioMOTIVO.Text = ""
cbESTADO.Text = ""
cbTEMPOR.Text = TemporadaActual

fg.Rows = 1
fg.Clear
'Call cbLista_click

End Sub

Private Sub cbCAJAS_Validate(Cancel As Boolean)
Call cbLista_click
End Sub

Private Sub cbImprimir_Click()

If fg.Rows <= 1 Then Exit Sub

'imprimir el vale de la devolucion seleccionada
If fg.TextMatrix(fg.Row, 1) <> "" And fg.TextMatrix(fg.Row, 14) <> "" Then
    
    'ver que no este anulada
    If fg.TextMatrix(fg.Row, 10) <> "ANULADA" Then
        Call Imprime_Vale(fg.TextMatrix(fg.Row, 1), fg.TextMatrix(fg.Row, 14), locCnn)
    Else
        MsgBox "La devolución esta ANULADA, no se puede imprimir un vale", vbExclamation, titulo
    End If

End If

End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean
Dim nuefech  As String
Dim tmpffin As Date

On Error GoTo cbLista_click_Error

miosql.ClearWhereClause '
                        '
                        ' masql.ClearWhereClause
                        '
artsql.ClearWhereClause '

'si solo tiene fecha inicial
If ((ioFECHAI.Text <> "") And IsDate(ioFECHAI.Text)) And ((ioFECHAF.Text = "")) Then

    'miosql.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'miosql.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'miOsql.AddComplexWhereClause "Year(FALTA IN (" & masql.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    nuefech = DateAdd("d", 1, ioFECHAI.Text)
    miosql.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAI.Text)), "0000") & Format(Month((ioFECHAI.Text)), "00") & Format(Day((ioFECHAI.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "'", LOGIC_AND
    usa_where = True

'si tiene fecha inicial y fecha final
ElseIf ((ioFECHAI.Text <> "") And IsDate(ioFECHAI.Text) And ((ioFECHAF.Text <> "") And IsDate(ioFECHAF.Text))) Then

    'añadir un dia mas (para decir < el dia mas para que saque todos
    'los registros, sino no lo hace bien, y los del ultimo dia no los saca.
        
    tmpffin = ioFECHAF.Text
    tmpffin = DateAdd("d", 1, tmpffin)
    
    miosql.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAI.Text)), "0000") & Format(Month((ioFECHAI.Text)), "00") & Format(Day((ioFECHAI.Text)), "00") & "' AND FMODI < '" & Format(Year((tmpffin)), "0000") & Format(Month((tmpffin)), "00") & Format(Day((tmpffin)), "00") & "'", LOGIC_AND
    usa_where = True

End If

'buscar por codbar
If ioCODBAR.Text <> "" Then
    miosql.AddSimpleWhereClause "CODART", miCod.CODIGO_ART, , , LOGIC_AND
    miosql.AddSimpleWhereClause "TEMPOR", miCod.TEMPORADA_ART, , , LOGIC_AND
    miosql.AddSimpleWhereClause "CODTALLA", miCod.TALLA_ART, , , LOGIC_AND
    miosql.AddSimpleWhereClause "CODCOL", miCod.COLOR_ART, , , LOGIC_AND
    usa_where = True
End If

'si hay codigo de vale, descomponer CODIGO de vale propiamente dicho y CAJA
If (ioCODVAL.Text <> "") Then
    miosql.AddSimpleWhereClause "CODVAL", Left(ioCODVAL.Text, Len(ioCODVAL.Text) - 3), , CLAUSE_LIKE, LOGIC_AND
    miosql.AddSimpleWhereClause "CODCAJA", Right(ioCODVAL.Text, 3), , , LOGIC_AND
    usa_where = True
End If

If (ioIMPORTE.Text <> "") Then
    If CDbl(ioIMPORTE.Text) > "0" Then
            miosql.AddSimpleWhereClause "IMPORTE", CDbl(ioIMPORTE.Text), , , LOGIC_AND
            usa_where = True
    End If
End If

If ioMOTIVO.Text <> "" Then
    miosql.AddSimpleWhereClause "MOTIVO", ioMOTIVO.Text, , CLAUSE_LIKE
    usa_where = True
End If

'si hay un vale, no evaluar el valor del combo caja
If (ioCODVAL.Text = "") Then

    If cbCAJAS.Text <> "" Then
        miosql.AddSimpleWhereClause "CODCAJA", cbCAJAS.Text, , , LOGIC_AND
        usa_where = True
    End If

End If


If (ioMODELO.Text <> "") And (cbTEMPOR.Text = "") Then
    MsgBox "Para buscar modelo debe seleccionar también una temporada", vbInformation, titulo
    usa_where = False
    Exit Sub
ElseIf (ioMODELO.Text <> "") And (cbTEMPOR.Text <> "") Then
    artsql.AddSimpleWhereClause "MODELO", ioMODELO.Text, , CLAUSE_LIKE, LOGIC_AND
    artsql.AddSimpleWhereClause "TEMPOR", cbTEMPOR.Text, , , LOGIC_AND
    usa_where = True
End If

If ioREF.Text <> "" Then
    artsql.AddSimpleWhereClause "REF", ioREF.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    miosql.AddSimpleWhereClause "MBAJA", cbESTADO.Text, , , LOGIC_AND
    usa_where = True
End If


Debug.Print miosql.SQL

'si deja todo en blanco, no mostrar ningun registro
If Not usa_where Then
    fg.Clear
    Exit Sub
End If

If ioMODELO.Text <> "" Then miosql.AddComplexWhereClause "CODART IN (" & artsql.SQL & ") and TEMPOR = " & cbTEMPOR.Text, LOGIC_AND

If miRc.State = 1 Then miRc.Close
miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic

fg.Rows = 1
'Set fg.DataSource = miRc

Call carga_grid

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy

'fg.ColHidden(fg.Cols - 1) = True

DoEvents

    With fg
    
    .ColFormat(0) = "000000000"
    .AutoSize 0, .Cols - 1
    
    DoEvents

 End With


   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexDev"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid
' Fecha/Hora    : 26/01/2004 09:59
' Autor         : JCastillo
' Propósito     :   Cargar el grid con las devoluciones
'---------------------------------------------------------------------------------------
'
'id
'fecha
'MODELO
'TALLA
'COLOR
'Motivo
'PVP
'USUARIO
'CAJA
'ESTADO:
'           0 = activa
'           1 = cancelada
'---------------------------------------------------------------------------------------
Private Sub carga_grid()
Dim tmpcodcolor As Long
Dim conta_filas As Long
Dim tart As Variant

   On Error GoTo carga_grid_Error

   With fg
   
    .Clear
    .Cols = 16
    .ColFormat(9) = "Currency"
    .ColAlignment(3) = flexAlignLeftCenter
    .ColHidden(0) = True
    .ColHidden(1) = True
    .ColHidden(14) = True
    .ColHidden(15) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Fecha"
    .TextMatrix(0, 3) = "Ref"
    .TextMatrix(0, 4) = "Modelo"
    .TextMatrix(0, 5) = "Prov."
    
    .TextMatrix(0, 6) = "Talla"
    .TextMatrix(0, 7) = "Color"
    .TextMatrix(0, 8) = "Motivo"
    .TextMatrix(0, 9) = "PVP"
    .TextMatrix(0, 10) = "Usuario"
    .TextMatrix(0, 11) = "Caja"
    .TextMatrix(0, 12) = "Estado"
    .TextMatrix(0, 13) = "C.Vale"
    
    
    '0 = activa
    '1 = cancelada
    
    
    If miRc.RecordCount <= 0 Then Exit Sub
        
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            conta_filas = conta_filas + 1
            
            tart = devuelve_matriz("SELECT MODELO, REF, CODPROV FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR"), locCnn)
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_filas
            'CODIGO DE VALE
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODVAL")
            
            'FECHA
            .TextMatrix(.Rows - 1, 2) = miRc.fields("FMODI")
            
            'REFERENCIA
            .TextMatrix(.Rows - 1, 3) = Trim(tart(1))
            
            'MODELO
            .TextMatrix(.Rows - 1, 4) = Format(miRc.fields("CODART"), "00000") & "-" & Trim(tart(0))
            
            'PROV
            .TextMatrix(.Rows - 1, 5) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tart(2), locCnn))
            
            'TALLA
            .TextMatrix(.Rows - 1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & miRc.fields("CODTALLA"), locCnn))
            
            'obtener el texto del color y su codigo de color (para colorear
            'la celda del grid)
            'COLOR
            If miRc.fields("CODCOL") > 0 Then
      
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL"), locCnn)
                .TextMatrix(.Rows - 1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL")))
                .Col = 7
                .Row = .Rows - 1
                .CellBackColor = tmpcodcolor
                .Col = 2
        
            End If
            
            'MOTIVO
            .TextMatrix(.Rows - 1, 8) = Trim(miRc.fields("MOTIVO"))
            
            'COSTURERA
           ' If Not IsNull(miRc.Fields("CODCOST")) Then .TextMatrix(.Rows - 1, 7) = devuelve_campo("SELECT NOMBRE FROM COSTURE WHERE CODIGO = " & miRc.Fields("CODCOST"), locCnn)
            
            'precio de venta
            .TextMatrix(.Rows - 1, 9) = miRc.fields("IMPORTE")
            
            'usuario
            .TextMatrix(.Rows - 1, 10) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & miRc.fields("CODUSR"), locCnn))
            
             'caja
            .TextMatrix(.Rows - 1, 11) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & miRc.fields("CODCAJA"), locCnn))
            
            '0=ACTIVA, 1=ANULADA
            Select Case miRc.fields("MBAJA")
            
            Case False
                      .TextMatrix(.Rows - 1, 12) = ""
            Case True
                      .TextMatrix(.Rows - 1, 12) = "ANULADA"
            'Case 3
                      '.TextMatrix(.Rows - 1, 10) = "CANCELADO"
                      
            End Select
                        
            'estado
            '.TextMatrix(.Rows - 1, 10) = miRc.Fields("ESTADO")
            
            .TextMatrix(.Rows - 1, 13) = miRc.fields("CODVAL") & Format(miRc.fields("CODCAJA"), "000")
            
            .TextMatrix(.Rows - 1, 14) = miRc.fields("CODCAJA")
            
            .TextMatrix(.Rows - 1, 15) = miRc.fields("CODIGO")
     
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        
        .subtotal flexSTCount, -1, 4, , vbBlue, vbWhite
        .subtotal flexSTSum, -1, 9, , vbBlue, vbWhite
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 4) = "Nº Devoluciones (" & .TextMatrix(1, 4) & ")"
        '.TextMatrix(1, 4) = ""
        .TextMatrix(1, 8) = "Total:"
        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
  
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexDev"
 
 End Sub

Private Sub cbTEMPOR_Validate(Cancel As Boolean)

If ioMODELO.Text <> "" And cbTEMPOR.Text <> "" Then Call cbLista_click

End Sub

Private Sub chameleonButton1_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpcaja As String
         

   On Error GoTo chameleonButton1_Click_Error
   
    DoEvents

    If cbCAJAS.Text <> "" Then
        tmpcaja = devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & cbCAJAS.Text, locCnn)
        If tmpcaja = "@" Then tmpcaja = ""
    End If
    
    linea1 = "Informe de Devoluciones. Motivo: " & ioMOTIVO.Text & ". F.Inicial: " & ioFECHAI.Text & ". F.Final: " & ioFECHAF.Text & ". Caja: " & tmpcaja
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    'fg.Font.Size = 9
    'fg.Refresh
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.Font.Size = 11
   ' fg.Refresh
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexDev"

End Sub

Private Sub fg_dblClick()
    seleccionado = True
    
    If fg.Rows <= 1 Then Exit Sub
        
    'si son datos validos, cargar ID devolucion y Caja
    If Trim(fg.TextMatrix(fg.Row, 14)) <> "" And Trim(fg.TextMatrix(fg.Row, 15)) <> "" Then
        If IsNumeric(fg.TextMatrix(fg.Row, 14)) And IsNumeric(fg.TextMatrix(fg.Row, 15)) Then
            
            ID_Devolucion = fg.TextMatrix(fg.Row, 15)
            Caja_Devolucion = fg.TextMatrix(fg.Row, 14)
            
        Else
            ID_Devolucion = 0
            Caja_Devolucion = 0
        
        End If
    
    Else
        ID_Devolucion = 0
        Caja_Devolucion = 0
    End If
    
    
    If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        'posicionarse en el registro
        miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
        DoEvents
        
        If desde_mnt Or desde_ventas Then
        
            Unload Me
        
        Else
        
          'If Trim(fg.TextMatrix(fg.Row, 14)) <> "" And Trim(fg.TextMatrix(fg.Row, 15)) <> "" Then
           '   If IsNumeric(fg.TextMatrix(fg.Row, 14)) And IsNumeric(fg.TextMatrix(fg.Row, 15)) Then
            
            If ID_Devolucion > 0 And Caja_Devolucion > 0 Then
            
                With frmMntDev
                    .codigo_devol = ID_Devolucion
                    .Caja_Devol = Caja_Devolucion
                    .Show
                End With
            
            End If
                   
            Unload Me
                    
        End If
                
    End If

End Sub
       
   ' Unload Me

'Exit Sub

'si corresponde a algun ID
'If fg.TextMatrix(fg.Row, 0) <> "" Then
'
'With frmNuArr''

'    .Solo_Actualizar = True
'    .Sel_Caja = fg.TextMatrix(fg.Row, 1)
'    .Sel_ID = fg.TextMatrix(fg.Row, 0)
'
'    .ioNOMBRE.Text = fg.TextMatrix(fg.Row, 7)
'    .ioDESCRIPCION.Text = fg.TextMatrix(fg.Row, 6)
'    .ioPVP.Text = fg.TextMatrix(fg.Row, 8)
'
 '   Select Case fg.TextMatrix(fg.Row, 11)
 '
 '           Case "PENDIENTE"
 ''
 '                   .cbESTADO.Text = 1
 '           Case "SERVIDO"
 '
  '                  .cbESTADO.Text = 2
'            Case "CANCELADO"
'
 '                   .cbESTADO.Text = 3
 '
'    End Select
'
'    '.cbESTADO.Text = fg.TextMatrix(fg.Row, 6)
'
'    .Show 1
'    Call cbLista_click
'
'End With
    
'End If
    
   ' Unload Me



Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13
    
    Call fg_dblClick
    seleccionado = True
    KeyAscii = 0
   ' Unload Me
    
End Select

End Sub

Private Sub fg_LostFocus()

fg.TabStop = False

End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    DoEvents
    
    If Not first Then
    
       ' Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortShow
          

        first = True
    End If
    
    fg.Visible = True
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioCODBAR.SetFocus
        Else
            fg.TabStop = True
            fg.Select 1, 1, 1, fg.Cols - 1
            fg.SetFocus
        End If
    End If
    KeyCode = 0

'salir del formulario actual
Case vbKeyF8

    KeyCode = 0
    Unload Me

End Select

End Sub

Private Sub Form_Load()

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

  fg.Visible = False
  fg.Rows = 1
  fg.Cols = 0
  
  'Cargar el micombo cajas
  With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    DoEvents
    .Text = CajaActual
    .Refresh
  End With
  
  With cbTEMPOR
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .DataField = "TEMPOR"
    .carga
    DoEvents
    .Text = TemporadaActual
   ' Set .DataSource = rc
  End With
  
  With ioFECHAI
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
   .Text = Date
  End With
  
  With ioFECHAF
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
  With ioIMPORTE
    .dspFormat = "Currency"
   .LongMaxima = 10
   .Alineacion = 1
  End With
  
 With ioMODELO
   .LongMaxima = 30
 End With
 
 With ioCODVAL
    .LongMaxima = 15
    .SoloNumeros = True
    .Alineacion = 1
 End With
  
   Select Case TipoPermiso
   
   'usuario comun, ver solo los pedidos de su almacén
   Case 0
        cbCAJAS.Enabled = False
   'supervisor, ver todos los pedidos
  ' Case 1
   
   End Select
   
With cbESTADO
    .LenCodigo = 1
    .añade_item "0  ACTIVA"
    .añade_item "1  ANULADA"
    .CodigoWidth = 300
End With
  
 artsql.AddTable "MAARTIC"
 artsql.AddField "CODIGO"
' masql.AddTable "COSTURE"
 miosql.AddTable "DEVOL"
 'masql.AddField "CODIGO"
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    tmpstrcombo = ""
    'Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    
    If (desde_mnt = False) And (desde_ventas = False) Then Set frmFlexDev = Nothing
    
End Sub



Private Sub ioCODBAR_GotFocus()

If Tab1.Tab <> 0 Then Tab1.Tab = 0

End Sub

Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim t_articulo As Variant
 
With ioCODBAR

If Trim(.Text) = "" Then Exit Sub

    'si es un codigo de barras con la longitud válidad
    If Len(Trim(.Text)) = LenCodBar Then
        
        
         'descomponer el codigo
        miCod = Descompone_CBAR(.Text)
        
        'comprobar si existe el artículo/temporada
        t_articulo = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & miCod.CODIGO_ART & " AND TEMPOR = " & miCod.TEMPORADA_ART, locCnn)
        
        If t_articulo = "@" Then
        
                MsgBox "No existe el artículo para esa temporada!, Codigo de Barras no Válido", titulo, vbExclamation
                ioCODBAR.Text = ""
                .CancelarValidacion
                Cancel = True
                Exit Sub
                
        End If
    
    Else
            MsgBox "Código de barras no válido", vbInformation, titulo
    
    End If
        
End With


End Sub



Private Sub ioCODVAL_GotFocus()

If Tab1.Tab <> 2 Then Tab1.Tab = 2

End Sub

Private Sub ioCODVAL_Validate(Cancel As Boolean)

   On Error GoTo ioCODVAL_Validate_Error

    If ioCODVAL.Text <> "" Then
        ioFECHAI.Text = ""
        ioFECHAF.Text = ""
        Call cbLista_click
    End If

   On Error GoTo 0
   Exit Sub

ioCODVAL_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODVAL_Validate de Formulario frmFlexDev"

End Sub

Private Sub ioFECHAI_GotFocus()

If Tab1.Tab <> 1 Then Tab1.Tab = 1

End Sub

Private Sub ioFECHAi_Validate(Cancel As Boolean)

If ioFECHAI.Text <> "" Then Call cbLista_click

End Sub


Private Sub ioFECHAf_Validate(Cancel As Boolean)

If (ioFECHAF.Text <> "") And (ioFECHAI.Text <> "") Then Call cbLista_click

End Sub



Private Sub ioVER_LostFocus()
ioCODBAR.SetFocus
End Sub
