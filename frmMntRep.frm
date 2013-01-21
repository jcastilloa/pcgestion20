VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMntRep 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Representantes"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9555
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9555
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   2715
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   3870
      Left            =   45
      TabIndex        =   37
      Top             =   495
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   6826
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmMntRep.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioTELEFONO2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioEMAIL"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ioTELEFONO1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ioPAIS"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ioFAX"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ioPROVINCIA"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ioCODPOS"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ioCIF"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ioLOCALIDAD"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ioDIRECCION"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ioNOMBRE"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmMntRep.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmBorrarFoto"
      Tab(1).Control(1)=   "ioCOMISION"
      Tab(1).Control(2)=   "ioENTIDAD"
      Tab(1).Control(3)=   "ioSUCURSAL"
      Tab(1).Control(4)=   "ioDC"
      Tab(1).Control(5)=   "ioCUENTA"
      Tab(1).Control(6)=   "ioCODBAN"
      Tab(1).Control(7)=   "ioFOTO"
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(10)=   "ioAPLICARCOM"
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(13)=   "Label19"
      Tab(1).Control(14)=   "Label18"
      Tab(1).Control(15)=   "BANCO"
      Tab(1).ControlCount=   16
      Begin PCGestion.ucGrdBttn cmBorrarFoto 
         Height          =   315
         Left            =   -65910
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3120
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   556
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "frmMntRep.frx":0038
      End
      Begin PCGestion.miText ioNOMBRE 
         Height          =   525
         Left            =   1320
         TabIndex        =   0
         Top             =   720
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioDIRECCION 
         Height          =   525
         Left            =   1320
         TabIndex        =   2
         Top             =   1260
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioLOCALIDAD 
         Height          =   525
         Left            =   1335
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioCIF 
         Height          =   525
         Left            =   7830
         TabIndex        =   1
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioCODPOS 
         Height          =   525
         Left            =   8280
         TabIndex        =   3
         Top             =   1260
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioPROVINCIA 
         Height          =   525
         Left            =   5910
         TabIndex        =   5
         Top             =   1815
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioFAX 
         Height          =   525
         Left            =   1320
         TabIndex        =   9
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioPAIS 
         Height          =   525
         Left            =   1320
         TabIndex        =   6
         Top             =   2355
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioTELEFONO1 
         Height          =   525
         Left            =   4575
         TabIndex        =   7
         Top             =   2355
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioEMAIL 
         Height          =   525
         Left            =   3930
         TabIndex        =   10
         Top             =   2880
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioTELEFONO2 
         Height          =   525
         Left            =   7530
         TabIndex        =   8
         Top             =   2355
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioCOMISION 
         Height          =   525
         Left            =   -73665
         TabIndex        =   11
         Top             =   675
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioENTIDAD 
         Height          =   525
         Left            =   -73665
         TabIndex        =   13
         Top             =   1665
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioSUCURSAL 
         Height          =   525
         Left            =   -72030
         TabIndex        =   14
         Top             =   1665
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioDC 
         Height          =   525
         Left            =   -71055
         TabIndex        =   15
         Top             =   1665
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miText ioCUENTA 
         Height          =   525
         Left            =   -70140
         TabIndex        =   16
         Top             =   1665
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin PCGestion.miCombo ioCODBAN 
         Height          =   495
         Left            =   -73665
         TabIndex        =   12
         Top             =   1185
         Width           =   5100
         _ExtentX        =   8996
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
      Begin VB.PictureBox ioFOTO 
         Height          =   2790
         Left            =   -68535
         ScaleHeight     =   2730
         ScaleWidth      =   2835
         TabIndex        =   58
         Top             =   660
         Width           =   2895
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COMISION"
         Height          =   360
         Left            =   -74850
         TabIndex        =   57
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   360
         Left            =   -72975
         TabIndex        =   56
         Top             =   750
         Width           =   210
      End
      Begin MSForms.CheckBox ioAPLICARCOM 
         Height          =   435
         Left            =   -72630
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   705
         Width           =   2385
         VariousPropertyBits=   746596371
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4207;767"
         Value           =   "0"
         Caption         =   "APLICAR COMISION"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CTA"
         Height          =   330
         Left            =   -70620
         TabIndex        =   54
         Top             =   1755
         Width           =   450
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC"
         Height          =   300
         Left            =   -71325
         TabIndex        =   53
         Top             =   1755
         Width           =   270
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ENTIDAD"
         Height          =   330
         Left            =   -74640
         TabIndex        =   52
         Top             =   1755
         Width           =   915
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFICINA"
         Height          =   300
         Left            =   -72840
         TabIndex        =   51
         Top             =   1755
         Width           =   810
      End
      Begin VB.Label BANCO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO"
         Height          =   330
         Left            =   -74865
         TabIndex        =   50
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAX"
         Height          =   360
         Left            =   675
         TabIndex        =   48
         Top             =   2955
         Width           =   525
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         Height          =   360
         Left            =   3105
         TabIndex        =   47
         Top             =   2955
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEF. 2"
         Height          =   360
         Left            =   6480
         TabIndex        =   46
         Top             =   2430
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEF. 1"
         Height          =   360
         Left            =   3435
         TabIndex        =   45
         Top             =   2430
         Width           =   1080
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAIS"
         Height          =   360
         Left            =   645
         TabIndex        =   44
         Top             =   2430
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE Y APELLIDOS"
         Height          =   600
         Left            =   120
         TabIndex        =   43
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION"
         Height          =   360
         Left            =   135
         TabIndex        =   42
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CP"
         Height          =   360
         Left            =   7770
         TabIndex        =   41
         Top             =   1335
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOCALIDAD"
         Height          =   360
         Left            =   60
         TabIndex        =   40
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA"
         Height          =   360
         Left            =   4680
         TabIndex        =   39
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NIF"
         Height          =   360
         Left            =   7365
         TabIndex        =   38
         Top             =   795
         Width           =   375
      End
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5250
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmMntRep.frx":0054
      PICN            =   "frmMntRep.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   30
      Top             =   4815
      Width           =   9495
      _ExtentX        =   16536
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   30
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5250
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmMntRep.frx":0D42
      PICN            =   "frmMntRep.frx":0D5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   4425
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5235
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "Lista F4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmMntRep.frx":1A94
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdNext 
      Height          =   630
      Left            =   7410
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5235
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmMntRep.frx":1AB0
      PICN            =   "frmMntRep.frx":1ACC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdLast 
      Height          =   630
      Left            =   8475
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5235
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmMntRep.frx":279E
      PICN            =   "frmMntRep.frx":27BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   15
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5910
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Agregar F1"
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
      MICON           =   "frmMntRep.frx":34F0
      PICN            =   "frmMntRep.frx":350C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbActualizar 
      Height          =   795
      Left            =   1110
      TabIndex        =   17
      Top             =   5910
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Actualizar F2"
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
      MICON           =   "frmMntRep.frx":41E6
      PICN            =   "frmMntRep.frx":4202
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEdicion 
      Height          =   795
      Left            =   2325
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5910
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Edicion F3"
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
      MICON           =   "frmMntRep.frx":4ADC
      PICN            =   "frmMntRep.frx":4AF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6390
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5895
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Cancelar"
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
      MICON           =   "frmMntRep.frx":5356
      PICN            =   "frmMntRep.frx":5372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEliminar 
      Height          =   795
      Left            =   7380
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5895
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "E&liminar F9"
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
      MICON           =   "frmMntRep.frx":5C4C
      PICN            =   "frmMntRep.frx":5C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCerrar 
      Height          =   795
      Left            =   8475
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5895
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "Cerrar ESC"
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
      MICON           =   "frmMntRep.frx":683A
      PICN            =   "frmMntRep.frx":6856
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   780
      TabIndex        =   36
      Top             =   4380
      Width           =   840
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1482;767"
      Value           =   "0"
      Caption         =   "Baja"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label ioFALTA 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5025
      TabIndex        =   25
      Top             =   4410
      Width           =   1290
   End
   Begin VB.Label ioFBAJA 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8190
      TabIndex        =   24
      Top             =   4410
      Width           =   1290
   End
   Begin VB.Label ioFMODI 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7095
      TabIndex        =   23
      Top             =   75
      Width           =   2445
   End
   Begin VB.Label ioCODIGO 
      Alignment       =   2  'Center
      BackColor       =   &H00AC998C&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1395
      TabIndex        =   22
      Top             =   60
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   435
      TabIndex        =   21
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   3885
      TabIndex        =   20
      Top             =   4440
      Width           =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   4815
      TabIndex        =   19
      Top             =   105
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   6975
      TabIndex        =   18
      Top             =   4425
      Width           =   1155
   End
End
Attribute VB_Name = "frmMntRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim prime As Boolean

Dim oSQL As New clsSmartSQL
Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub cmBorrarFoto_Click()

If mbEditFlag = False And mbAddNewFlag = False Then Exit Sub

If MsgBox("¿Desea quitar la imagen?", vbQuestion + vbYesNo, titulo) = vbYes Then

    rc.fields(ioFOTO.DataField).Value = Null
    ioFOTO.Picture = Nothing

End If

End Sub

Private Sub Form_Activate()

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Representantes. ¿Crear?", vbYesNo + vbQuestion, "Representantes") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
         
        Call cmdFirst_Click
        Call cbCancelar_Click
        
  End If
  
prime = True
End If
    
End Sub

Private Sub Form_Load()
  
  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
     With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "REPRESEN"
  oSQL.AddOrderClause "CODIGO"
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioNOMBRE
  Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "NOMBRE"
        .LongMaxima = 40
  End With
  
 With ioCIF
  Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "NIF"
        .LongMaxima = 15
  End With
  
  With ioDIRECCION
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DIRECCION"
        .LongMaxima = 40
  End With
  
  With ioCODPOS
  Set .DataSource = rc
       .PermitirBlanco = True
        .DataField = "CODPOS"
        .LongMaxima = 5
  End With

  With ioLOCALIDAD
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "LOCALIDAD"
        .LongMaxima = 40
  End With
  
  With ioPROVINCIA
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PROVINCIA"
        .LongMaxima = 40
  End With
  
  With ioPAIS
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PAIS"
        .LongMaxima = 25
  End With

 With ioTELEFONO1
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELEFONO1"
        .LongMaxima = 17
  End With
  
   With ioTELEFONO2
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELEFONO2"
        .LongMaxima = 17
  End With
  
  With ioFAX
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "FAX"
        .LongMaxima = 17
  End With
  
  With ioEMAIL
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "EMAIL"
        .LongMaxima = 50
  End With
  
  With ioCOMISION
  Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "COMISION"
        .LongMaxima = 2
  End With
  
  With ioAPLICARCOM
  Set .DataSource = rc
'        .PermitirBlanco = True
        .DataField = "APLICARCOM"
  End With
  
  
  With ioENTIDAD
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "ENTIDAD"
        .LongMaxima = 4
  End With
  
  With ioSUCURSAL
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "SUCURSAL"
        .LongMaxima = 4
  End With
  
  With ioDC
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DC"
        .LongMaxima = 2
  End With
  
  With ioCUENTA
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CUENTA"
        .LongMaxima = 10
  End With
  
  With ioCODBAN
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM BANCOS ORDER BY CODIGO"
    .LenCodigo = 2
    .DataField = "CODBAN"
    .carga
    .CodigoWidth = 700
    Set .DataSource = rc

  End With
  
  With ioFOTO
    .DataField = "FOTO"
    Set .DataSource = rc
  End With
   
  With ioFBAJA
  Set .DataSource = rc
        .DataField = "FBAJA"
  End With
  
  With ioFALTA
  Set .DataSource = rc
        .DataField = "FALTA"
  End With
  
        With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
  
  With ioMBAJA
  Set .DataSource = rc
        .DataField = "MBAJA"
  End With
       
       
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
      
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
      
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
      Case vbKeyF3
            Call cbedicion_Click
        
      Case vbKeyF4
            Call cbLista_click
      
      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
       Case vbKeyF8
        Call cmdLast_Click
      
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate


rc.Close
Set rc = Nothing

  ' With locCnn
 '   If .State <> 0 Then .Close
 '  End With
   
Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmMntRep = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

Dim tmprc As New ADODB.Recordset
Dim tmpstrcombo As String
  
tmprc.Open "SELECT CODIGO, NOMBRE FROM BANCOS ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly

With frmFlexSimple
        
    With .fg
            Set .DataSource = rc
             tmpstrcombo = .BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
            .ColComboList(16) = tmpstrcombo
            .ColFormat(1) = "00000"
            .AutoSize 1, .Cols - 1
    End With
    
    .Caption = "Representantes ..."
    .Show 1
   ' .SetFocus
End With

tmprc.Close
Set tmprc = Nothing




Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub





Private Sub ioCOMISION_GotFocus()
If Tab1.Tab = 0 Then Tab1.Tab = 1
End Sub

Private Sub ioCUENTA_Validate(Cancel As Boolean)
Tab1.Tab = 0
End Sub

Private Sub ioEMAIL_gotFocus()

If Tab1.Tab = 1 Then Tab1.Tab = 0

End Sub

Private Sub ioFOTO_Click()

If mbEditFlag = False And mbAddNewFlag = False Then Exit Sub

With Dialogo
.ShowOpen

If (.CancelError = True) Or (.filename = "") Then Exit Sub

'cargar imagen
ioFOTO.Picture = LoadPicture(Dialogo.filename)
GuardarBinary rc.fields(ioFOTO.DataField), ioFOTO

End With


End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then _
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
End Sub

Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  
    With ioNOMBRE
        If .Text = "" Then
            lblstatus.Caption = "NOMBRE no puede estar en blanco"
            .SetFocus
            bCancel = True
        End If
    End With
    
    With ioCIF
        If .Text = "" Then
            lblstatus.Caption = "CIF no puede estar en blanco"
            .SetFocus
            bCancel = True
        End If
    End With
       
    With ioCOMISION
        If .Text = "" Then
            lblstatus.Caption = "COMISION no puede estar en blanco"
            .SetFocus
            bCancel = True
        End If
    End With
    
    'With ioAPLICARCOM
     '  ' If .Text = "" Then
      '      lblstatus.Caption = "APLICAR COMISIÓN no puede estar en blanco"
       '     ioAPLICARCOM.SetFocus
        '    bCancel = True
      '  End If
    'End With
  
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cbAgregar_Click()
  Dim tmpcodigo As Variant
  
  On Error GoTo AddErr
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from REPRESEN")
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons (False)
    cbActualizar.Visible = True
    
    Tab1.SetFocus
    ioNOMBRE.SetFocus
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
    On Error GoTo DeleteErr
  With rc
    '.Delete
    '.MoveNext
    .fields("mbaja") = True
    .fields("FBAJA") = Date
    If .EOF Then .MoveLast
  End With
 
  Call cbactualizar_Click
   
  
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  Tab1.SetFocus
  ioNOMBRE.SetFocus
  
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cbCancelar_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rc.CancelUpdate
  If mvBookMark > 0 Then
    rc.Bookmark = mvBookMark
  Else
    rc.MoveFirst
  End If
  mbDataChanged = False
  Tab1.Tab = 0
  

End Sub

Private Sub cbactualizar_Click()
  On Error GoTo UpdateErr
 
  
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  Tab1.Tab = 0
  
  cbAgregar.SetFocus

  Exit Sub
UpdateErr:
  If Err.Number = -2147217887 Then Exit Sub
  MsgBox Err.Description, vbInformation, "Atención"
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rc.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rc.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rc.EOF Then rc.MoveNext
  If rc.EOF And rc.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    rc.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rc.BOF Then rc.MovePrevious
  If rc.BOF And rc.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    rc.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
  cbLista.Visible = bVal
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  ioCODBAN.Locked = bVal
  
End Sub
