VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMntPer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personal"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   5115
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   3450
      Left            =   15
      TabIndex        =   41
      Top             =   495
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   6085
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmMntDep.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioEMAIL"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioFAX"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ioTELEFONO2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ioTELEFONO1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ioPAIS"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ioPROVINCIA"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ioPOBLACION"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "ioCODPOS"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ioDIRECCION"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ioCIF"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ioNOMBRE"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmMntDep.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmBorrarFoto"
      Tab(1).Control(1)=   "cbBancos"
      Tab(1).Control(2)=   "ioAFILIA"
      Tab(1).Control(3)=   "ioENTIDAD"
      Tab(1).Control(4)=   "ioSUCURSAL"
      Tab(1).Control(5)=   "ioDC"
      Tab(1).Control(6)=   "ioCUENTA"
      Tab(1).Control(7)=   "cbFPAGO"
      Tab(1).Control(8)=   "cbCENTROS"
      Tab(1).Control(9)=   "cbCAJAS"
      Tab(1).Control(10)=   "cbCambioPWD"
      Tab(1).Control(11)=   "cbTIPPERM"
      Tab(1).Control(12)=   "ioFOTO"
      Tab(1).Control(13)=   "Label25"
      Tab(1).Control(14)=   "Label23"
      Tab(1).Control(15)=   "Label22"
      Tab(1).Control(16)=   "Label16"
      Tab(1).Control(17)=   "Label18"
      Tab(1).Control(18)=   "Label19"
      Tab(1).Control(19)=   "Label20"
      Tab(1).Control(20)=   "Label21"
      Tab(1).Control(21)=   "Label17"
      Tab(1).ControlCount=   22
      Begin PCGestion.ucGrdBttn cmBorrarFoto 
         Height          =   315
         Left            =   -65085
         TabIndex        =   60
         Top             =   3015
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
         Image           =   "frmMntDep.frx":0038
      End
      Begin PCGestion.miText ioNOMBRE 
         Height          =   525
         Left            =   1350
         TabIndex        =   0
         Top             =   465
         Width           =   6615
         _ExtentX        =   11668
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
         Left            =   8640
         TabIndex        =   1
         Top             =   465
         Width           =   1545
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
      Begin PCGestion.miText ioDIRECCION 
         Height          =   525
         Left            =   1350
         TabIndex        =   2
         Top             =   1050
         Width           =   7050
         _ExtentX        =   12435
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
         Left            =   9420
         TabIndex        =   3
         Top             =   1050
         Width           =   765
         _ExtentX        =   2011
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
      Begin PCGestion.miText ioPOBLACION 
         Height          =   525
         Left            =   1350
         TabIndex        =   4
         Top             =   1605
         Width           =   3765
         _ExtentX        =   6641
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
         Left            =   6420
         TabIndex        =   5
         Top             =   1620
         Width           =   3795
         _ExtentX        =   6694
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
         Left            =   1350
         TabIndex        =   6
         Top             =   2160
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
         Left            =   4995
         TabIndex        =   7
         Top             =   2145
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
      Begin PCGestion.miText ioTELEFONO2 
         Height          =   525
         Left            =   8385
         TabIndex        =   8
         Top             =   2160
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
      Begin PCGestion.miText ioFAX 
         Height          =   525
         Left            =   1350
         TabIndex        =   9
         Top             =   2730
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
         Left            =   4110
         TabIndex        =   10
         Top             =   2730
         Width           =   6105
         _ExtentX        =   10769
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
      Begin PCGestion.miCombo cbBancos 
         Height          =   510
         Left            =   -73755
         TabIndex        =   15
         Top             =   1905
         Width           =   5925
         _ExtentX        =   10451
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
      End
      Begin PCGestion.miText ioAFILIA 
         Height          =   525
         Left            =   -73740
         TabIndex        =   13
         Top             =   1425
         Width           =   1680
         _ExtentX        =   4445
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
         Left            =   -73740
         TabIndex        =   16
         Top             =   2385
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
         Left            =   -71250
         TabIndex        =   17
         Top             =   2385
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
         Left            =   -73740
         TabIndex        =   18
         Top             =   2895
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         BackColor       =   15
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
         Left            =   -72165
         TabIndex        =   19
         Top             =   2880
         Width           =   1590
         _ExtentX        =   2884
         _ExtentY        =   926
         BackColor       =   16
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
      Begin PCGestion.miCombo cbFPAGO 
         Height          =   510
         Left            =   -71235
         TabIndex        =   14
         Top             =   1440
         Width           =   3405
         _ExtentX        =   6006
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
      End
      Begin PCGestion.miCombo cbCENTROS 
         Height          =   510
         Left            =   -73755
         TabIndex        =   11
         Top             =   480
         Width           =   5925
         _ExtentX        =   10451
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
      End
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   510
         Left            =   -73755
         TabIndex        =   12
         Top             =   945
         Width           =   5925
         _ExtentX        =   10451
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
      End
      Begin PCGestion.chameleonButton cbCambioPWD 
         Height          =   450
         Left            =   -70215
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2910
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "Cambiar Clave"
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
         MICON           =   "frmMntDep.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miCombo cbTIPPERM 
         Height          =   510
         Left            =   -70620
         TabIndex        =   20
         Top             =   2385
         Width           =   2775
         _ExtentX        =   4233
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
      End
      Begin VB.PictureBox ioFOTO 
         Height          =   2850
         Left            =   -67725
         ScaleHeight     =   2790
         ScaleWidth      =   2895
         TabIndex        =   64
         Top             =   495
         Width           =   2955
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   240
         Left            =   -74685
         TabIndex        =   63
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CENTRO"
         Height          =   240
         Left            =   -74730
         TabIndex        =   59
         Top             =   555
         Width           =   945
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.PAGO"
         Height          =   255
         Left            =   -72105
         TabIndex        =   61
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "AFILIACION"
         Height          =   360
         Left            =   -74925
         TabIndex        =   58
         Top             =   1485
         Width           =   1125
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFICINA"
         Height          =   300
         Left            =   -72180
         TabIndex        =   57
         Top             =   2490
         Width           =   810
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ENTIDAD"
         Height          =   330
         Left            =   -74700
         TabIndex        =   56
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC"
         Height          =   300
         Left            =   -74085
         TabIndex        =   55
         Top             =   3000
         Width           =   270
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA"
         Height          =   330
         Left            =   -73050
         TabIndex        =   54
         Top             =   2985
         Width           =   840
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO"
         Height          =   330
         Left            =   -74700
         TabIndex        =   53
         Top             =   2010
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "POBLACION"
         Height          =   300
         Left            =   105
         TabIndex        =   52
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA"
         Height          =   300
         Left            =   5175
         TabIndex        =   51
         Top             =   1710
         Width           =   1200
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAIS"
         Height          =   360
         Left            =   675
         TabIndex        =   50
         Top             =   2235
         Width           =   585
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEF. 1"
         Height          =   360
         Left            =   3855
         TabIndex        =   49
         Top             =   2235
         Width           =   1080
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEF. 2"
         Height          =   360
         Left            =   7350
         TabIndex        =   48
         Top             =   2235
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAX"
         Height          =   360
         Left            =   705
         TabIndex        =   47
         Top             =   2805
         Width           =   555
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         Height          =   360
         Left            =   3285
         TabIndex        =   46
         Top             =   2805
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE Y APELLIDOS"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   435
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NIF"
         Height          =   360
         Left            =   8205
         TabIndex        =   44
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION"
         Height          =   300
         Left            =   45
         TabIndex        =   43
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CP"
         Height          =   360
         Left            =   9000
         TabIndex        =   42
         Top             =   1140
         Width           =   375
      End
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4845
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
      MICON           =   "frmMntDep.frx":0070
      PICN            =   "frmMntDep.frx":008C
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
      Left            =   15
      Top             =   4425
      Width           =   10365
      _ExtentX        =   18283
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
      Left            =   45
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4845
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
      MICON           =   "frmMntDep.frx":0D5E
      PICN            =   "frmMntDep.frx":0D7A
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
      Left            =   4845
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4845
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
      MICON           =   "frmMntDep.frx":1AB0
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
      Left            =   8205
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4845
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
      MICON           =   "frmMntDep.frx":1ACC
      PICN            =   "frmMntDep.frx":1AE8
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
      Left            =   9270
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4845
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
      MICON           =   "frmMntDep.frx":27BA
      PICN            =   "frmMntDep.frx":27D6
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
      Left            =   30
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":350C
      PICN            =   "frmMntDep.frx":3528
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
      Left            =   1125
      TabIndex        =   21
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":4202
      PICN            =   "frmMntDep.frx":421E
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
      Left            =   2340
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":4AF8
      PICN            =   "frmMntDep.frx":4B14
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
      Left            =   7170
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":5372
      PICN            =   "frmMntDep.frx":538E
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
      Left            =   8160
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":5C68
      PICN            =   "frmMntDep.frx":5C84
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
      Left            =   9270
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmMntDep.frx":6856
      PICN            =   "frmMntDep.frx":6872
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
      Left            =   765
      TabIndex        =   40
      Top             =   3945
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
      Left            =   4845
      TabIndex        =   29
      Top             =   4005
      Width           =   1245
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
      Left            =   9090
      TabIndex        =   28
      Top             =   4005
      Width           =   1245
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
      Left            =   7890
      TabIndex        =   27
      Top             =   60
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
      Left            =   1425
      TabIndex        =   26
      Top             =   60
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   510
      TabIndex        =   25
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   3510
      TabIndex        =   24
      Top             =   4035
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5640
      TabIndex        =   23
      Top             =   90
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   7665
      TabIndex        =   22
      Top             =   4050
      Width           =   1350
   End
End
Attribute VB_Name = "frmMntPer"
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
Dim nif As New clsNIF




Private Sub cbCambioPWD_Click()

If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub
   
With frmCambioPwd
    .otro_usuario = True
    .ID_Otro_usuario = rc.fields("CODIGO").Value
    .Show 1
'    ioCLAVE.SetFocus
End With

End Sub

Private Sub cbCENTROS_gotFocus()
Tab1.Tab = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cbCENTROS_Validate
' Fecha/Hora     : 10/12/2003 12:53
' Autor             : JCastillo
' Propósito       : Cargar el combo de cajas con las cajas correspondientes al
'                       centro seleccionado
'---------------------------------------------------------------------------------------
'
Private Sub cbCENTROS_Validate(Cancel As Boolean)

  On Error GoTo cbCENTROS_Validate_Error
   
If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub

If cbCENTROS.Text <> "" Then

    With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE CODCEN = " & cbCENTROS.Text & " ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "CODCAJA"
    .carga
    Set .DataSource = rc
  End With

End If

   On Error GoTo 0
   Exit Sub

cbCENTROS_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCENTROS_Validate de Formulario frmMntPer"

End Sub

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
        
        If MsgBox("No se encuentra Personal. ¿Crear?", vbYesNo + vbQuestion, "Personal") = vbNo Then
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

Private Sub ioEMAIL_gotFocus()

Tab1.Tab = 0

End Sub

Private Sub ioEMAIL_Validate(Cancel As Boolean)
    
    If Trim(ioEMAIL.Text) <> "" Then
        'devuelve true si el email es correcto
        If Not ValidEmail(ioEMAIL.Text) Then
            
            ioEMAIL.CancelarValidacion
            Cancel = True
                
        Else
            Tab1.Tab = 1
        
        End If
        
    Else
        Tab1.Tab = 1
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
  oSQL.AddTable "PERSONAL"
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
  
   With ioFOTO
  Set .DataSource = rc
        .DataField = "FOTO"
  End With
  
  With ioNOMBRE
  Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "NOMBRE"
        .LongMaxima = 50
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

  With ioPOBLACION
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "POBLACION"
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
  
  With ioAFILIA
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "AFILIA"
        .LongMaxima = 15
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
  
  'With ioCLAVE
  'Set .DataSource = rc
  '      .PermitirBlanco = False
  '      .EsPassword = True
  '      .DataField = "CLAVE"
  '      .LongMaxima = 10
  'End With
  
  
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
  
  With cbBancos
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM BANCOS ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "CODBAN"
    .carga
    Set .DataSource = rc
  End With
  
  With cBFPAGO
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FPAGO ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "FPAGO"
    .carga
    Set .DataSource = rc
  End With
  
  With cbCENTROS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CENTROS ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "CODCEN"
    .carga
    Set .DataSource = rc
  End With
  
    With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "CODCAJA"
    .carga
    Set .DataSource = rc
  End With
  
  With cbTIPPERM
    .añade_item "0   Dependiente"
    .añade_item "1   Supervisor"
    .LenCodigo = 1
    .CodigoWidth = 300
Set .DataSource = rc
    .DataField = "TIPPERM"
    
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
  '  If .State <> 0 Then .Close
  ' End With
   
Set oSQL = Nothing
Set nif = Nothing
'If Me.MDIChild = True Then frmMenuTactil.Show

Set frmMntPer = Nothing

End Sub

Private Sub Form_Resize()
  DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub




Private Sub cbLista_click()

'If rc.EditMode = adEditNone Then


Dim tmprc As New ADODB.Recordset
Dim tmpstrcombo As String
Dim tmpcodcen As String

Dim tmpcodfpago As String
 
'cargar listas
With tmprc
    .Open "SELECT CODIGO, NOMBRE FROM BANCOS ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpstrcombo = frmFlexSimple.fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
    .Close
    .Open "SELECT CODIGO, DESCRIPCION FROM CENTROS ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpcodcen = frmFlexSimple.fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
     .Close
    .Open "SELECT CODIGO, DESCRIPCION FROM FPAGO ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpcodfpago = frmFlexSimple.fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
End With

With frmFlexSimple
        
    With .fg
            Set .DataSource = rc
            .ColComboList(3) = tmpcodcen
            .ColComboList(16) = tmpstrcombo
            .ColComboList(21) = tmpcodfpago
            .ColFormat(1) = "00000"
            .AutoSize 1, .Cols - 1
    End With
    
    .Caption = "Personal ..."
    .Show 1
   ' .SetFocus
End With

tmprc.Close
Set tmprc = Nothing

tmpstrcombo = ""
tmpcodcen = ""
tmpcodfpago = ""


'With frmFlexSimple
'
 '   .Caption = "Personal ..."
  '
   ' With .fg
    '        Set .DataSource = rc
     '       .ColFormat(1) = "000"
      '      DoEvents
'            .AutoSize 1, .Cols - 1
 '           .Refresh
  '  End With
   '
'    .Show 1

'End With

'Else

 '   MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

'End If

End Sub






Private Sub ioCIF_Validate(Cancel As Boolean)

'si esta a blancos salir
If Trim(ioCIF.Text) = "" Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
End If

nif.DarFormato = True
nif.nif = ioCIF.Text

If nif.Err Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioCIF.Text = nif.nif
End If

'comprueba_DNI ioCIF.Text, ioCIF
End Sub





Private Sub ioCUENTA_Validate(Cancel As Boolean)
Tab1.Tab = 0
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
  
    If Trim(ioNOMBRE.Text = "") Then
        lblstatus.Caption = "NOMBRE no puede quedar en blanco"
         bCancel = True
        ioNOMBRE.SetFocus
    End If
    
    If Trim(ioCIF.Text = "") Then
        lblstatus.Caption = "NIF no puede quedar en blanco"
         bCancel = True
        ioCIF.SetFocus
    End If
    
    If Trim(ioDIRECCION.Text = "") Then
        lblstatus.Caption = "DIRECCION no puede quedar en blanco"
        bCancel = True
        ioDIRECCION.SetFocus
    End If
  
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
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from PERSONAL")
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
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
  
  cbAgregar.SetFocus

  Exit Sub
UpdateErr:
   'el error q gebe
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
  cbBancos.Locked = bVal
  cbCENTROS.Locked = bVal
  cBFPAGO.Locked = bVal
  
  cbTIPPERM.Locked = bVal
  
  cbCambioPWD.Enabled = Not bVal
  
  
  
  
  cbCAJAS.Locked = bVal
End Sub

