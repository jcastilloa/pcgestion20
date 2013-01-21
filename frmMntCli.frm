VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMntCli 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
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
   ScaleHeight     =   6870
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   3480
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5385
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":002C
      picn            =   "frmMntCli.frx":004A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   360
      Left            =   15
      Top             =   5010
      Width           =   9900
      _extentx        =   17463
      _extenty        =   635
      caption         =   ""
      fount           =   "frmMntCli.frx":0D1E
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   30
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5385
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":0D4C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":0D78
      picn            =   "frmMntCli.frx":0D96
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   4620
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1111
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":1ACE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":1AFA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdNext 
      Height          =   630
      Left            =   7800
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5385
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":1B18
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":1B44
      picn            =   "frmMntCli.frx":1B62
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   1
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdLast 
      Height          =   630
      Left            =   8865
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5385
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":2836
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":2862
      picn            =   "frmMntCli.frx":2880
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   1
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   15
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   6045
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":35B8
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":35E4
      picn            =   "frmMntCli.frx":3602
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbActualizar 
      Height          =   795
      Left            =   1110
      TabIndex        =   31
      Top             =   6060
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":42DE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":430A
      picn            =   "frmMntCli.frx":4328
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEdicion 
      Height          =   795
      Left            =   2325
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6060
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":4C04
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":4C30
      picn            =   "frmMntCli.frx":4C4E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6780
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6060
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":54AE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":54DA
      picn            =   "frmMntCli.frx":54F8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEliminar 
      Height          =   795
      Left            =   7755
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6060
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":5DD4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":5E00
      picn            =   "frmMntCli.frx":5E1E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCerrar 
      Height          =   795
      Left            =   8865
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   6060
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmMntCli.frx":69F2
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntCli.frx":6A1E
      picn            =   "frmMntCli.frx":6A3C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4245
      Left            =   15
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   390
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmMntCli.frx":7718
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label35"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LOCALIDAD"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label20"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label21"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ioDCTOPP"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ioDCTO"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ioPAIS"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "ioNIF"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ioTITULAR"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ioTELCONTA"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cbREPRESEN"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ioPERCONTA"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ioWEB"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ioEMAIL"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ioFAX"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ioTELEFONO2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "ioTELEFONO1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "ioPROVINCIA"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ioCODPOS"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ioPOBLACION"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ioDIRECCION"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "ioRAZO"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmMntCli.frx":7734
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label34"
      Tab(1).Control(1)=   "Label33"
      Tab(1).Control(2)=   "Label32"
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(4)=   "Label30"
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(6)=   "Label24"
      Tab(1).Control(7)=   "Label25"
      Tab(1).Control(8)=   "BANCO"
      Tab(1).Control(9)=   "Label26"
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(11)=   "Label28"
      Tab(1).Control(12)=   "Label29"
      Tab(1).Control(13)=   "ioFOTO"
      Tab(1).Control(14)=   "ioCPENVIO"
      Tab(1).Control(15)=   "ioDIRECENVIO"
      Tab(1).Control(16)=   "bsGradientLabel1"
      Tab(1).Control(17)=   "ioPAISENVIO"
      Tab(1).Control(18)=   "ioLOCENVIO"
      Tab(1).Control(19)=   "ioPROVENVIO"
      Tab(1).Control(20)=   "ioCUENTA"
      Tab(1).Control(21)=   "ioDC"
      Tab(1).Control(22)=   "ioSUCURSAL"
      Tab(1).Control(23)=   "ioENTIDAD"
      Tab(1).Control(24)=   "ioCODBAN"
      Tab(1).Control(25)=   "ioDIAPAGO2"
      Tab(1).Control(26)=   "ioDIAPAGO1"
      Tab(1).Control(27)=   "ioFCOBRO"
      Tab(1).Control(28)=   "cmBorrarFoto"
      Tab(1).ControlCount=   29
      Begin PCGestion.ucGrdBttn cmBorrarFoto 
         Height          =   315
         Left            =   -65460
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3870
         Width           =   270
         _extentx        =   476
         _extenty        =   556
         caption         =   "X"
         font            =   "frmMntCli.frx":7750
         image           =   "frmMntCli.frx":777C
      End
      Begin PCGestion.miText ioRAZO 
         Height          =   525
         Left            =   1185
         TabIndex        =   0
         Top             =   420
         Width           =   3675
         _extentx        =   6482
         _extenty        =   926
         font            =   "frmMntCli.frx":779A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioDIRECCION 
         Height          =   525
         Left            =   3900
         TabIndex        =   3
         Top             =   975
         Width           =   4590
         _extentx        =   8096
         _extenty        =   926
         font            =   "frmMntCli.frx":77C6
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPOBLACION 
         Height          =   525
         Left            =   1185
         TabIndex        =   5
         Top             =   1515
         Width           =   3600
         _extentx        =   6350
         _extenty        =   926
         font            =   "frmMntCli.frx":77F2
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioCODPOS 
         Height          =   525
         Left            =   8985
         TabIndex        =   4
         Top             =   975
         Width           =   765
         _extentx        =   1349
         _extenty        =   926
         font            =   "frmMntCli.frx":781E
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPROVINCIA 
         Height          =   525
         Left            =   5895
         TabIndex        =   6
         Top             =   1515
         Width           =   3855
         _extentx        =   6800
         _extenty        =   926
         font            =   "frmMntCli.frx":784A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioTELEFONO1 
         Height          =   525
         Left            =   5880
         TabIndex        =   9
         Top             =   2070
         Width           =   1665
         _extentx        =   2937
         _extenty        =   926
         font            =   "frmMntCli.frx":7876
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioTELEFONO2 
         Height          =   525
         Left            =   8115
         TabIndex        =   10
         Top             =   2070
         Width           =   1635
         _extentx        =   2884
         _extenty        =   926
         font            =   "frmMntCli.frx":78A2
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFAX 
         Height          =   525
         Left            =   3645
         TabIndex        =   8
         Top             =   2070
         Width           =   1665
         _extentx        =   3201
         _extenty        =   926
         font            =   "frmMntCli.frx":78CE
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioEMAIL 
         Height          =   525
         Left            =   1185
         TabIndex        =   11
         Top             =   2610
         Width           =   4125
         _extentx        =   7276
         _extenty        =   926
         font            =   "frmMntCli.frx":78FA
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioWEB 
         Height          =   525
         Left            =   5880
         TabIndex        =   12
         Top             =   2610
         Width           =   3870
         _extentx        =   6826
         _extenty        =   926
         font            =   "frmMntCli.frx":7926
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPERCONTA 
         Height          =   525
         Left            =   1185
         TabIndex        =   13
         Top             =   3135
         Width           =   4305
         _extentx        =   7594
         _extenty        =   926
         font            =   "frmMntCli.frx":7952
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbREPRESEN 
         Height          =   555
         Left            =   1185
         TabIndex        =   15
         Top             =   3660
         Width           =   4305
         _extentx        =   7594
         _extenty        =   979
         font            =   "frmMntCli.frx":797E
      End
      Begin PCGestion.miText ioTELCONTA 
         Height          =   525
         Left            =   7935
         TabIndex        =   14
         Top             =   3135
         Width           =   1815
         _extentx        =   3201
         _extenty        =   926
         font            =   "frmMntCli.frx":79AA
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioTITULAR 
         Height          =   525
         Left            =   5775
         TabIndex        =   1
         Top             =   450
         Width           =   3975
         _extentx        =   7011
         _extenty        =   926
         font            =   "frmMntCli.frx":79D6
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo ioFCOBRO 
         Height          =   555
         Left            =   -73755
         TabIndex        =   18
         Top             =   375
         Width           =   4365
         _extentx        =   7699
         _extenty        =   979
         enabled         =   0   'False
         font            =   "frmMntCli.frx":7A02
      End
      Begin PCGestion.miText ioDIAPAGO1 
         Height          =   525
         Left            =   -67605
         TabIndex        =   19
         Top             =   375
         Width           =   600
         _extentx        =   1058
         _extenty        =   926
         font            =   "frmMntCli.frx":7A2E
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioDIAPAGO2 
         Height          =   525
         Left            =   -65715
         TabIndex        =   20
         Top             =   375
         Width           =   570
         _extentx        =   1005
         _extenty        =   926
         font            =   "frmMntCli.frx":7A5A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo ioCODBAN 
         Height          =   495
         Left            =   -73755
         TabIndex        =   21
         Top             =   900
         Width           =   4365
         _extentx        =   7699
         _extenty        =   873
         enabled         =   0   'False
         font            =   "frmMntCli.frx":7A86
      End
      Begin PCGestion.miText ioENTIDAD 
         Height          =   525
         Left            =   -73755
         TabIndex        =   22
         Top             =   1395
         Width           =   825
         _extentx        =   1455
         _extenty        =   926
         font            =   "frmMntCli.frx":7AB2
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioSUCURSAL 
         Height          =   525
         Left            =   -72120
         TabIndex        =   23
         Top             =   1395
         Width           =   675
         _extentx        =   1191
         _extenty        =   926
         font            =   "frmMntCli.frx":7ADE
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioDC 
         Height          =   525
         Left            =   -71115
         TabIndex        =   24
         Top             =   1395
         Width           =   495
         _extentx        =   873
         _extenty        =   926
         font            =   "frmMntCli.frx":7B0A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioCUENTA 
         Height          =   525
         Left            =   -69840
         TabIndex        =   25
         Top             =   1395
         Width           =   1305
         _extentx        =   2302
         _extenty        =   926
         font            =   "frmMntCli.frx":7B36
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioNIF 
         Height          =   525
         Left            =   1185
         TabIndex        =   2
         Top             =   975
         Width           =   1545
         _extentx        =   2725
         _extenty        =   926
         font            =   "frmMntCli.frx":7B62
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPROVENVIO 
         Height          =   525
         Left            =   -73755
         TabIndex        =   28
         Top             =   3255
         Width           =   2775
         _extentx        =   4895
         _extenty        =   926
         font            =   "frmMntCli.frx":7B8E
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioLOCENVIO 
         Height          =   525
         Left            =   -73755
         TabIndex        =   27
         Top             =   2775
         Width           =   4365
         _extentx        =   7699
         _extenty        =   926
         font            =   "frmMntCli.frx":7BBA
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPAISENVIO 
         Height          =   525
         Left            =   -73755
         TabIndex        =   30
         Top             =   3735
         Width           =   2775
         _extentx        =   4895
         _extenty        =   926
         font            =   "frmMntCli.frx":7BE6
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.bsGradientLabel bsGradientLabel1 
         Height          =   375
         Left            =   -74895
         Top             =   1905
         Width           =   6345
         _extentx        =   11192
         _extenty        =   661
         caption         =   "Datos de Envío"
         fount           =   "frmMntCli.frx":7C12
         captioncolour   =   0
         colour1         =   15640462
         colour2         =   7177785
         captionalignment=   1
      End
      Begin PCGestion.miText ioDIRECENVIO 
         Height          =   525
         Left            =   -73755
         TabIndex        =   26
         Top             =   2280
         Width           =   5235
         _extentx        =   9234
         _extenty        =   926
         font            =   "frmMntCli.frx":7C40
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioCPENVIO 
         Height          =   525
         Left            =   -70305
         TabIndex        =   29
         Top             =   3255
         Width           =   915
         _extentx        =   1614
         _extenty        =   926
         font            =   "frmMntCli.frx":7C6C
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioPAIS 
         Height          =   525
         Left            =   1185
         TabIndex        =   7
         Top             =   2085
         Width           =   1995
         _extentx        =   3519
         _extenty        =   926
         font            =   "frmMntCli.frx":7C98
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioDCTO 
         Height          =   525
         Left            =   6525
         TabIndex        =   16
         Top             =   3660
         Width           =   975
         _extentx        =   1720
         _extenty        =   926
         font            =   "frmMntCli.frx":7CC4
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioDCTOPP 
         Height          =   525
         Left            =   8775
         TabIndex        =   17
         Top             =   3660
         Width           =   975
         _extentx        =   1720
         _extenty        =   926
         font            =   "frmMntCli.frx":7CF0
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin VB.PictureBox ioFOTO 
         Height          =   3345
         Left            =   -68520
         ScaleHeight     =   3285
         ScaleWidth      =   3285
         TabIndex        =   86
         Top             =   855
         Width           =   3345
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   360
         Left            =   8595
         TabIndex        =   53
         Top             =   3735
         Width           =   225
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   360
         Left            =   7410
         TabIndex        =   85
         Top             =   3735
         Width           =   225
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TITULAR"
         Height          =   345
         Left            =   4860
         TabIndex        =   84
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DCTO PP"
         Height          =   360
         Left            =   7755
         TabIndex        =   83
         Top             =   3735
         Width           =   915
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DCTO"
         Height          =   360
         Left            =   5880
         TabIndex        =   82
         Top             =   3735
         Width           =   630
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEF. CONTACTO"
         Height          =   360
         Left            =   5985
         TabIndex        =   81
         Top             =   3225
         Width           =   1905
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REPRESEN."
         Height          =   300
         Left            =   60
         TabIndex        =   80
         Top             =   3765
         Width           =   1140
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PERSONA CONTACTO"
         Height          =   570
         Left            =   60
         TabIndex        =   79
         Top             =   3090
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RAZO / NOMBRE"
         Height          =   615
         Left            =   180
         TabIndex        =   78
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION"
         Height          =   300
         Left            =   2775
         TabIndex        =   77
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CP"
         Height          =   300
         Left            =   8625
         TabIndex        =   76
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA"
         Height          =   300
         Left            =   4725
         TabIndex        =   75
         Top             =   1590
         Width           =   1140
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TLF 1"
         Height          =   360
         Left            =   5220
         TabIndex        =   74
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TLF 2"
         Height          =   360
         Left            =   7140
         TabIndex        =   73
         Top             =   2145
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAX"
         Height          =   360
         Left            =   3060
         TabIndex        =   72
         Top             =   2145
         Width           =   555
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         Height          =   300
         Left            =   45
         TabIndex        =   71
         Top             =   2685
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "WEB"
         Height          =   300
         Left            =   5325
         TabIndex        =   70
         Top             =   2715
         Width           =   510
      End
      Begin VB.Label LOCALIDAD 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOCALIDAD"
         Height          =   300
         Left            =   45
         TabIndex        =   69
         Top             =   1605
         Width           =   1140
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFICINA"
         Height          =   300
         Left            =   -72930
         TabIndex        =   68
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ENTIDAD"
         Height          =   330
         Left            =   -74685
         TabIndex        =   67
         Top             =   1485
         Width           =   915
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC"
         Height          =   300
         Left            =   -71400
         TabIndex        =   66
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA"
         Height          =   330
         Left            =   -70620
         TabIndex        =   65
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label BANCO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO"
         Height          =   330
         Left            =   -74910
         TabIndex        =   64
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIA PAGO 1"
         Height          =   360
         Left            =   -68835
         TabIndex        =   63
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIA PAGO 2"
         Height          =   360
         Left            =   -67050
         TabIndex        =   62
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA DE COBRO"
         Height          =   615
         Left            =   -74865
         TabIndex        =   61
         Top             =   345
         Width           =   1110
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NIF"
         Height          =   300
         Left            =   660
         TabIndex        =   60
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOCALIDAD"
         Height          =   300
         Left            =   -74910
         TabIndex        =   59
         Top             =   2865
         Width           =   1140
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   58
         Top             =   3360
         Width           =   1140
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CP"
         Height          =   300
         Left            =   -70815
         TabIndex        =   57
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION"
         Height          =   300
         Left            =   -74880
         TabIndex        =   56
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAIS"
         Height          =   300
         Left            =   -74295
         TabIndex        =   55
         Top             =   3825
         Width           =   510
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAIS"
         Height          =   300
         Left            =   705
         TabIndex        =   54
         Top             =   2160
         Width           =   450
      End
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   615
      TabIndex        =   50
      Top             =   4620
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
      Left            =   8670
      TabIndex        =   39
      Top             =   4650
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
      Left            =   5385
      TabIndex        =   38
      Top             =   4650
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
      Left            =   7455
      TabIndex        =   37
      Top             =   15
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
      Left            =   1215
      TabIndex        =   36
      Top             =   15
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   300
      TabIndex        =   35
      Top             =   30
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   7320
      TabIndex        =   34
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5205
      TabIndex        =   33
      Top             =   45
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   3990
      TabIndex        =   32
      Top             =   4665
      Width           =   1350
   End
End
Attribute VB_Name = "frmMntCli"
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



Private Sub cmBorrarFoto_Click()

If TipoServer = 1 Then

If mbEditFlag = False And mbAddNewFlag = False Then Exit Sub

If MsgBox("¿Desea quitar la imagen?", vbQuestion + vbYesNo, titulo) = vbYes Then

    rc.fields(ioFOTO.DataField).Value = Null
    ioFOTO.Picture = Nothing

End If

End If

End Sub

Private Sub Form_Activate()

If Not prime Then

  If TipoServer = 1 Then
  
        If rc.RecordCount = 0 Then
            If MsgBox("No se encuentran Clientes. ¿Crear?", vbYesNo + vbQuestion, "Clientes") = vbNo Then
                Unload Me
            Else
                Call cbAgregar_Click
            End If
        
        Else
         
                Call cmdFirst_Click
                Call cbCancelar_Click
        
        End If
        
  Else
        
        Call cmdFirst_Click
        Call cbCancelar_Click

        
  End If
  

prime = True
End If
    
End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000000")
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
  oSQL.AddTable "CLIENTES"
  oSQL.AddOrderClause "CODIGO"
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

  
  Call Enlaza_Controles
  
  With ioDCTO
 ' .displayformat = "00.00 %"
 ' .Format = "##.##"
 ' Set .DataSource = rc
 '     .DataField = "DCTO"
    .Alineacion = 1
   ' .dspFormat = "00.00"
    .SoloNumeros = True
    .LongMaxima = 5
  End With
  
  With ioDCTOPP
 ' .displayformat = "00.00 %"
 ' .Format = "##.##"
 ' Set .DataSource = rc
  '    .DataField = "DCTOPP"
    .Alineacion = 1
   ' .dspFormat = "00.00"
    .SoloNumeros = True
    .LongMaxima = 5
  End With
  
  Dialogo.Filter = "*.bmp; *.gif|*.bmp; *.gif"
       
  mbDataChanged = False
End Sub


Private Sub Enlaza_Controles()
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioRAZO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "RAZO"
        .LongMaxima = 40
  End With
  
  With ioTITULAR
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TITULAR"
        .LongMaxima = 40
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
        .SoloNumeros = True
        .Alineacion = 1
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
        .LongMaxima = 25
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
  
  With ioWEB
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "WEB"
        .LongMaxima = 50
  End With
  
  With ioPERCONTA
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PERCONTA"
        .LongMaxima = 40
  End With
  
  With ioTELCONTA
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELCONTA"
        .LongMaxima = 17
  End With

  With ioNIF
  Set .DataSource = rc
        .DataField = "NIF"
        .LongMaxima = 15
        .PermitirBlanco = False
  End With
  
  
  With ioDIAPAGO1
  Set .DataSource = rc
        .DataField = "DIAPAGO1"
        .LongMaxima = 2
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioDIAPAGO2
  Set .DataSource = rc
        .DataField = "DIAPAGO2"
        .LongMaxima = 2
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  
  With ioENTIDAD
  Set .DataSource = rc
        .DataField = "ENTIDAD"
        .LongMaxima = 4
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioSUCURSAL
  Set .DataSource = rc
        .DataField = "SUCURSAL"
        .LongMaxima = 4
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioDC
  Set .DataSource = rc
        .DataField = "DC"
        .LongMaxima = 2
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioCUENTA
  Set .DataSource = rc
        .DataField = "CUENTA"
        .LongMaxima = 10
        .PermitirBlanco = True
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioDIRECENVIO
  Set .DataSource = rc
        .DataField = "DIRECENVIO"
        .LongMaxima = 40
        .PermitirBlanco = True
  End With
  
  With ioCPENVIO
  Set .DataSource = rc
        .DataField = "CPENVIO"
        .LongMaxima = 5
        .PermitirBlanco = True
  End With
  
  With ioLOCENVIO
  Set .DataSource = rc
        .DataField = "LOCENVIO"
        .LongMaxima = 40
        .PermitirBlanco = True
  End With
  
  With ioPAISENVIO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PAISENVIO"
        .LongMaxima = 25
  End With
  
  With cbREPRESEN
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM REPRESEN WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .DataField = "REPRESEN"
    .carga
    Set .DataSource = rc
  End With
  
  With ioCODBAN
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM BANCOS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 2
    .CodigoWidth = 700
    .DataField = "CODBAN"
    .carga
    Set .DataSource = rc
  End With
  
    'Forma de pago (que nos pagan, forma de cobro para los clientes)
  With ioFCOBRO
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FCOBRO WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 3
    .CodigoWidth = 700
    .DataField = "FCOBRO"
    .carga
    Set .DataSource = rc
  End With
  
  If TipoServer = 1 Then
  
   With ioFOTO
    .DataField = "FOTO"
    Set .DataSource = rc
  End With
  
  End If
  
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
End Sub

Private Sub Des_Enlaza_Controles()

 With ioCODIGO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioRAZO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioTITULAR
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioDIRECCION
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCODPOS
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioPOBLACION
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioPROVINCIA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioPAIS
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioTELEFONO1
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioTELEFONO2
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFAX
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioEMAIL
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioWEB
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioPERCONTA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioTELCONTA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioNIF
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioDIAPAGO1
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioDIAPAGO2
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioENTIDAD
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioSUCURSAL
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioDC
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCUENTA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioDIRECENVIO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCPENVIO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioLOCENVIO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioPAISENVIO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With cbREPRESEN
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCODBAN
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFCOBRO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFOTO
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFBAJA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFALTA
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioFMODI
  Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioMBAJA
  Set .DataSource = Nothing
        .DataField = ""
  End With

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

 '  With locCnn
 '   If .State <> 0 Then .Close
 '  End With

Set oSQL = Nothing
Set nif = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmMntCli = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cbLista_click
' DateTime  : 01/11/2003 21:21
' Author    : Administrador
' Purpose   : Cargar strings para los colcombolists
'---------------------------------------------------------------------------------------
'
Private Sub cbLista_click()
'Dim tmprc As New ADODB.Recordset
'Dim tmpcodrep As String
'Dim tmpcodban As String
'Dim tmpcodfcobro As String

   On Error GoTo cbLista_click_Error

'With tmprc
'tmprc.Open "SELECT CODIGO, NOMBRE FROM BANCOS WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
'tmpcodban = frmFlexSimple.fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
'.Close
'tmprc.Open "SELECT CODIGO, DESCRIPCION FROM FCOBRO WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
'tmpcodfcobro = frmFlexSimple.fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
'.Close
'tmprc.Open "SELECT CODIGO, NOMBRE FROM REPRESEN WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
'tmpcodrep = frmFlexSimple.fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
'.Close
'End With

With frmFlexCli

    .Caption = "Clientes ..."
    Set .miosql = oSQL
            
    Set .miRc = rc
    
    'With .fg
         '   Set .DataSource = rc
    '        .ColComboList(17) = tmpcodrep
    '        .ColComboList(22) = tmpcodfcobro
    '        .ColComboList(25) = tmpcodban
    '        .ColFormat(1) = "00000"
    '        .AutoSize 1, .Cols - 1
    'End With
    
    DoEvents
    
    Call Des_Enlaza_Controles
    DoEvents
    
    .Show 1
    
    Set frmFlexCli = Nothing
    
    DoEvents
    Call Enlaza_Controles

End With

'tmpcodfcobro = ""
'tmpcodban = ""
'tmpcodrep = ""

   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cbLista_click of Formulario frmMntCli"

End Sub


Private Sub ioDCTO_Validate(Cancel As Boolean)

On Error GoTo ioDCTO_Validate_Error

With ioDCTO

If Trim(.Text) <> "" Then
    If CDbl(Replace(.Text, ".", ",")) >= 100 Then
    lblstatus.Caption = "No se permite un Descuento mayor o igual al 100%"
  '  .CancelarValidacion
    Cancel = True
    Else
        lblstatus.Caption = ""
    End If
End If

End With

   On Error GoTo 0
   Exit Sub

ioDCTO_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDCTO_Validate of Formulario frmMntCli"

End Sub

Private Sub ioDCTOPP_Validate(Cancel As Boolean)

On Error GoTo ioDCTOPP_Validate_Error

With ioDCTOPP
If Trim(.Text) <> "" Then
    If Replace(.Text, ".", ",") >= 100 Then
        lblstatus.Caption = "No se permite un Descuento por Pronto Pago mayor o igual al 100%"
        .CancelarValidacion
        Cancel = True
    Else
        lblstatus.Caption = ""
        Tab1.Tab = 1
    End If
Else
    Tab1.Tab = 1
End If
End With

   On Error GoTo 0
   Exit Sub

ioDCTOPP_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDCTOPP_Validate of Formulario frmMntCli"

End Sub

Private Sub ioDIAPAGO1_Validate(Cancel As Boolean)

   On Error GoTo ioDIAPAGO1_Validate_Error

If Trim(ioDIAPAGO1.Text) <> "" Then

    If CByte(ioDIAPAGO1.Text) < 1 Or CByte(ioDIAPAGO1.Text) > 31 Then

        Cancel = True
        ioDIAPAGO1.CancelarValidacion
        lblstatus.Caption = "Día de pago incorrecto (1-31)"
        
    Else
        lblstatus.Caption = ""
    End If

End If

   On Error GoTo 0
   Exit Sub

ioDIAPAGO1_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDIAPAGO1_Validate of Formulario frmMntCli"

End Sub

Private Sub ioDIAPAGO2_Validate(Cancel As Boolean)

   On Error GoTo ioDIAPAGO2_Validate_Error

If Trim(ioDIAPAGO2.Text) <> "" Then

    If CByte(ioDIAPAGO2.Text) < 1 Or CByte(ioDIAPAGO2.Text) > 31 Then

        Cancel = True
        ioDIAPAGO2.CancelarValidacion
        lblstatus.Caption = "Día de pago incorrecto (1-31)"
        
    Else
        lblstatus.Caption = ""
    End If

End If

   On Error GoTo 0
   Exit Sub

ioDIAPAGO2_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDIAPAGO2_Validate of Formulario frmMntCli"

End Sub



Private Sub ioEMAIL_Validate(Cancel As Boolean)
    
   On Error GoTo ioEMAIL_Validate_Error

    If Trim(ioEMAIL.Text) <> "" Then
        'devuelve true si el email es correcto
        If Not ValidEmail(ioEMAIL.Text) Then
            
            ioEMAIL.CancelarValidacion
            Cancel = True
                
        End If
        
    End If

   On Error GoTo 0
   Exit Sub

ioEMAIL_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioEMAIL_Validate of Formulario frmMntCli"
    
End Sub

Private Sub ioNIF_Validate(Cancel As Boolean)

'si esta a blancos salir
   On Error GoTo ioNIF_Validate_Error

If Trim(ioNIF.Text) = "" Then
    ioNIF.CancelarValidacion
    Cancel = True
    Exit Sub
End If

nif.DarFormato = True
nif.nif = ioNIF.Text

If nif.Err Then
    ioNIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioNIF.Text = nif.nif
End If

   On Error GoTo 0
   Exit Sub

ioNIF_Validate_Error:
    Cancel = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioNIF_Validate of Formulario frmMntCli"
 
End Sub

Private Sub ioPAISENVIO_Validate(Cancel As Boolean)
Tab1.Tab = 0
End Sub

Private Sub ioRAZO_Validate(Cancel As Boolean)

'si es blanco, cancelar validación
If Trim(ioRAZO.Text) = "" Then
    ioRAZO.CancelarValidacion
    Cancel = True
    lblstatus.Caption = "Razón Social no puede estar en blanco"
End If

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
  If rc.AbsolutePosition > 0 Then
  
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
  
  ioDCTO.Text = rc.fields("DCTO")
  ioDCTOPP.Text = rc.fields("DCTOPP")
  
  End If
  
End Sub

Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
    
   '     If ioRAZO.Text = "" Then
    '        lblstatus.Caption = "Razón Social no puede estar en blanco"
     '       ioRAZO.SetFocus
      '  End If
    
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
  
        With ioRAZO
        If .Text = "" Then
            lblstatus.Caption = "Razón Social no puede estar en blanco"
            .CancelarValidacion
            .SetFocus
        End If
        End With
        
        With ioNIF
        If .Text = "" Then
            lblstatus.Caption = "NIF/CIF no puede estar en blanco"
            .CancelarValidacion
            .SetFocus
            bCancel = True
        End If
        End With
         
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
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from clientes where codcaja = " & CajaActual)

    'tmpcodigo = devuelve_campo("select max(codigo) + 1 from clientes")
    .AddNew '"CODIGO", tmpcodigo
    
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    .fields("CODCAJA") = CajaActual
    
    'End If
    
    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    Tab1.SetFocus
    ioRAZO.SetFocus
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
  ioRAZO.SetFocus
  
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
 
 
  If ioDCTO.Text = "" Then ioDCTO.Text = 0
  If ioDCTOPP.Text = "" Then ioDCTOPP.Text = 0
    
  rc.fields("DCTO") = ioDCTO.Text
  rc.fields("DCTOPP") = ioDCTOPP.Text
  
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
If Err.Number = -2147467259 Then Exit Sub
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
  cbREPRESEN.Locked = bVal
  ioFCOBRO.Locked = bVal
  ioCODBAN.Locked = bVal
End Sub




