VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntEmp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11130
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
   ScaleHeight     =   5970
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4485
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
      MICON           =   "frmMntEmp.frx":0000
      PICN            =   "frmMntEmp.frx":001C
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
      Top             =   4065
      Width           =   11100
      _ExtentX        =   19579
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
   Begin PCGestion.miText ioRAZO 
      Height          =   525
      Left            =   2880
      TabIndex        =   0
      Top             =   135
      Width           =   6090
      _ExtentX        =   10742
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
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4485
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
      MICON           =   "frmMntEmp.frx":0CEE
      PICN            =   "frmMntEmp.frx":0D0A
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
      Left            =   5145
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4485
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
      MICON           =   "frmMntEmp.frx":1A40
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
      Left            =   9000
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4485
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
      MICON           =   "frmMntEmp.frx":1A5C
      PICN            =   "frmMntEmp.frx":1A78
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
      Left            =   10065
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4485
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
      MICON           =   "frmMntEmp.frx":274A
      PICN            =   "frmMntEmp.frx":2766
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":349C
      PICN            =   "frmMntEmp.frx":34B8
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
      TabIndex        =   29
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":4192
      PICN            =   "frmMntEmp.frx":41AE
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":4A88
      PICN            =   "frmMntEmp.frx":4AA4
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
      Left            =   7965
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":5302
      PICN            =   "frmMntEmp.frx":531E
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
      Left            =   8940
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":5BF8
      PICN            =   "frmMntEmp.frx":5C14
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
      Left            =   10065
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5145
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
      MICON           =   "frmMntEmp.frx":67E6
      PICN            =   "frmMntEmp.frx":6802
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioCIF 
      Height          =   480
      Left            =   9525
      TabIndex        =   1
      Top             =   135
      Width           =   1545
      _ExtentX        =   2725
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
   Begin PCGestion.miText ioDIRECC 
      Height          =   450
      Left            =   1125
      TabIndex        =   2
      Top             =   675
      Width           =   4290
      _ExtentX        =   7567
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
   Begin PCGestion.miText ioPROVIN 
      Height          =   525
      Left            =   1125
      TabIndex        =   5
      Top             =   1155
      Width           =   3525
      _ExtentX        =   6218
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
   Begin PCGestion.miText ioLOCALI 
      Height          =   525
      Left            =   7665
      TabIndex        =   4
      Top             =   690
      Width           =   3450
      _ExtentX        =   6085
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
      Left            =   5745
      TabIndex        =   3
      Top             =   675
      Width           =   765
      _ExtentX        =   1349
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
   Begin PCGestion.miText ioTELEF 
      Height          =   525
      Left            =   5745
      TabIndex        =   6
      Top             =   1155
      Width           =   1485
      _ExtentX        =   2619
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
      Left            =   7665
      TabIndex        =   7
      Top             =   1155
      Width           =   1650
      _ExtentX        =   2910
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
   Begin PCGestion.miText ioNOMBRE 
      Height          =   525
      Left            =   1125
      TabIndex        =   8
      Top             =   1635
      Width           =   6375
      _ExtentX        =   11245
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
   Begin PCGestion.miText ioIPCLI 
      Height          =   525
      Left            =   5595
      TabIndex        =   13
      Top             =   2115
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
   Begin PCGestion.miText ioBBDDCLI 
      Height          =   525
      Left            =   3345
      TabIndex        =   11
      Top             =   2115
      Width           =   1485
      _ExtentX        =   2619
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
   Begin PCGestion.miText ioIPSRV 
      Height          =   525
      Left            =   5580
      TabIndex        =   14
      Top             =   2550
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
   Begin PCGestion.miText ioBBDDSRV 
      Height          =   525
      Left            =   3345
      TabIndex        =   12
      Top             =   2565
      Width           =   1485
      _ExtentX        =   2619
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
   Begin PCGestion.miText ioCL2 
      Height          =   525
      Left            =   4290
      TabIndex        =   16
      Top             =   3045
      Width           =   3180
      _ExtentX        =   5609
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
   Begin PCGestion.miText ioCL1 
      Height          =   525
      Left            =   810
      TabIndex        =   15
      Top             =   3045
      Width           =   2625
      _ExtentX        =   4630
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
   Begin PCGestion.miText ioPL1 
      Height          =   525
      Left            =   810
      TabIndex        =   18
      Top             =   3525
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
   Begin PCGestion.miText ioCL3 
      Height          =   525
      Left            =   8325
      TabIndex        =   17
      Top             =   3045
      Width           =   2760
      _ExtentX        =   4868
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
   Begin PCGestion.miText ioPL2 
      Height          =   525
      Left            =   5430
      TabIndex        =   19
      Top             =   3540
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
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA 5"
      Height          =   300
      Left            =   4605
      TabIndex        =   50
      Top             =   3630
      Width           =   780
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA 2"
      Height          =   300
      Left            =   3465
      TabIndex        =   49
      Top             =   3135
      Width           =   780
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA 1"
      Height          =   300
      Left            =   30
      TabIndex        =   48
      Top             =   3105
      Width           =   780
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA 3"
      Height          =   300
      Left            =   7500
      TabIndex        =   47
      Top             =   3135
      Width           =   780
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA 4"
      Height          =   300
      Left            =   30
      TabIndex        =   46
      Top             =   3585
      Width           =   780
   End
   Begin MSForms.CheckBox ioCREADASRV 
      Height          =   435
      Left            =   1065
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1185
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2090;767"
      Value           =   "0"
      Caption         =   "EXISTE"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox ioCREADACLI 
      Height          =   435
      Left            =   1065
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1170
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2064;767"
      Value           =   "0"
      Caption         =   "EXISTE"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP SRV"
      Height          =   300
      Left            =   4935
      TabIndex        =   45
      Top             =   2670
      Width           =   630
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BBDD SRV"
      Height          =   300
      Left            =   2340
      TabIndex        =   44
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BBDD LOC"
      Height          =   300
      Left            =   2325
      TabIndex        =   43
      Top             =   2175
      Width           =   990
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP CLI"
      Height          =   300
      Left            =   4995
      TabIndex        =   42
      Top             =   2220
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   300
      Left            =   240
      TabIndex        =   41
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      Height          =   300
      Left            =   4635
      TabIndex        =   40
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      Height          =   330
      Left            =   7170
      TabIndex        =   39
      Top             =   1230
      Width           =   450
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA"
      Height          =   330
      Left            =   -45
      TabIndex        =   38
      Top             =   1245
      Width           =   1155
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOCALIDAD"
      Height          =   300
      Left            =   6555
      TabIndex        =   37
      Top             =   765
      Width           =   1125
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CP"
      Height          =   330
      Left            =   5265
      TabIndex        =   36
      Top             =   750
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION"
      Height          =   330
      Left            =   -105
      TabIndex        =   35
      Top             =   750
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CIF"
      Height          =   330
      Left            =   9120
      TabIndex        =   34
      Top             =   210
      Width           =   360
   End
   Begin VB.Label ioID 
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
      Left            =   1155
      TabIndex        =   22
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   240
      TabIndex        =   21
      Top             =   225
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL"
      Height          =   705
      Left            =   1980
      TabIndex        =   20
      Top             =   75
      Width           =   900
   End
End
Attribute VB_Name = "frmMntEmp"
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


Public Configuracion_Inicial As Boolean

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

'If Trim(ioCIF.Text) <> "" Then Call comprueba_DNI(Trim(ioCIF.Text), ioCIF)
End Sub

Private Sub ioID_Change()
ioID.Caption = Format(ioID.Caption, "000")
End Sub

Private Sub Form_Activate()
  
  If Not prime Then
  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Empresas. ¿Crear?", vbYesNo + vbQuestion, "Empresas") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
        Call cmdFirst_Click
        Call cbCancelar_Click
        
  End If
  
  'si entramos desde la configuración inicial, editar el registro introducido
  'por defecto
  If Configuracion_Inicial Then
    Call cbedicion_Click
  End If
  
  prime = True
  
  End If
    
End Sub

Private Sub Form_Load()
  
  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  Set EmpCnn = New ADODB.Connection
  
  With EmpCnn
    
        .CursorLocation = adUseClient
        .Open strEmpCnn
   
  End With
   
  Set rc = New Recordset
  oSQL.AddTable "EMPRESAS"
  oSQL.AddOrderClause "ID"
  rc.Open oSQL.SQL, EmpCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioID
  Set .DataSource = rc
        .DataField = "ID"
  End With
  
  With ioRAZO
    Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "RAZO"
        .LongMaxima = 50
  End With
  
  With ioCIF
        .PermitirBlanco = False
    Set .DataSource = rc
        .DataField = "CIF"
        .LongMaxima = 12
  End With
  
  With ioDIRECC
    Set .DataSource = rc
        .DataField = "DIRECC"
        .LongMaxima = 50
  End With
  
  With ioCODPOS
  Set .DataSource = rc
        .DataField = "CODPOS"
        .LongMaxima = 5
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioLOCALI
    Set .DataSource = rc
        .DataField = "LOCALI"
        .LongMaxima = 20
  End With
  
  With ioPROVIN
  Set .DataSource = rc
        .DataField = "PROVIN"
        .LongMaxima = 20
  End With
  
  
  With ioTELEF
  Set .DataSource = rc
        .DataField = "TELEF"
        .LongMaxima = 12
  End With
  
  With ioFAX
    Set .DataSource = rc
        .DataField = "FAX"
        .LongMaxima = 12
  End With
  
    
  With ioNOMBRE
    Set .DataSource = rc
        .DataField = "NOMBRE"
        .LongMaxima = 50
  End With
  
  With ioBBDDSRV
  Set .DataSource = rc
        .DataField = "BBDDSRV"
        .LongMaxima = 10
  End With
  
  With ioBBDDCLI
    Set .DataSource = rc
        .DataField = "BBDDCLI"
        .LongMaxima = 10
        .PermitirBlanco = False
  End With
  
  
  With ioCREADASRV
  Set .DataSource = rc
        .DataField = "CREADASRV"
        .Locked = True
  End With
  
  With ioCREADACLI
  Set .DataSource = rc
        .DataField = "CREADACLI"
        .Locked = True
  End With
  
  
  With ioIPSRV
  Set .DataSource = rc
        .DataField = "IPSRV"
        .LongMaxima = 50
        .PermitirBlanco = False
  End With
  
  With ioIPCLI
    Set .DataSource = rc
        .DataField = "IPCLI"
        .LongMaxima = 50
        .PermitirBlanco = False
  End With
  
  With ioCL1
    Set .DataSource = rc
        .DataField = "CL1"
        .LongMaxima = 30
        .PermitirBlanco = True
  End With
  
  With ioCL2
    Set .DataSource = rc
        .DataField = "CL2"
        .LongMaxima = 30
        .PermitirBlanco = True
  End With
  
  With ioCL3
    Set .DataSource = rc
        .DataField = "CL3"
        .LongMaxima = 30
        .PermitirBlanco = True
  End With
         
  With ioPL1
    Set .DataSource = rc
        .DataField = "PL1"
        .LongMaxima = 30
        .PermitirBlanco = True
  End With
  
    With ioPL2
    Set .DataSource = rc
        .DataField = "PL2"
        .LongMaxima = 30
        .PermitirBlanco = True
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
            
            cbLista_click
      
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

EmpCnn.Close
Set EmpCnn = Nothing

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

Set oSQL = Nothing
Set nif = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmMntEmp = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub cbLista_click()

'If rc.EditMode = adEditNone Then

With frmFlexSimple

    .Caption = "Empresas ..."
        
    With .fg
            Set .DataSource = rc
            .ColFormat(1) = "000"
            DoEvents
            .AutoSize 1, .Cols - 1
            .Refresh
    End With
    
    .Show 1

End With

'Else

    'MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

'End If

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
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cbAgregar_Click()
  'Dim tmpcodigo As Variant
  
  On Error GoTo AddErr
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    'tmpcodigo = devuelve_campo("select max(codigo) + 1 from secciones")
    
    'If tmpcodigo <> "@" Then
    '
    '.Fields("CODIGO") = tmpcodigo
    
    'End If

    'Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar empresa"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
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

  lblstatus.Caption = "Modificar datos de empresa"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
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
  

End Sub

Private Sub cbactualizar_Click()
Dim tmpid As Long
Dim tmpcreadasrv As Boolean
Dim tmpcreadacli As Boolean
Dim tmpbbddsrv As String
Dim tmpbbddcli As String


  On Error GoTo UpdateErr
  
  If Trim(ioRAZO.Text) = "" Then
    MsgBox "No se permite Razón Social en blanco", vbInformation, titulo
    ioRAZO.SetFocus
    Exit Sub
  End If
  
  If Trim(ioCIF.Text) = "" Then
    MsgBox "No se permite CIF en blanco", vbInformation, titulo
    ioCIF.SetFocus
    Exit Sub
  End If
  
  If Trim(ioNOMBRE.Text) = "" Then
    MsgBox "No se permite NOMBRE del TITULAR en blanco", vbInformation, titulo
    ioNOMBRE.SetFocus
    Exit Sub
  End If
 
  With rc
  
  If .EditMode = adEditNone Then Exit Sub
  
  If Not IsNull(.fields("CREADASRV").Value) Then tmpcreadasrv = .fields("CREADASRV").Value
  If Not IsNull(.fields("CREADACLI").Value) Then tmpcreadacli = .fields("CREADACLI").Value
  
  
  If Not IsNull(.fields("ID").Value) Then tmpid = .fields("ID").Value
  If Not IsNull(.fields("BBDDSRV").Value) Then tmpbbddsrv = .fields("BBDDSRV").Value
  If Not IsNull(.fields("BBDDCLI").Value) Then tmpbbddcli = .fields("BBDDCLI").Value
  
  .fields("CONSTRING") = " "
  
  .UpdateBatch adAffectAll
  
  .fields("CONSTRING") = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Trim(rc.fields("BBDDCLI").Value) & ";Data Source=" & Trim(rc.fields("IPCLI").Value)
  
  .UpdateBatch adAffectAll
  
  '//////////////////////////////////////////////
  ' Llama a la subrutina de crear la base de datos en el servidor
  '//////////////////////////////////////////////
  If (Not tmpcreadasrv) Or (Not tmpcreadacli) Then
  
    
    If crear_bbdd(tmpid, tmpbbddsrv, tmpbbddcli, tmpcreadasrv, tmpcreadacli) Then
        tmpbbddsrv = ""
        tmpbbddcli = ""
        Exit Sub
    End If
    
  End If

  If mbAddNewFlag Then
    .MoveLast              'va al nuevo registro
  End If
  
  End With

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  tmpbbddsrv = ""
  tmpbbddcli = ""
  
  DoEvents
  Call leer_configuracion
  DoEvents
  
  cbAgregar.SetFocus

  Exit Sub
UpdateErr:
  MsgBox Err.Description, vbInformation, titulo
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
End Sub


'---------------------------------------------------------------------------------------
' Procedure : crear_bbdd
' DateTime  : 10/10/2003 01:33
' Author    : Administrador
' Purpose   : Crea las bases de datos en el servidor, y marca el registro
'             correspondiente con creada = 1.
'   Si devuelve TRUE es que se ha producido un error
'---------------------------------------------------------------------------------------
'
Private Function crear_bbdd(id As Long, BBDDSRV As String, BBDDCLI As String, CreadaEnServer As Boolean, CreadaEnCliente As Boolean) As Boolean
Dim cm As New ADODB.Command
Dim var As Long
Dim cerrarSrvCnn As Boolean

On Error GoTo crear_bbdd_Error

lblstatus.Caption = "Creando Bases de datos ..."
DoEvents

EmpCnn.BeginTrans
        
    'crear base de datos en el servidor LOCAL
    If Not CreadaEnCliente Then
    
        cm.ActiveConnection = EmpCnn
        cm.CommandText = "CREATE DATABASE " & BBDDCLI
        cm.Execute var
    
        'Marcar como base de datos creada correctamente en la
        'Tabla empresas
        cm.ActiveConnection = EmpCnn
        cm.CommandText = "UPDATE EMPRESAS SET CREADACLI = 1 WHERE ID = " & id
        cm.Execute var
    
    End If
        
EmpCnn.CommitTrans
EmpCnn.BeginTrans
        
'preparar la conexión del servidor CENTRAL (remoto)
With SrvCnn
    If .State = 0 Then
    .Open strSrvCnn
        'cerrar al terminar
        cerrarSrvCnn = True
    End If

    .BeginTrans
    
    'crear base de datos en el servidor CENTRAL (remoto)
    If Not CreadaEnServer Then
    
        cm.ActiveConnection = SrvCnn
        cm.CommandText = "CREATE DATABASE " & BBDDSRV
        cm.Execute var
        
        'Marcar como base de datos creada correctamente en la
        'Tabla empresas
        cm.ActiveConnection = EmpCnn
        cm.CommandText = "UPDATE EMPRESAS SET CREADASRV = 1 WHERE ID = " & id
        cm.Execute var
            
    
    End If
    
    .CommitTrans
    'si la hemos abierto en esta función, volver a cerrarla
    If cerrarSrvCnn Then .Close
End With


lblstatus.Caption = "Operación Finalizada Correctamente"
DoEvents

On Error GoTo 0
Exit Function

crear_bbdd_Error:
    'cancelar transacciones
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crear_bbdd of Formulario frmMntEmp", vbExclamation, titulo
    On Error Resume Next
    EmpCnn.RollbackTrans
    SrvCnn.RollbackTrans
 
End Function



