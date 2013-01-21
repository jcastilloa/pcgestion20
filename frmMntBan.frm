VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntBan 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3405
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
      MICON           =   "frmMntBan.frx":0000
      PICN            =   "frmMntBan.frx":001C
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
      Top             =   2985
      Width           =   9540
      _ExtentX        =   16828
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
   Begin PCGestion.miText ioNOMBRE 
      Height          =   525
      Left            =   1155
      TabIndex        =   0
      Top             =   510
      Width           =   3300
      _ExtentX        =   5821
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3405
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
      MICON           =   "frmMntBan.frx":0CEE
      PICN            =   "frmMntBan.frx":0D0A
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
      Left            =   4410
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3405
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
      MICON           =   "frmMntBan.frx":1A40
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
      Left            =   7455
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3405
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
      MICON           =   "frmMntBan.frx":1A5C
      PICN            =   "frmMntBan.frx":1A78
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
      Left            =   8520
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3405
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
      MICON           =   "frmMntBan.frx":274A
      PICN            =   "frmMntBan.frx":2766
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":349C
      PICN            =   "frmMntBan.frx":34B8
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
      TabIndex        =   24
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":4192
      PICN            =   "frmMntBan.frx":41AE
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":4A88
      PICN            =   "frmMntBan.frx":4AA4
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
      Left            =   6435
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":5302
      PICN            =   "frmMntBan.frx":531E
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
      Left            =   7410
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":5BF8
      PICN            =   "frmMntBan.frx":5C14
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
      Left            =   8520
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4065
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
      MICON           =   "frmMntBan.frx":67E6
      PICN            =   "frmMntBan.frx":6802
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioDOMICILIO 
      Height          =   525
      Left            =   5520
      TabIndex        =   1
      Top             =   525
      Width           =   4065
      _ExtentX        =   7170
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
      Left            =   1155
      TabIndex        =   2
      Top             =   1020
      Width           =   3600
      _ExtentX        =   6350
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
      Left            =   1155
      TabIndex        =   4
      Top             =   1500
      Width           =   915
      _ExtentX        =   1614
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
      Left            =   5850
      TabIndex        =   3
      Top             =   1020
      Width           =   3735
      _ExtentX        =   6588
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
   Begin PCGestion.miText ioTELEFONO 
      Height          =   525
      Left            =   4830
      TabIndex        =   6
      Top             =   1515
      Width           =   1665
      _ExtentX        =   2937
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
      Left            =   2505
      TabIndex        =   5
      Top             =   1515
      Width           =   1665
      _ExtentX        =   2937
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
   Begin PCGestion.miText ioPERCON 
      Height          =   525
      Left            =   1155
      TabIndex        =   8
      Top             =   2085
      Width           =   4500
      _ExtentX        =   7990
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
   Begin PCGestion.miText ioTELEFPCON 
      Height          =   525
      Left            =   7770
      TabIndex        =   7
      Top             =   1515
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
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TLF. CONTACTO"
      Height          =   585
      Left            =   6330
      TabIndex        =   37
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONA CONTACTO"
      Height          =   645
      Left            =   45
      TabIndex        =   36
      Top             =   2070
      Width           =   1110
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO POSTAL."
      Height          =   600
      Left            =   180
      TabIndex        =   35
      Top             =   1455
      Width           =   900
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA"
      Height          =   300
      Left            =   4635
      TabIndex        =   34
      Top             =   1095
      Width           =   1200
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TLF 1"
      Height          =   360
      Left            =   4230
      TabIndex        =   33
      Top             =   1590
      Width           =   585
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      Height          =   360
      Left            =   2040
      TabIndex        =   32
      Top             =   1590
      Width           =   435
   End
   Begin VB.Label LOCALIDAD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCALIDAD"
      Height          =   300
      Left            =   30
      TabIndex        =   31
      Top             =   1065
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DOMICILIO"
      Height          =   360
      Left            =   4440
      TabIndex        =   30
      Top             =   570
      Width           =   1050
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   555
      TabIndex        =   29
      Top             =   2565
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
      Left            =   8325
      TabIndex        =   17
      Top             =   2595
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
      Left            =   4410
      TabIndex        =   16
      Top             =   2595
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
      Left            =   7125
      TabIndex        =   15
      Top             =   90
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
      Left            =   1200
      TabIndex        =   14
      Top             =   75
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   285
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   6975
      TabIndex        =   12
      Top             =   2625
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   4905
      TabIndex        =   11
      Top             =   105
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   3000
      TabIndex        =   10
      Top             =   2610
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   360
      Left            =   195
      TabIndex        =   9
      Top             =   585
      Width           =   930
   End
End
Attribute VB_Name = "frmMntBan"
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

Private Sub CheckBox1_Click()

End Sub

Private Sub Form_Activate()

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Bancos. ¿Crear?", vbYesNo + vbQuestion, "Bancos") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
        Call cmdFirst_Click
        
  End If
  

prime = True
End If
    
End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
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
  oSQL.AddTable "BANCOS"
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
        .LongMaxima = 50
  End With
  
  With ioDOMICILIO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DOMICILIO"
        .LongMaxima = 50
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
        .LongMaxima = 25
  End With
  
  With ioCODPOS
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CODPOS"
        .LongMaxima = 5
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
    
  With ioTELEFONO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELEFONO"
        .LongMaxima = 12
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioFAX
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "FAX"
        .LongMaxima = 12
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
    
  With ioPERCON
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PERCON"
        .LongMaxima = 30
  End With
  
  With ioTELEFPCON
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELEFPCON"
        .LongMaxima = 12
        .SoloNumeros = True
        .Alineacion = 1
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

 '  With locCnn
  '  If .State <> 0 Then .Close
 '  End With

Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmMntBan = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

'If rc.EditMode = adEditNone Then

With frmFlexSimple

    .Caption = "Bancos ..."
        
    With .fg
            Set .DataSource = rc
            .ColFormat(1) = "0000"
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
  
        If ioNOMBRE.Text = "" Then
            lblstatus.Caption = "Nombre no puede estar en blanco"
            ioNOMBRE.SetFocus
            ioNOMBRE.CancelarValidacion
            bCancel = True
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
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from BANCOS")
    
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
End Sub
