VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntProv 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11100
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
   ScaleHeight     =   5985
   ScaleWidth      =   11100
   Begin PCGestion.miCombo cbSECTOR 
      Height          =   495
      Left            =   1215
      TabIndex        =   2
      Top             =   930
      Width           =   3750
      _extentx        =   6615
      _extenty        =   873
      font            =   "frmMntProv.frx":0000
   End
   Begin PCGestion.miText ioNOMBRE 
      Height          =   525
      Left            =   1215
      TabIndex        =   0
      Top             =   435
      Width           =   5340
      _extentx        =   9419
      _extenty        =   926
      font            =   "frmMntProv.frx":002C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioDIRECC 
      Height          =   525
      Left            =   6165
      TabIndex        =   3
      Top             =   915
      Width           =   4935
      _extentx        =   8705
      _extenty        =   926
      font            =   "frmMntProv.frx":0058
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioPROVIN 
      Height          =   525
      Left            =   6165
      TabIndex        =   5
      Top             =   1425
      Width           =   3630
      _extentx        =   6403
      _extenty        =   926
      font            =   "frmMntProv.frx":0084
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioLOCALI 
      Height          =   525
      Left            =   1215
      TabIndex        =   4
      Top             =   1425
      Width           =   3765
      _extentx        =   6641
      _extenty        =   926
      font            =   "frmMntProv.frx":00B0
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCIF 
      Height          =   525
      Left            =   9450
      TabIndex        =   1
      Top             =   435
      Width           =   1635
      _extentx        =   2566
      _extenty        =   926
      font            =   "frmMntProv.frx":00DC
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCODPOS 
      Height          =   525
      Left            =   10140
      TabIndex        =   6
      Top             =   1425
      Width           =   945
      _extentx        =   1667
      _extenty        =   926
      font            =   "frmMntProv.frx":0108
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioTELFNO 
      Height          =   525
      Left            =   1215
      TabIndex        =   7
      Top             =   1950
      Width           =   1485
      _extentx        =   2619
      _extenty        =   926
      font            =   "frmMntProv.frx":0134
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioFAX 
      Height          =   525
      Left            =   3330
      TabIndex        =   8
      Top             =   1950
      Width           =   1650
      _extentx        =   2910
      _extenty        =   926
      font            =   "frmMntProv.frx":0160
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioPERCON1 
      Height          =   525
      Left            =   6915
      TabIndex        =   9
      Top             =   1950
      Width           =   4185
      _extentx        =   7382
      _extenty        =   926
      font            =   "frmMntProv.frx":018C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioPERCON2 
      Height          =   525
      Left            =   1215
      TabIndex        =   10
      Top             =   2445
      Width           =   3765
      _extentx        =   6641
      _extenty        =   926
      font            =   "frmMntProv.frx":01B8
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioDCTO 
      Height          =   525
      Left            =   6915
      TabIndex        =   11
      Top             =   2453
      Width           =   675
      _extentx        =   1191
      _extenty        =   926
      font            =   "frmMntProv.frx":01E4
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioDCTOPP 
      Height          =   525
      Left            =   8475
      TabIndex        =   12
      Top             =   2453
      Width           =   675
      _extentx        =   1191
      _extenty        =   926
      font            =   "frmMntProv.frx":0210
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCCENTI 
      Height          =   525
      Left            =   6135
      TabIndex        =   13
      Top             =   3090
      Width           =   825
      _extentx        =   1455
      _extenty        =   926
      font            =   "frmMntProv.frx":023C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCCOFICI 
      Height          =   525
      Left            =   7800
      TabIndex        =   14
      Top             =   3090
      Width           =   675
      _extentx        =   1191
      _extenty        =   926
      font            =   "frmMntProv.frx":0268
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCCDC 
      Height          =   525
      Left            =   8775
      TabIndex        =   15
      Top             =   3090
      Width           =   495
      _extentx        =   873
      _extenty        =   926
      font            =   "frmMntProv.frx":0294
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioCCCUEN 
      Height          =   525
      Left            =   9780
      TabIndex        =   16
      Top             =   3083
      Width           =   1305
      _extentx        =   2302
      _extenty        =   926
      font            =   "frmMntProv.frx":02C0
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1050
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":02EC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":0318
      picn            =   "frmMntProv.frx":0336
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
      Left            =   5040
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1111
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":100A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":1036
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
      Left            =   8940
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":1054
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":1080
      picn            =   "frmMntProv.frx":109E
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
      Left            =   10020
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":1D72
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":1D9E
      picn            =   "frmMntProv.frx":1DBC
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":2AF4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":2B20
      picn            =   "frmMntProv.frx":2B3E
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
      TabIndex        =   43
      Top             =   5175
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":381A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":3846
      picn            =   "frmMntProv.frx":3864
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
      Left            =   2340
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5175
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":4140
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":416C
      picn            =   "frmMntProv.frx":418A
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
      Left            =   7920
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5175
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":49EA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":4A16
      picn            =   "frmMntProv.frx":4A34
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
      Left            =   8910
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":5310
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":533C
      picn            =   "frmMntProv.frx":535A
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
      Left            =   10020
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":5F2E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":5F5A
      picn            =   "frmMntProv.frx":5F78
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   30
      Top             =   4095
      Width           =   11055
      _extentx        =   19500
      _extenty        =   661
      caption         =   "-"
      fount           =   "frmMntProv.frx":6C54
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   15
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmMntProv.frx":6C82
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntProv.frx":6CAE
      picn            =   "frmMntProv.frx":6CCC
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miCombo ioCODBAN 
      Height          =   495
      Left            =   1185
      TabIndex        =   56
      Top             =   3075
      Width           =   3810
      _extentx        =   6720
      _extenty        =   873
      font            =   "frmMntProv.frx":7A04
   End
   Begin PCGestion.ucGrdBttn cmComentario 
      Height          =   375
      Left            =   2025
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3630
      Width           =   2985
      _extentx        =   5265
      _extenty        =   661
      caption         =   "Modificar Comentario"
      font            =   "frmMntProv.frx":7A30
      image           =   "frmMntProv.frx":7A5C
   End
   Begin PCGestion.bsGradientLabel lblExisteCom 
      Height          =   375
      Left            =   3435
      Top             =   30
      Visible         =   0   'False
      Width           =   2430
      _extentx        =   4286
      _extenty        =   661
      caption         =   "¡ Existe Comentario!"
      fount           =   "frmMntProv.frx":7A7A
      captioncolour   =   0
      colour1         =   8454143
      colour2         =   49152
      captionalignment=   1
   End
   Begin VB.Label BANCO 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO"
      Height          =   330
      Left            =   0
      TabIndex        =   57
      Top             =   3150
      Width           =   1110
   End
   Begin MSForms.CheckBox ioEXENTO 
      Height          =   435
      Left            =   9885
      TabIndex        =   55
      Top             =   2498
      Width           =   1155
      VariousPropertyBits=   746588183
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2037;767"
      Value           =   "0"
      Caption         =   "EXENTO"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox ioRE 
      Height          =   435
      Left            =   9255
      TabIndex        =   54
      Top             =   2498
      Width           =   660
      VariousPropertyBits=   746588183
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1164;767"
      Value           =   "0"
      Caption         =   "RE"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   1260
      TabIndex        =   53
      Top             =   45
      Width           =   870
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
      Left            =   8535
      TabIndex        =   52
      Top             =   45
      Width           =   2520
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
      Left            =   6930
      TabIndex        =   50
      Top             =   3690
      Width           =   1245
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
      Left            =   9810
      TabIndex        =   49
      Top             =   3690
      Width           =   1245
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   570
      TabIndex        =   48
      Top             =   3615
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
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CTA"
      Height          =   330
      Left            =   9225
      TabIndex        =   35
      Top             =   3180
      Width           =   525
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DC"
      Height          =   300
      Left            =   8505
      TabIndex        =   34
      Top             =   3195
      Width           =   270
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENTIDAD"
      Height          =   330
      Left            =   5205
      TabIndex        =   32
      Top             =   3180
      Width           =   915
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OFICINA"
      Height          =   300
      Left            =   6990
      TabIndex        =   33
      Top             =   3195
      Width           =   810
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO PP"
      Height          =   300
      Left            =   7605
      TabIndex        =   31
      Top             =   2565
      Width           =   870
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO"
      Height          =   330
      Left            =   6300
      TabIndex        =   30
      Top             =   2550
      Width           =   600
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONA CONT. 2"
      Height          =   600
      Left            =   90
      TabIndex        =   29
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONA CONT. 1"
      Height          =   300
      Left            =   5025
      TabIndex        =   28
      Top             =   2055
      Width           =   1845
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      Height          =   330
      Left            =   2835
      TabIndex        =   27
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TELEFONO"
      Height          =   300
      Left            =   120
      TabIndex        =   26
      Top             =   2055
      Width           =   1065
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CP"
      Height          =   330
      Left            =   9660
      TabIndex        =   25
      Top             =   1515
      Width           =   450
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CIF"
      Height          =   330
      Left            =   8820
      TabIndex        =   20
      Top             =   510
      Width           =   600
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOCALIDAD"
      Height          =   300
      Left            =   60
      TabIndex        =   23
      Top             =   1530
      Width           =   1125
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA"
      Height          =   330
      Left            =   4965
      TabIndex        =   24
      Top             =   1515
      Width           =   1155
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION"
      Height          =   330
      Left            =   4950
      TabIndex        =   22
      Top             =   1005
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SECTOR"
      Height          =   300
      Left            =   360
      TabIndex        =   21
      Top             =   1020
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   315
      TabIndex        =   17
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   8280
      TabIndex        =   37
      Top             =   3720
      Width           =   1485
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   6150
      TabIndex        =   18
      Top             =   75
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   5505
      TabIndex        =   36
      Top             =   3705
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      Height          =   300
      Left            =   345
      TabIndex        =   19
      Top             =   540
      Width           =   840
   End
End
Attribute VB_Name = "FrmMntProv"
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




Private Sub cmComentario_Click()

FrmInicio.Editor.carga "Comentario de Proveedores [" & ioNOMBRE.Text & "]", rc.fields("COMEN"), ""

End Sub


Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()
 
 If Not prime Then
 
 If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Proveedores. ¿Crear?", vbYesNo + vbQuestion, "Proveedores") = vbNo Then
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
  
  oSQL.AddTable "MAPROV"
  oSQL.AddOrderClause "CODIGO"
  oSQL.AddSimpleWhereClause "MBAJA", 0
  
  If TipoServer = 1 Then
    rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
  Else
    rc.Open oSQL.SQL, locCnn, adOpenKeyset, adLockOptimistic
  End If


''

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
        'no dejar pasar con el campo en blanco
        .PermitirBlanco = False
        .DataField = "NOMBRE"
        .LongMaxima = 30
  End With
  
  With ioCIF
  Set .DataSource = rc
        'no dejar pasar con el campo en blanco
        .PermitirBlanco = False
        .DataField = "CIF"
        .LongMaxima = 12
  End With
  
  With ioDIRECC
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DIRECC"
        .LongMaxima = 20
  End With
  
  With ioPROVIN
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PROVIN"
        .LongMaxima = 15
  End With
  
  With ioLOCALI
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "LOCALI"
        .LongMaxima = 10
  End With
  
  With ioCODPOS
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CODPOS"
        .LongMaxima = 5
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioTELFNO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "TELFNO"
        .LongMaxima = 9
        .SoloNumeros = True
        .Alineacion = 1
  End With
  
  With ioPERCON1
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PERCON1"
        .LongMaxima = 20
  End With
  
 With ioPERCON2
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "PERCON2"
        .LongMaxima = 20
  End With
    
  With ioFAX
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "FAX"
        .LongMaxima = 9
         .SoloNumeros = True
         .Alineacion = 1
  End With
    
  With ioDCTO
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DCTO"
        .LongMaxima = 2
        .dspFormat = "00"
         .SoloNumeros = True
         .Alineacion = 1
  End With
  
  With ioDCTOPP
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DCTOPP"
        .LongMaxima = 2
        .dspFormat = "00"
         .SoloNumeros = True
         .Alineacion = 1
  End With
  
 With ioRE
  Set .DataSource = rc
        .DataField = "RE"
  End With
  
  With ioEXENTO
  Set .DataSource = rc
        .DataField = "EXENTO"
  End With
  
  'datos bancarios:
  With ioCCENTI
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CCENTI"
        .LongMaxima = 4
         .SoloNumeros = True
         .Alineacion = 1
  End With

  With ioCCOFICI
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CCOFICI"
        .LongMaxima = 4
         .SoloNumeros = True
         .Alineacion = 1
  End With
    
  With ioCCDC
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CCDC"
        .LongMaxima = 2
         .SoloNumeros = True
         .Alineacion = 1
  End With

  With ioCCCUEN
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "CCCUEN"
        .LongMaxima = 10
         .SoloNumeros = True
         .Alineacion = 1
  End With

  
  'Cargar el micombo sectores
  With cbSECTOR
    .ConexionString = locCnn
    .SQLString = "SELECT CODST, SECTOR FROM SECTORES WHERE MBAJA = 0 ORDER BY CODST"
    .LenCodigo = 2
    .DataField = "SECTOR"
    .carga
    .CodigoWidth = 500
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
      cbcerrar_Click
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
  ' End With
   
Set oSQL = Nothing
Set nif = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set FrmMntProv = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

With frmFlexProv
    
    Set .miosql = oSQL
    
    With .fg
             Set frmFlexProv.miRc = rc
    End With
    
    .Caption = "Proveedores ..."
    
    End With
    
'////////////////////////////////////////////////////////
' Des - Enlazar controles:
'////////////////////////////////////////////////////////


  With ioCODIGO
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioNOMBRE
    Set .DataSource = Nothing
        .DataField = ""
End With
  
  With ioCIF
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioDIRECC
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioPROVIN
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioLOCALI
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioCODPOS
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioTELFNO
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioPERCON1
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
 With ioPERCON2
    Set .DataSource = Nothing
        .DataField = ""
  End With
    
  With ioFAX
    Set .DataSource = Nothing
        .DataField = ""
  End With
    
  With ioDCTO
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioDCTOPP
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
 With ioRE
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  With ioEXENTO
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
  'datos bancarios:
  With ioCCENTI
    Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCCOFICI
    Set .DataSource = Nothing
        .DataField = ""
  End With
    
  With ioCCDC
    Set .DataSource = Nothing
        .DataField = ""
  End With

  With ioCCCUEN
    Set .DataSource = Nothing
        .DataField = ""
  End With

  
  'Cargar el micombo sectores
  With cbSECTOR
    Set .DataSource = Nothing
        .DataField = ""
  End With
  
    With ioCODBAN
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

    
    frmFlexProv.Show 1
    
    
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioNOMBRE
  Set .DataSource = rc
        .DataField = "NOMBRE"
  End With
  
  With ioCIF
  Set .DataSource = rc
        .DataField = "CIF"
  End With
  
  With ioDIRECC
  Set .DataSource = rc
        .DataField = "DIRECC"
  End With
  
  With ioPROVIN
  Set .DataSource = rc
        .DataField = "PROVIN"
  End With
  
  With ioLOCALI
  Set .DataSource = rc
        .DataField = "LOCALI"
  End With
  
  With ioCODPOS
  Set .DataSource = rc
        .DataField = "CODPOS"
  End With
  
  With ioTELFNO
  Set .DataSource = rc
        .DataField = "TELFNO"
  End With
  
  With ioPERCON1
  Set .DataSource = rc
        .DataField = "PERCON1"
  End With
  
 With ioPERCON2
  Set .DataSource = rc
        .DataField = "PERCON2"
  End With
    
  With ioFAX
  Set .DataSource = rc
        .DataField = "FAX"
  End With
    
  With ioDCTO
  Set .DataSource = rc
        .DataField = "DCTO"
  End With
  
  With ioDCTOPP
  Set .DataSource = rc
      .DataField = "DCTOPP"
  End With
  
 With ioRE
  Set .DataSource = rc
      .DataField = "RE"
  End With
  
  With ioEXENTO
  Set .DataSource = rc
      .DataField = "EXENTO"
  End With
  
  'datos bancarios:
  With ioCCENTI
  Set .DataSource = rc
        .DataField = "CCENTI"
  End With

  With ioCCOFICI
  Set .DataSource = rc
        .DataField = "CCOFICI"
  End With
    
  With ioCCDC
  Set .DataSource = rc
        .DataField = "CCDC"
  End With

  With ioCCCUEN
  Set .DataSource = rc
        .DataField = "CCCUEN"
  End With

  
  'Cargar el micombo sectores
  With cbSECTOR
    .DataField = "SECTOR"
    Set .DataSource = rc
  End With
  
    With ioCODBAN
    .DataField = "CODBAN"
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


Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub

Private Sub ioCIF_Validate(Cancel As Boolean)
'Call comprueba_DNI(ioCIF.Text, ioCIF)


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
End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
  
    If Not IsNull(rc.fields("COMEN")) Then
    lblExisteCom.Visible = True
     Else
    lblExisteCom.Visible = False
    End If
  
  End If
  
  
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
  
        If ioDCTO.Text = "" Then ioDCTO.Text = "0"
        If ioDCTOPP.Text = "" Then ioDCTOPP.Text = "0"
        If ioCODPOS.Text = "" Then ioCODPOS.Text = "0"
        If ioCCENTI.Text = "" Then ioCCENTI.Text = "0"
        If ioCCOFICI.Text = "" Then ioCCOFICI.Text = "0"
        If ioCCOFICI.Text = "" Then ioCCOFICI.Text = "0"
        If ioCCDC.Text = "" Then ioCCDC.Text = "0"
        If ioCCCUEN.Text = "" Then ioCCCUEN.Text = "0"
  
        If ioCIF.Text = "" Then
            lblstatus.Caption = "CIF no puede estar en blanco"
            ioCIF.SetFocus
            bCancel = True
        End If
        
        If ioNOMBRE.Text = "" Then
            lblstatus.Caption = "NOMBRE no puede estar en blanco"
            ioNOMBRE.SetFocus
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
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from MAPROV")

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

  'Para MySQL darle los valores de fecha
  If TipoServer = 2 And mbAddNewFlag Then
    rc.fields("FALTA") = Now
    rc.fields("FMODI") = Now
  ElseIf TipoServer = 2 And mbEditFlag = False Then
    rc.fields("FMODI") = Now
  End If

  rc.Update 'Batch adAffectAll

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
  If Err.Number = -2147217887 Then Exit Sub
  MsgBox Err.Description
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
  
  cbSECTOR.Locked = bVal
  ioCODBAN.Locked = bVal
  cmComentario.Enabled = Not bVal
  
End Sub
