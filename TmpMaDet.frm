VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form TmpMaDet 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pedidos a Proveedores"
   ClientHeight    =   8025
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10785
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.ucGrdBttn ucGrdBttn1 
      Height          =   435
      Left            =   3600
      TabIndex        =   41
      Top             =   7350
      Width           =   1185
      _extentx        =   2090
      _extenty        =   767
      caption         =   "Etiquetas"
      forecolor       =   16711680
      gradientcolor1  =   16777215
      gradientcolor2  =   16777215
      angle           =   274
      font            =   "TmpMaDet.frx":0000
      image           =   "TmpMaDet.frx":002C
   End
   Begin TabDlg.SSTab Tab2 
      Height          =   4020
      Left            =   30
      TabIndex        =   27
      Top             =   1485
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   7091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Artículos"
      TabPicture(0)   =   "TmpMaDet.frx":004A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Totales"
      TabPicture(1)   =   "TmpMaDet.frx":0066
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   3600
         Left            =   60
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   10620
         _cx             =   18732
         _cy             =   6350
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"TmpMaDet.frx":0082
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
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6540
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":0160
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":018C
      picn            =   "TmpMaDet.frx":01AA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6540
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":0E7E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":0EAA
      picn            =   "TmpMaDet.frx":0EC8
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
      Left            =   8640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6540
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":1C00
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":1C2C
      picn            =   "TmpMaDet.frx":1C4A
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
      Left            =   9705
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6540
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":291E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":294A
      picn            =   "TmpMaDet.frx":2968
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
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":36A0
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":36CC
      picn            =   "TmpMaDet.frx":36EA
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
      Left            =   1125
      TabIndex        =   15
      Top             =   7200
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":43C6
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":43F2
      picn            =   "TmpMaDet.frx":4410
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
      Left            =   2355
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7200
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":4CEC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":4D18
      picn            =   "TmpMaDet.frx":4D36
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
      Left            =   7620
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7200
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":5596
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":55C2
      picn            =   "TmpMaDet.frx":55E0
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
      Left            =   8595
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":5EBC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":5EE8
      picn            =   "TmpMaDet.frx":5F06
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
      Left            =   9705
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":6ADA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":6B06
      picn            =   "TmpMaDet.frx":6B24
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   4905
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6540
      Width           =   1185
      _extentx        =   2090
      _extenty        =   1111
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":7800
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":782C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtAgregar 
      Height          =   585
      Left            =   30
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5550
      Width           =   735
      _extentx        =   1296
      _extenty        =   1032
      btype           =   9
      tx              =   "Agregar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":784A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":7876
      picn            =   "TmpMaDet.frx":7894
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtActualizar 
      Height          =   585
      Left            =   780
      TabIndex        =   22
      Top             =   5550
      Width           =   870
      _extentx        =   1535
      _extenty        =   1032
      btype           =   9
      tx              =   "Actualizar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":7AA0
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":7ACC
      picn            =   "TmpMaDet.frx":7AEA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtEdicion 
      Height          =   585
      Left            =   1680
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5550
      Width           =   720
      _extentx        =   1270
      _extenty        =   1032
      btype           =   9
      tx              =   "Edicion"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":7CFE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":7D2A
      picn            =   "TmpMaDet.frx":7D48
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtCancelar 
      Height          =   585
      Left            =   8445
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5550
      Width           =   810
      _extentx        =   1429
      _extenty        =   1032
      btype           =   9
      tx              =   "Cancelar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":7F48
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":7F74
      picn            =   "TmpMaDet.frx":7F92
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtEliminar 
      Height          =   585
      Left            =   9285
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5550
      Width           =   720
      _extentx        =   1270
      _extenty        =   1032
      btype           =   9
      tx              =   "Eliminar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":816E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":819A
      picn            =   "TmpMaDet.frx":81B8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton dtCerrar 
      Height          =   585
      Left            =   10020
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5550
      Width           =   720
      _extentx        =   1270
      _extenty        =   1032
      btype           =   9
      tx              =   "Cerrar"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":8384
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":83B0
      picn            =   "TmpMaDet.frx":83CE
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   30
      Top             =   6150
      Width           =   10725
      _extentx        =   16854
      _extenty        =   661
      caption         =   ""
      fount           =   "TmpMaDet.frx":85D6
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel lblStatusD 
      Height          =   375
      Left            =   2430
      Top             =   5550
      Width           =   5985
      _extentx        =   10557
      _extenty        =   661
      caption         =   ""
      fount           =   "TmpMaDet.frx":8604
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.miText ioNUMERO 
      Height          =   525
      Left            =   825
      TabIndex        =   0
      Top             =   30
      Width           =   960
      _extentx        =   1693
      _extenty        =   926
      font            =   "TmpMaDet.frx":8632
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin PCGestion.miText ioFECHA 
      Height          =   525
      Left            =   2340
      TabIndex        =   1
      Top             =   30
      Width           =   1230
      _extentx        =   2170
      _extenty        =   926
      font            =   "TmpMaDet.frx":865E
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin PCGestion.miText ioTRNSPORTI 
      Height          =   525
      Left            =   8685
      TabIndex        =   9
      Top             =   975
      Width           =   2115
      _extentx        =   3731
      _extenty        =   926
      font            =   "TmpMaDet.frx":868A
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin PCGestion.miCombo cbCODPROV 
      Height          =   540
      Left            =   4605
      TabIndex        =   2
      Top             =   45
      Width           =   6150
      _extentx        =   10848
      _extenty        =   953
      font            =   "TmpMaDet.frx":86B6
   End
   Begin PCGestion.miCombo cBFPAGO 
      Height          =   480
      Left            =   810
      TabIndex        =   3
      Top             =   525
      Width           =   3840
      _extentx        =   6773
      _extenty        =   847
      font            =   "TmpMaDet.frx":86E2
   End
   Begin PCGestion.miCombo cbPLAZOE 
      Height          =   570
      Left            =   5655
      TabIndex        =   4
      Top             =   525
      Width           =   3645
      _extentx        =   6429
      _extenty        =   1005
      font            =   "TmpMaDet.frx":870E
   End
   Begin PCGestion.chameleonButton cbNuevoArticulo 
      Height          =   525
      Left            =   3495
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6615
      Width           =   1380
      _extentx        =   2434
      _extenty        =   926
      btype           =   9
      tx              =   "Nuevo Artículo"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":873A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":8766
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbSeleccionaArticulo 
      Height          =   525
      Left            =   6135
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1395
      _extentx        =   2461
      _extenty        =   926
      btype           =   9
      tx              =   "Buscar Artículo"
      enab            =   -1  'True
      font            =   "TmpMaDet.frx":8784
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "TmpMaDet.frx":87B0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALM 
      Height          =   480
      Left            =   810
      TabIndex        =   6
      Top             =   990
      Width           =   3285
      _extentx        =   5794
      _extenty        =   847
      font            =   "TmpMaDet.frx":87CE
   End
   Begin PCGestion.ucGrdBttn cmEntrada 
      Height          =   435
      Left            =   6150
      TabIndex        =   42
      Top             =   7350
      Width           =   1185
      _extentx        =   2090
      _extenty        =   767
      caption         =   "Entrada"
      forecolor       =   16711680
      gradientcolor1  =   16777215
      gradientcolor2  =   16777215
      angle           =   274
      font            =   "TmpMaDet.frx":87FA
      image           =   "TmpMaDet.frx":8826
   End
   Begin PCGestion.miText ioPORTES 
      Height          =   480
      Left            =   4830
      TabIndex        =   7
      Top             =   1005
      Width           =   1155
      _extentx        =   2037
      _extenty        =   847
      font            =   "TmpMaDet.frx":8844
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin PCGestion.miText ioGASTOS 
      Height          =   525
      Left            =   6720
      TabIndex        =   8
      Top             =   990
      Width           =   1155
      _extentx        =   2037
      _extenty        =   926
      font            =   "TmpMaDet.frx":8870
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin PCGestion.miText ioDCTOPP 
      Height          =   525
      Left            =   10005
      TabIndex        =   5
      Top             =   525
      Width           =   795
      _extentx        =   1402
      _extenty        =   926
      font            =   "TmpMaDet.frx":889C
      dspformat       =   ""
      enabled         =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALMAC."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   40
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PLAZO    ENT."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4530
      TabIndex        =   36
      Top             =   495
      Width           =   1170
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FORMA PAGO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   45
      TabIndex        =   35
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PORTES"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4065
      TabIndex        =   34
      Top             =   1095
      Width           =   765
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GASTOS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   1095
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSP."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7875
      TabIndex        =   32
      Top             =   1110
      Width           =   825
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3555
      TabIndex        =   31
      Top             =   135
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO PP"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9360
      TabIndex        =   30
      Top             =   480
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1770
      TabIndex        =   29
      Top             =   135
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   28
      Top             =   135
      Width           =   765
   End
End
Attribute VB_Name = "TmpMaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module     : TmpMaDet
' DateTime : 10/10/2003 17:40
' Author     :  José Castillo
' Purpose   : Plantilla para Maestro/Detalle.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Notas:
'---------------------------------------------------------------------------------------
' Para esta versión se utiliza un recordset jerarquico.
' Primario (Maestro) :    adoPrimaryRS
' Secundario (Detalle) :  adoDetalleRS

' poner el código de validación de los campos en detalle en:
'
' * fg_ValidateEdit (evento del grid)
'
' y si fuera necesario, también poner código en:
' * adoDetalleRS_WillChangeRecord (evento del recordset)

' poner replaces (.EditText = Replace(.EditText, ",", ".")) para aquellos campos
' que sean con decimales. Tambien para poner solo_numericos:

' * fg_ChangeEdit
'
' Poner los DataSource, DataFields y miCombos en
' Private Sub Asigna_Campos()

' Los formatos y ColCombos para el grid poner en
' Private Sub Asigna_Grid()
'---------------------------------------------------------------------------------------




'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  'cambiar el codigo de usuario (DEPENDIENTE)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoDetalleRS As Recordset
Attribute adoDetalleRS.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim editagrid As Boolean

Dim lblValue As String
Dim lblOriginalValue As String
Dim Casignados As Boolean
Dim Gasignados As Boolean

'para almacenar el IDTEM de la temporada de trabajo.
'Dim tmptemporada As Long

Dim oSQL As New clsSmartSQL


Dim tmptempor As String
'Dim tmpcodigo As String
Dim tmptalla As String
Dim tmpcolor As String


Private Sub cbNuevoArticulo_Click()

With FrmMntArt
    .NumeroPedido = ioNUMERO.Text
    .add_en_detalle = True
    Set .rc_detalle = adoDetalleRS
    .Show
    DoEvents
End With

End Sub

Private Sub UpdateLabels()
    'On error Resume Next
    Dim fld As String
    
  '  With fg
  '
  '  fld = .TextMatrix(0, .Col)
  '
  '  End With
  '
  '  If fld = "" Then Exit Sub
    
    'grdDatagrid.Text
 
'    With adoDetalleRS(fld)
'        If Not IsNull(.Value) Then lblValue = .Value
        'If Not IsNull(.OriginalValue) Then lblOriginalValue = .OriginalValue
'    End With
    
    'grdDatagrid.Text
    'If lblValue <> lblOriginalValue Then
      
'    With fg
'
'     If .Text <> lblOriginalValue Then
'
'        .BackColorSel = vbBlack
'        .CellFontBold = True
'        .CellForeColor = vbWhite
'        .CellBackColor = vbBlue

'        'lblValue.ForeColor = vbRed
''    Else
 '       .BackColorSel = &H8000000D 'default selection color
 '                                                 'lblValue.ForeColor = vbBlack
 '   End If
 '
 '   End With
    
End Sub



Private Sub adoDetalleRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
  
  'If adoDetalleRS.Fields("DENOMINACION") = "" Then
  '  bCancel = True
  '  lblStatusD.Caption = "Denominación no puede estar en blanco"
  '  DoEvents
  'End If
  
  'adoPrimaryRS.UpdateBatch adAffectAllChapters
  'Call UpdateLabels
  
  End Select


  If bCancel Then adStatus = adStatusCancel
End Sub




'---------------------------------------------------------------------------------------
' Procedure : Asigna_Campos
' DateTime  : 22/10/2003 21:38
' Author    : Administrador
' Purpose   : Configura los Datafield y Datasources de los campos
'             asi como los miCombos si fueran necesarios
'---------------------------------------------------------------------------------------
Private Sub Asigna_Campos()
   
  With ioNUMERO
    .LongMaxima = 9
    Set .DataSource = adoPrimaryRS
    .DataField = "NUMERO"
    .SoloNumeros = True
    .Alineacion = 1
    .Locked = True  'que entre como bloqueado y solo desbloquear
                    'para añadir un nuevo registro
  End With
  
  With ioFECHA
    .LongMaxima = 10
    Set .DataSource = adoPrimaryRS
    .DataField = "FECHA"
  End With
  
  With ioTRNSPORTI
    .LongMaxima = 40
    Set .DataSource = adoPrimaryRS
    .DataField = "TRNSPORTI"
  End With
  
  With ioDCTOPP
  .Alineacion = 1
  .dspFormat = "00.00"
  .SoloNumeros = True
   ' Set .DataSource = adoPrimaryRS
    '.DataField = "DCTOPP"
 '   .displayformat = "00.00 %"
  '  .Format = "##.##"
  End With
  
  'With ioDCTO
 ' .Alineacion = 1
 ' .dspFormat = "00.00"
 ' .SoloNumeros = True
   ' Set .DataSource = adoPrimaryRS
    '.DataField = "DCTO"
    '.displayformat = "00.00 %"
    '.Format = "##.##"
  'End With
    
  With ioGASTOS
  .Alineacion = 1
  .dspFormat = "Currency"
  .SoloNumeros = True
   ' Set .DataSource = adoPrimaryRS
  '  .DataField = "GASTOS"
   ' .displayformat = "0000.00 €"
    '.Format = "####.##"
  End With
   
    With ioPORTES
  .Alineacion = 1
  .dspFormat = "Currency"
  .SoloNumeros = True
   ' Set .DataSource = adoPrimaryRS
   ' .DataField = "PORTES"
   ' .displayformat = "0000.00 €"
   ' .Format = "####.##"
  End With
    
With cbCODPROV
    .ConexionString = locCnnSP
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .DataField = "CODPROV"
    .carga
    Set .DataSource = adoPrimaryRS
    .CodigoWidth = 800
End With

With cbFPAGO
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FPAGO WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "FPAGO"
    .carga
    Set .DataSource = adoPrimaryRS
End With

With cbCODALM
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "CODALM"
    .carga
    Set .DataSource = adoPrimaryRS
End With

With cbPLAZOE
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM PLAZOE WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "PLAZOE"
    .carga
    Set .DataSource = adoPrimaryRS
    
End With

  mbDataChanged = False
  
  Casignados = True
  
  If Not adoPrimaryRS.EOF Then
  If Not IsNull(adoPrimaryRS("ChildCMD").UnderlyingValue) Then
    Call Asigna_Grid
  End If
  End If
 
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Asigna_Grid
' DateTime  : 22/10/2003 21:45
' Author    : Administrador
' Purpose   : Asigna los datos a adoDetalleRs y al grid
'---------------------------------------------------------------------------------------
Private Sub Re_Asigna_Grid()
Dim tmpcodart As Integer
Dim tmpuds As Double
Dim totuds As Double
Dim fila As Long


'1 NUMERO
'2 linea
'3 CODART
'4 TEMPOR
'5 PRECOM
'6 CODTALLA
'7 CODCOL
'8 unidades
'9 DCTO1
'10 DCTO2
'11 IVA
'12 RE
'13 FMODI

 fg.Clear
 fg.Rows = 1
 
 If (adoDetalleRS Is Nothing) Then
  
    Exit Sub
    
 ElseIf (adoDetalleRS.EOF And adoDetalleRS.BOF) Then
 
    Exit Sub
 
 End If
  
 With fg
    .Rows = 1
    .AllowBigSelection = False
    .AllowSelection = True
    .SelectionMode = flexSelectionByRow
    .AllowUserResizing = flexResizeColumns
    .HighLight = flexHighlightWithFocus
    .FocusRect = flexFocusSolid
 End With
 
 With adoDetalleRS
 
    fg.Cols = .Fields.Count

    If Not .BOF Then .MoveFirst
    fg.Redraw = flexRDNone
    
    'If tmpcodart = 0 Then 'asignar el codigo si es la primera vez
    tmpcodart = .Fields("CODART")
   ' End If
            
        
    fila = .AbsolutePosition
     
    Do Until .EOF
    
        
    fg.Rows = fg.Rows + 1
    
   ' If Not IsNull(.Fields("NUMERO")) Then ' fg.TextMatrix(.AbsolutePosition, 1) = .Fields("NUMERO")
    
   
    If tmpcodart = .Fields("CODART") Then  'romper por codigo
    
        tmpuds = tmpuds + .Fields("UNIDADES")  'sumar unidades
    
        
   
        If Not IsNull(.Fields("CODART")) And Not IsNull(.Fields("TEMPOR")) Then
    
                   
            tmpcodart = .Fields("CODART") 'asignar nuevo codigo
            fg.TextMatrix(fila, 2) = Format(.Fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & .Fields("CODART") & " AND TEMPOR = " & .Fields("TEMPOR"))
            fg.TextMatrix(fila, 3) = .Fields("TEMPOR")
      
        End If
    
        If Not IsNull(.Fields("LINEA")) Then fg.TextMatrix(fila, 1) = .Fields("LINEA")
       
        If Not IsNull(.Fields("PRECOM")) Then fg.TextMatrix(fila, 4) = .Fields("PRECOM")
    
        If Not IsNull(.Fields("CODTALLA")) Then fg.TextMatrix(fila, 5) = .Fields("CODTALLA")
    
        If Not IsNull(.Fields("CODCOL")) Then
            fg.TextMatrix(fila, 6) = .Fields("CODCOL")
            fg.Col = 6
            fg.Row = fila
            fg.CellBackColor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & .Fields("CODCOL"))
        End If
    
        If Not IsNull(.Fields("unidades")) Then fg.TextMatrix(fila, 7) = .Fields("unidades")
    
        If Not IsNull(.Fields("DCTO1")) Then fg.TextMatrix(fila, 8) = .Fields("DCTO1")
    
        If Not IsNull(.Fields("DCTO2")) Then fg.TextMatrix(fila, 9) = .Fields("DCTO2")
    
        If Not IsNull(.Fields("IVA")) Then fg.TextMatrix(fila, 10) = .Fields("IVA")
    
        If Not IsNull(.Fields("RE")) Then fg.TextMatrix(fila, 11) = .Fields("RE")
    
        If Not IsNull(.Fields("FMODI")) Then fg.TextMatrix(fila, 12) = .Fields("FMODI")
    
       .MoveNext
        fila = fila + 1

        
    ElseIf (tmpcodart <> .Fields("CODART") Or .EOF) Then
    
    'insertar fila de subtotal
    Call imprime_Subtotal(tmpuds, fila, "Total Uds. por Artículo:", False)
    tmpcodart = .Fields("CODART")
    fila = fila + 1
    totuds = totuds + tmpuds
    tmpuds = 0
    
    
    End If
    
    Loop
        
    fg.Rows = fg.Rows + 1
    totuds = totuds + tmpuds
    'insertar fila de subtotal
    
    Call imprime_Subtotal(tmpuds, fila, "Total Uds. por Artículo:", False)
    
    'añadir una linea al principio
    fg.AddItem "", 1
    'insertar el total general
    Call imprime_Subtotal(totuds, 1, "Total Uds. General:", True)
        
 End With
 
 
 With fg
    
   ' .Subtotal flexSTClear
    .TextMatrix(0, 1) = "Línea"
    .TextMatrix(0, 2) = "Artículo"
    .TextMatrix(0, 3) = "Temporada"
    .TextMatrix(0, 4) = "Precio Com"
    .TextMatrix(0, 5) = "Talla"
    .TextMatrix(0, 6) = "Color"
    .TextMatrix(0, 7) = "Uds."
    .TextMatrix(0, 8) = "Dcto. 1"
    .TextMatrix(0, 9) = "Dcto. 2"
    .TextMatrix(0, 10) = "IVA"
    .TextMatrix(0, 11) = "RE"
    .TextMatrix(0, 12) = "Ul.Modif."
    
    '.Subtotal flexSTSum, 0, 7, , vbBlue, vbWhite, True, "Total Uds. General:"
    '.SubtotalPosition = flexSTAbove
    '.Subtotal flexSTSum, 2, 7, , 16744576, , True, "Total por Articulo:"

    
    .ColFormat(1) = "00000"
    
    .ColComboList(3) = tmptempor
    .ColComboList(5) = tmptalla
    .ColComboList(6) = tmpcolor
    
    .ColFormat(4) = "Currency"
    .AutoSize 1, .Cols - 1
    
  End With
  
  With adoDetalleRS
  
    If Not .BOF Then .MoveFirst
  
  End With

  fg.Redraw = True
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : imprime_Subtotal
' Fecha/Hora  : 16/11/2003 21:08
' Autor       : JCASTILLO
' Propósito   : Imprimir la linea de subtotales en el flexgrid
'
'---------------------------------------------------------------------------------------
Private Sub imprime_Subtotal(unidades As Double, fila As Long, texto As String, invertido As Boolean)
Dim tmpcol As Byte

   'On Error GoTo imprime_Subtotal_Error
    
    fg.Row = fila
    
        For tmpcol = 1 To fg.Cols - 1
            fg.Col = tmpcol
            
            If Not invertido Then
            
                fg.CellBackColor = vbWhite
                fg.CellForeColor = vbBlue
            
            Else
            
                fg.CellBackColor = vbBlue
                fg.CellForeColor = vbWhite
            
            End If
            
            fg.CellFontBold = True
        Next
    
    fg.TextMatrix(fila, 1) = texto
    fg.TextMatrix(fila, 2) = fg.TextMatrix(fila - 1, 2) 'poner la descripción del articulo

    fg.TextMatrix(fila, 7) = unidades
    
           

   On Error GoTo 0
   Exit Sub

imprime_Subtotal_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento imprime_Subtotal de Formulario TmpMaDet"
End Sub

Private Sub Asigna_Grid()
Dim tmprc As New ADODB.Recordset



  If Not adoPrimaryRS.EOF Then
  Set adoDetalleRS = adoPrimaryRS("ChildCMD").UnderlyingValue
  
  
'1. numberformat
'2. numberformat
'3. numberformat
'4. temporada
'6. talla
'7. color
'9 formato %
'10 formato %
'11 formato %
'12 formato %



    With tmprc
        .Open "SELECT IDTEM, TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY IDTEM", locCnn, adOpenDynamic, adLockReadOnly
        tmptempor = fg.BuildComboList(tmprc, "TEMPORADA", "IDTEM", vbBlue)
        .Close
        .Open "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
        tmptalla = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
        .Close
        .Open "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
        tmpcolor = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
        .Close
  '      .Open "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND TEMPOR = " & TemporadaActual & " ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
'        tmpcodigo = fg.BuildComboList(tmprc, "MODELO", "CODIGO", vbBlue)
      '  .Close
    End With
    
    
    
     'Set fg.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
    'fg.DataMode = flexDMBoundBatch
    
    Call Re_Asigna_Grid
  
  'formato para el grid
 
  Gasignados = True
  
  End If


Set tmprc = Nothing
End Sub


Private Sub cbSeleccionaArticulo_Click()
Dim artSQL As New clsSmartSQL
Dim mrc As New ADODB.Recordset

With artSQL
    .AddTable "MAARTIC"
    .AddOrderClause "CODIGO"
    .AddSimpleWhereClause "MBAJA", "0"
    .AddSimpleWhereClause "HIST", "0"
End With

mrc.Open artSQL.SQL, locCnn, adOpenStatic, adLockReadOnly

    With frmFlexArt
    
    Set .miOsql = artSQL
    Set .miRc = mrc
    'el recordset de este formulario
    
    Set .rc_detalle = adoDetalleRS
    .NumeroPedido = adoPrimaryRS.Fields("NUMERO")
    .add_en_detalle = True
    
    .Show 1
    DoEvents
    
    End With
    
    mrc.Close
    Set mrc = Nothing
'tmpbook = rc.Bookmark
     
'rc.Bookmark = tmpbook
'Set tmpbook = Nothing
Set artSQL = Nothing

End Sub

Private Sub cmEntrada_Click()

With adoDetalleRS

  If .EOF Then Exit Sub
  .MoveFirst
  
    Do Until .EOF
    
    
    Loop
    
  .Update
  
  'End If
'stock  .Fields("CODART"),
End With

End Sub

Private Sub dtActualizar_Click()

   On Error GoTo dtActualizar_Click_Error

With adoDetalleRS
    If .EditMode <> adEditNone And .EditMode <> adEditInProgress Then
            .UpdateBatch
    End If
End With

   On Error GoTo 0
   Exit Sub

dtActualizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure dtActualizar_Click of Formulario TmpMaDet"

End Sub

Private Sub dtAgregar_Click()
'Dim tmpcodigo As Long

''**********************************************************
' ATENCION
' Poner en el detalle de pedidos como clave 2 campos:
' NUMERO
' LINEA
' donde linea empieza a contar desde 1 en cada pedido
''**********************************************************
If Not Gasignados Then Call Asigna_Grid

'If Not adoDetalleRS.EOF Then
 '   If adoDetalleRS.EditMode <> adEditNone Then
  '  Call dtActualizar_Click
'    End If
'End If

'tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.Fields("NUMERO").Value, locCnnSP)

'If adoPrimaryRS.EOF Then

'With adoDetalleRS
    'editagrid = True
   ' .AddNew
  '  .Fields("LINEA") = tmpcodigo
 '   .Fields("TEMPOR") = TemporadaActual
'End With

'End If

 'With fg
  '  .SetFocus
   ' .Row = .Row + 1
 '   .Col = 3
'    .EditCell
 'End With
 
DoEvents


With FrmMntArt
    .NumeroPedido = adoPrimaryRS.Fields("NUMERO")
    .add_en_detalle = True
    Set .rc_detalle = adoPrimaryRS
    .Show
    DoEvents
End With





End Sub



Private Sub dtEdicion_Click()

editagrid = True

With fg
    .SetFocus
    .Col = 3
    .Row = .Rows - 1
    .EditCell
End With

End Sub

Private Sub dtEliminar_Click()

On Error Resume Next
With adoDetalleRS
    If Not .EOF Then
             .Delete
             .MoveFirst
    End If
End With

End Sub

Private Sub fg_dblClick()

   If adoPrimaryRS.RecordCount = 0 Then Exit Sub
    
   With fg

   If .TextMatrix(.Row, 1) = "" Or Not IsNumeric(.TextMatrix(.Row, 1)) Then Exit Sub

        'ir al registro especificado
        adoDetalleRS.MoveFirst
        adoDetalleRS.Find "LINEA = " & .TextMatrix(.Row, 1), , adSearchForward
        
   End With
    
   If Not Gasignados Then Call Asigna_Grid
   
   With frmDPedPro
      .NumeroPedido = ioNUMERO.Text
      If Not adoPrimaryRS.EOF Then
      Set .rc = adoDetalleRS
      End If
    .Show
   End With
   
End Sub

Private Sub Form_Load()
 
  Dim cabSQL As String
  Dim detSQL As String
  
  'Detalle de pedidos
  oSQL.AddTable "DETPEDPRO"
  oSQL.AddOrderClause "CODART"
  oSQL.AddOrderClause "LINEA"
  

  detSQL = oSQL.SQL

  Set oSQL = Nothing
  
  'Cabecera de pedidos
  oSQL.AddTable "CABPEDPRO"
  oSQL.AddOrderClause "NUMERO"
  
    
  cabSQL = oSQL.SQL
  
   With locCnnSP
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnnSP
    End If
   End With
     
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {" & cabSQL & "} AS ParentCMD APPEND ({" & detSQL & "} AS ChildCMD RELATE NUMERO TO NUMERO) AS ChildCMD", locCnnSP, adOpenStatic, adLockBatchOptimistic

  cbCODPROV.CodigoWidth = 800

  'obtener el IDTEM de la temporada actual
  'tmptemporada = devuelve_campo("select IDTEM from TEMPOR where ACTUAL = 1", locCnnSP)

  'si no hay registros
  If adoPrimaryRS.EOF Then Exit Sub
  'adoPrimaryRS.MoveFirst
  
  Call cmdFirst_Click
  DoEvents
  
  'cargar campos y grid
  Call Asigna_Campos
  
    fg.AllowSelection = True
    fg.HighLight = flexHighlightAlways
    fg.SelectionMode = flexSelectionByRow
    
    Call SetButtons(True)
    
  '
  
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
    If editagrid = False Then cbcerrar_Click
    Case vbKeyEnd
      If editagrid = False Then cmdLast_Click
    Case vbKeyHome
      If editagrid = False Then cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      
      If editagrid = False Then
      
        If Shift = vbCtrlMask Then
            cmdFirst_Click
        Else
            cmdPrevious_Click
        End If
      
      End If
      
      Case vbKeyDown, vbKeyPageDown
      
      If editagrid = False Then
      
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
      
      End If
      
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

tmptempor = ""
tmptalla = ""
tmpcolor = ""
'tmpcodigo = ""

Set TmpMaDet = Nothing
Set oSQL = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
 
 With adoPrimaryRS
 
 If .AbsolutePosition > 0 Then
 lblstatus.caption = "Registro: " & CStr(adoPrimaryRS.AbsolutePosition)
  
   ' ioDCTO.Text = .Fields("DCTO")
    ioDCTOPP.Text = .Fields("DCTOPP")
    ioGASTOS.Text = .Fields("GASTOS")
    ioPORTES.Text = .Fields("PORTES")
    
    DoEvents
   ' Call Re_Asigna_Grid
    
 
 End If
 
 End With
 
  
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
 On Error GoTo AddErr
  
  If Not Casignados Then Call Asigna_Campos
  If Not Gasignados Then Call Asigna_Grid
  
  'Call Asigna_Campos
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblstatus.caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    
    ioNUMERO.SetFocus
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbeliminar_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub





Private Sub cbedicion_Click()
  On Error GoTo EditErr

  If adoPrimaryRS.EOF Then Exit Sub

  lblstatus.caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  
  ioNUMERO.SetFocus
  Exit Sub
  

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cbCancelar_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cbactualizar_Click()
  On Error GoTo UpdateErr
      
    
    With adoPrimaryRS
        If .EOF Then Exit Sub
        
        
      '  If ioDCTO.Text = "" Then ioDCTO.Text = 0
        If ioDCTOPP.Text = "" Then ioDCTOPP.Text = "" = 0
        If ioGASTOS.Text = "" Then ioGASTOS.Text = 0
        If ioPORTES.Text = "" Then ioPORTES.Text = 0
    
        
      '  .Fields("DCTO") = ioDCTO.Text
        .Fields("DCTOPP") = ioDCTOPP.Text
        .Fields("GASTOS") = ioGASTOS.Text
        .Fields("PORTES") = ioPORTES.Text
    End With
  
  
  ioNUMERO.Locked = True
  
  'cambiar el codigo de usuario (DEPENDIENTE)
  adoPrimaryRS.Fields("CODUSR") = 6
  
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
On Error GoTo GoFirstError

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MoveFirst
  mbDataChanged = False
  
  Call Re_Asigna_Grid

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
  
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError
  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  
  Call Re_Asigna_Grid

  Exit Sub

GoLastError:
 If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description & Err.Number
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

    If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
   adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False
  
  Call Re_Asigna_Grid

  Exit Sub
GoNextError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False
  
   Call Re_Asigna_Grid

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
   
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  cbCODPROV.Locked = bVal
  cbFPAGO.Locked = bVal
  cbPLAZOE.Locked = bVal
  cbCODALM.Locked = bVal
  
'  Call Habilita_miTextNum(Me, Not bVal)
  
End Sub

'Private Sub fg_ChangeEdit()
'
'With fg

'Select Case .Col
 '
 '
''IMPORTE y UNIDADES
'Case 5, 6
'
'    .EditText = Replace(.EditText, ",", ".")
    
'End Select
'
'End With
'
'
'End Sub

'Private Sub fg_EnterCell()
'
'With fg
'        If .DataSource Is Nothing Then Exit Sub
'
'        .BackColorSel = vbBlack
'        .CellFontBold = True
'        .CellForeColor = vbWhite
'        .CellBackColor = vbBlue
'End With
'
'
'        'lblValue.ForeColor = vbRed
'  '  Else
'        '.BackColorSel = &H8000000D 'd
'End Sub
'
'Private Sub fg_GotFocus()
'
'    editagrid = True
'
'End Sub



'Private Sub fg_LeaveCell()

'On error Resume Next
' With fg
'    If .DataSource Is Nothing Then Exit Sub
'   .CellBackColor = &H80000005
'   .BackColorSel = &H80000005
'   .CellFontBold = False
'   .CellForeColor = vbBlack
' End With
'
'End Sub

'Private Sub fg_LostFocus()
'
'    editagrid = False
'
'End Sub


'Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'
'    With fg
'
'    Select Case Col
'
'    'CODIGO
'    Case 3
'
'
'     If Not IsNumeric(.EditText) Or Trim(.EditText) = "" Or IsNull(.EditText) Then
'
'        lblStatusD.Caption = "CODIGO en blanco, o no valido"
'        .CellBackColor = vbYellow
'        Cancel = True
'        .EditText = ""
'        .Text = ""
'
'
'     Else
'
'        lblStatusD.Caption = ""
'      '  .Col = .Col + 1
'       ' .EditCell
'
'     End If
'
'
'
'    'DENOMINACION
'    Case 4
'
'
'     If Trim(.EditText) = "" Or IsNull(.EditText) Then
'
'        lblStatusD.Caption = "No se permite DENOMINACION en blanco"
'        .CellBackColor = vbYellow
'        Cancel = True
'        .EditText = ""
'        .Text = ""
'
'
'     Else
'
'        lblStatusD.Caption = ""
'      '  .Col = .Col + 1
'      '  .EditCell
'
'     End If
'
'     'IMPORTE
'     Case 5
'
'
'     If Not IsNumeric(.EditText) Or Trim(.EditText) = "" Or IsNull(.EditText) Then
'
'        lblStatusD.Caption = "IMPORTE en blanco o no valido."
'        .CellBackColor = vbYellow
'        Cancel = True
'        .EditText = ""
'        .Text = ""

'     'si es valido
'     Else
'
'        lblStatusD.Caption = ""
'      '  .Col = .Col + 1
'      '  .EditCell
'
'     End If
'
'
'    End Select
'
'    End With
'

'End Sub
'
Private Sub ioDCTOPP_Validate(Cancel As Boolean)

If mbEditFlag Or mbAddNewFlag Then
    
With ioDCTOPP

    If Trim(.Text) <> "" Then
        If CDbl(Replace(.Text, ".", ",")) >= 100 Then
            lblstatus.caption = "No se permite un descuento por pronto pago del 100%"
            .CancelarValidacion
            Cancel = True
            DoEvents
            Exit Sub
        
        Else
            lblstatus.caption = ""
            DoEvents
        End If
    End If

End With

End If
 
End Sub



'Private Sub ioDCTO_Validate(Cancel As Boolean)

'If mbEditFlag Or mbAddNewFlag Then
 '
'With ioDCTO
'
 '   If Trim(.Text) <> "" Then
  '      If CDbl(.Text) >= 100 Then
   '         lblstatus.Caption = "No se permite un descuento del 100%"
    '        .CancelarValidacion
     '       Cancel = True
      '      DoEvents
       '     Exit Sub
        
 '       Else
  '          lblstatus.Caption = ""
   '         DoEvents
    '    End If
 '   End If
'
'End With

'Tab1.Tab = 1

'End If
 
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : CreateDBEtiquetas
' DateTime  : 10/11/2003 20:36
' Author    : Administrador
' Purpose   : Rutina que crea la base de datos temporal donde se alma
'             cenaran los registros que se van al imprimir como etiquetas
'             (un registro por cada unidad articulo/talla/color.
'---------------------------------------------------------------------------------------
'
Private Sub CreateDBEtiquetas(rc_detped As ADODB.Recordset)
On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Const fichero = "c:\TempEtiquetasDB.mdb"
Dim etiqrc As New ADODB.Recordset
Dim nveces As Long

ChDir ("c:\")

'si existe uno previo, borrar
If Dir(fichero) <> "" Then Kill fichero

sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & fichero

Cat.Create sCnn

  '----------* Table Definition of ETIQUETAS *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "ETIQUETAS"
    .Columns.Append "ABREVIA", adVarWChar, 20
    .Columns.Append "CODCOLOR", adVarWChar, 3
      .Columns("CODCOLOR").Properties("Default").Value = "0"
    .Columns.Append "CODIGO", adVarWChar, 5
      .Columns("CODIGO").Properties("Default").Value = "0"
    .Columns.Append "CODTALLA", adVarWChar, 3
      .Columns("CODTALLA").Properties("Default").Value = "0"
    .Columns.Append "DESCOLOR", adVarWChar, 15
    .Columns.Append "DESCTALLA", adVarWChar, 15
    .Columns.Append "DESCTEMP", adVarWChar, 20
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    .Columns.Append "PVP", adCurrency
      .Columns("PVP").Properties("Default").Value = "0"
    .Columns.Append "TEMPOR", adVarWChar, 3
      .Columns("TEMPOR").Properties("Default").Value = "0"
  End With
  '----------* Index Definitions of ETIQUETAS *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Tbl(0).Indexes.Append Idx(0)

  Cat.Tables.Append Tbl(0)

  Set Cat = Nothing
  
'tempor
'codart
'codtalla
'codcol
'preven
  
  etiqrc.Open "SELECT * FROM ETIQUETAS", sCnn, adOpenDynamic, adLockOptimistic
  
  'ahora meterle los datos ....
  With rc_detped
    
    If Not .BOF Then .MoveFirst
    
    Do Until .EOF
    
        'un registro por cada unidad para el mismo articulo
        For nveces = 1 To .Fields("UNIDADES").Value
    
        etiqrc.AddNew
        etiqrc.Fields("ABREVIA") = .Fields("CODART") & "-" & devuelve_campo("SELECT ABREVIA FROM MAARTIC WHERE CODIGO =" & .Fields("CODART"))
        etiqrc.Fields("TEMPOR") = Format(.Fields("TEMPOR"), "000")
        etiqrc.Fields("CODIGO") = Format(.Fields("CODART"), "00000")
        etiqrc.Fields("CODTALLA") = Format(.Fields("CODTALLA"), "000")
        etiqrc.Fields("CODCOLOR") = Format(.Fields("CODCOL"), "000")
        etiqrc.Fields("DESCTEMP") = .Fields("TEMPOR") & "-" & devuelve_campo("SELECT TEMPORADA FROM TEMPOR WHERE IDTEM =" & .Fields("TEMPOR"))
        etiqrc.Fields("DESCOLOR") = .Fields("CODCOL") & "-" & devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO =" & .Fields("CODCOL"))
        etiqrc.Fields("DESCTALLA") = .Fields("CODTALLA") & "-" & devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO =" & .Fields("CODTALLA"))
        etiqrc.Update
        
        Next
        
    .MoveNext
    Loop
    
  End With
  
  etiqrc.Close
  Set etiqrc = Nothing
  
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Code Gen Error")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : generar_etiquegtas
' DateTime  : 09/11/2003 21:45
' Author    : Administrador
' Purpose   : Genera las etiquetas y el informe Crystal Report
'---------------------------------------------------------------------------------------
'
Private Sub generar_etiquetas()
Dim tmprcped As New ADODB.Recordset

 '  On Error GoTo generar_etiquegtas_Error

   'crear la base de datos etiquetas'
   
   If adoPrimaryRS.RecordCount = 0 Then Exit Sub
   
   
   If MsgBox("¿Desea generar las etiquetas para el pedido actual nº " & ioNUMERO.Text, vbQuestion + vbYesNo) = vbYes Then
    
        tmprcped.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & ioNUMERO.Text, locCnnSP
    
        Call CreateDBEtiquetas(tmprcped)
        DoEvents
        Call procesa_informes(1)
        DoEvents
        adoPrimaryRS.Requery
        DoEvents
        Call Asigna_Grid
    
        tmprcped.Close
        Set tmprcped = Nothing
   
   End If
   
  
   Set tmprcped = Nothing
   On Error GoTo 0
   Exit Sub

generar_etiquegtas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_etiquetas of Formulario TmpMaDet"

End Sub

Private Sub ucGrdBttn1_Click()
Call generar_etiquetas
End Sub



