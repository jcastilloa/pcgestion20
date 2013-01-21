VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias de Mercancía"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
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
   ScaleHeight     =   6960
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin VSFlex8LCtl.VSFlexGrid fg 
      Height          =   4095
      Left            =   30
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   930
      Width           =   11160
      _cx             =   19685
      _cy             =   7223
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14260872
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"frmPtrans.frx":0000
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
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5460
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":00DE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":010A
      picn            =   "frmPtrans.frx":0128
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
      Height          =   375
      Left            =   0
      Top             =   5055
      Width           =   11190
      _extentx        =   19738
      _extenty        =   661
      caption         =   ""
      fount           =   "frmPtrans.frx":0DFC
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   15
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":0E2A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":0E56
      picn            =   "frmPtrans.frx":0E74
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
      Left            =   9045
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":1BAC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":1BD8
      picn            =   "frmPtrans.frx":1BF6
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
      Left            =   10125
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":28CA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":28F6
      picn            =   "frmPtrans.frx":2914
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6135
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":364C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":3678
      picn            =   "frmPtrans.frx":3696
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
      TabIndex        =   4
      Top             =   6135
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":4372
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":439E
      picn            =   "frmPtrans.frx":43BC
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
      Left            =   8925
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6135
      Width           =   1170
      _extentx        =   2064
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":4C98
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":4CC4
      picn            =   "frmPtrans.frx":4CE2
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
      Left            =   10125
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6135
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":55BE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":55EA
      picn            =   "frmPtrans.frx":5608
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALMDEST 
      Height          =   480
      Left            =   810
      TabIndex        =   0
      Top             =   420
      Width           =   4095
      _extentx        =   7223
      _extenty        =   847
      font            =   "frmPtrans.frx":62E4
   End
   Begin PCGestion.bsGradientLabel lblMensajes 
      Height          =   375
      Left            =   5790
      Top             =   6075
      Width           =   2565
      _extentx        =   4524
      _extenty        =   661
      caption         =   ""
      fount           =   "frmPtrans.frx":6310
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   15640462
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   540
      Left            =   6000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5460
      Width           =   1065
      _extentx        =   1879
      _extenty        =   953
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":633E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":636A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblEstado 
      Height          =   375
      Left            =   2895
      Top             =   6075
      Width           =   2655
      _extentx        =   4683
      _extenty        =   661
      caption         =   ""
      fount           =   "frmPtrans.frx":6388
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   15640462
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton dtAgregar 
      Height          =   540
      Left            =   4380
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5460
      Width           =   780
      _extentx        =   1376
      _extenty        =   953
      btype           =   9
      tx              =   "Agregar"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":63B6
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":63E2
      picn            =   "frmPtrans.frx":6400
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdVerMensajes 
      Height          =   540
      Left            =   7095
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5460
      Width           =   1485
      _extentx        =   2619
      _extenty        =   953
      btype           =   9
      tx              =   "Ver Mensajes"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":660C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":6638
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdGenerarCodigo 
      Height          =   420
      Left            =   2505
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5535
      Visible         =   0   'False
      Width           =   1845
      _extentx        =   3254
      _extenty        =   741
      btype           =   3
      tx              =   "Generar Código"
      enab            =   0   'False
      font            =   "frmPtrans.frx":6656
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   16776960
      fcolo           =   16776960
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":6682
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioDCTO 
      Height          =   525
      Left            =   10095
      TabIndex        =   3
      Top             =   435
      Width           =   645
      _extentx        =   1138
      _extenty        =   926
      font            =   "frmPtrans.frx":66A0
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblAlmOrig 
      Height          =   375
      Left            =   4230
      Top             =   30
      Width           =   3225
      _extentx        =   5689
      _extenty        =   661
      caption         =   ""
      fount           =   "frmPtrans.frx":66CC
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   14457707
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmCambioDCTO 
      Height          =   420
      Left            =   10725
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   450
      Width           =   435
      _extentx        =   767
      _extenty        =   741
      btype           =   9
      tx              =   "DCTO"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":66FA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":6726
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmAnularTrans 
      Height          =   405
      Left            =   5775
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6495
      Width           =   2595
      _extentx        =   4577
      _extenty        =   714
      btype           =   3
      tx              =   "Anular Transferencia"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":6744
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":6770
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmTerminarTrans 
      Height          =   405
      Left            =   2880
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2670
      _extentx        =   4710
      _extenty        =   714
      btype           =   3
      tx              =   "&Terminar Transferencia"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":678E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":67BA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdAceptarTrans 
      Height          =   420
      Left            =   2895
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   2670
      _extentx        =   4710
      _extenty        =   741
      btype           =   3
      tx              =   "Aceptar Transferencia"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":67D8
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   16776960
      fcolo           =   16776960
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":6804
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   540
      Left            =   5190
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5460
      Width           =   780
      _extentx        =   1376
      _extenty        =   953
      btype           =   9
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmPtrans.frx":6822
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":684E
      picn            =   "frmPtrans.frx":686C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioGASTOS 
      Height          =   525
      Left            =   7950
      TabIndex        =   2
      Top             =   435
      Width           =   900
      _extentx        =   1588
      _extenty        =   926
      font            =   "frmPtrans.frx":7548
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmCambioGastos 
      Height          =   420
      Left            =   8835
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   450
      Width           =   675
      _extentx        =   1191
      _extenty        =   741
      btype           =   9
      tx              =   "GASTOS"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":7574
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":75A0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioNUMPED 
      Height          =   525
      Left            =   5595
      TabIndex        =   1
      Top             =   435
      Width           =   900
      _extentx        =   1588
      _extenty        =   926
      font            =   "frmPtrans.frx":75BE
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmCambioPedido 
      Height          =   420
      Left            =   6480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   465
      Width           =   675
      _extentx        =   1191
      _extenty        =   741
      btype           =   9
      tx              =   "PEDIDO"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":75EA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":7616
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbIrA 
      Height          =   360
      Left            =   2580
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   45
      Width           =   525
      _extentx        =   926
      _extenty        =   635
      btype           =   9
      tx              =   "&Ir A"
      enab            =   -1  'True
      font            =   "frmPtrans.frx":7634
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmPtrans.frx":7660
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDO"
      Height          =   300
      Left            =   4905
      TabIndex        =   31
      Top             =   510
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GASTOS"
      Height          =   300
      Left            =   7170
      TabIndex        =   29
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO"
      Height          =   300
      Left            =   9510
      TabIndex        =   24
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALM. DESTINO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15
      TabIndex        =   17
      Top             =   375
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN ORIGEN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2955
      TabIndex        =   16
      Top             =   -45
      Width           =   1200
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
      Left            =   8685
      TabIndex        =   9
      Top             =   30
      Width           =   2490
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
      Left            =   855
      TabIndex        =   8
      Top             =   45
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   45
      TabIndex        =   7
      Top             =   75
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ult. modif."
      Height          =   315
      Left            =   7470
      TabIndex        =   6
      Top             =   60
      Width           =   1200
   End
End
Attribute VB_Name = "frmPtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Module    : Plantilla
' DateTime  : 31/10/2003 10:08
' Author    : Administrador
' Purpose   : Plantilla de código para los formularios de maestros.
'---------------------------------------------------------------------------------------
'·································································································································
' Convenio:
'·································································································································
'
' Para los campos de texto:  usar miText
' Para los combos:              usar miCombo.
'
'
' - Instrucciones:
'
' Enlazar los controles a los campos en Form_Load()   (ver ). Y especificar la tabla
' y orden (y otras cosas que se pudieran necesitar) mediante los parametros del
' oSQL (objecto SmartSQL):
'
'  oSQL.AddTable "SECCIONES"
'  oSQL.AddOrderClause "CODIGO"

'---------------------------------------------------------------------------------------
' - Cambiar en:
'
' Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' set plantilla = nothing
' por el nombre del formulario. Ej:   set frmMntPrue = nothing

'---------------------------------------------------------------------------------------
' - Si se utiliza algun campo simulando que sea incremental (que se incremente en cada
' registro) cambia en Private Sub cbAgregar_Click()
'
' tmpcodigo = devuelve_campo("select max(codigo) + 1 from secciones")
'
' y poner el SQL correcto para que nos devuelva el proximo codigo para nuestro campo
'
'
'
'---------------------------------------------------------------------------------------
' - Colocar 2 tipos de validaciones para los datos.
'
'  Una validación a nivel de campo. Por ejemplo, comprobar al salir del campo
'  que la información es correcta, usando el evento validate. (si es > X, <> "", etc)
'
'- Otra validación es en:
'
'Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
' El evento cuando cambia un registro del recordset, pones el codigo por ejemplo en

'        Case adRsnUpdate
'
' Mira el formulario frmMntCli para ver un ejemplo de esto.
'
'---------------------------------------------------------------------------------------
' - Formularios de Lista. Para llamar al formulario de lista estandar FrmFlexSimple, ver el
' codigo de cbLista_Click. Cambiar los colformats y otras cosas que puedan ser
' necesarias, para adecuar a cada formulario
'
'
'---------------------------------------------------------------------------------------
'Otras notas:
'
' - comprobar el orden correcto de los tabindex para permitir recorrer miText y miCombo
' del formulario con el teclado (soportan el avance con ENTER). Desde el primero hasta
' el ultimo.
'
'- cambiar en cbAgregar_clik y cbEditar_click
'
' ioDescripcion.setfocus
'
' por el nombre del control que tengamos que activar en primer lugar.
'
' cambiar en Private Sub ioCODIGO_Change(), y poner el numero de 000
' correcto en cada caso

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

Public añade_nueva As Boolean
Public cargando_Grid As Boolean



'---------------------------------------------------------------------------------------
' Procedimiento : cbImprimir_Click
' Fecha/Hora    : 04/02/2004 11:12
' Autor         : JCastillo
' Propósito     :  Imprimir la transferencia actual.
'---------------------------------------------------------------------------------------
'
Private Sub cbImprimir_Click()
Dim linea1 As String
Dim linea2 As String

    On Error GoTo cbImprimir_Click_Error
        
    DoEvents


    linea1 = "Transferencia de Mercancía. Codigo: " & rc.fields("CODIGO") & ". Origen: " & lblAlmOrig.Caption & ". Destino: " & Trim(devuelve_campo("select descripcion from almacenes where codigo = " & cbCODALMDEST.Text, locCnn)) & "   DCTO: " & rc.fields("DCTO") & " %"
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & rc.fields("FMODI")
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 12, 2)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmPtrans"
End Sub


Private Sub cbIrA_Click()

   On Error GoTo cbIrA_Click_Error

With frmBusTrn
    
    .Show 1
    
    If .miCODALM_ORIG = 0 Then Exit Sub

    cargando_Grid = True

    If Not rc.BOF Then rc.MoveFirst

    Do Until rc.EOF

        'si lo encontramos, salir
        If (rc.fields("CODIGO") = .miCODTRN) And (rc.fields("CODALMORIG") = .miCODALM_ORIG) Then
            lblstatus.Caption = "Se ha encontrado satisfactoriamente"
            Exit Do
        End If

        rc.MoveNext
        
    Loop
    
    cargando_Grid = False
    
    If rc.EOF Then
       lblstatus.Caption = "No se ha encontrado la transferencia"
    Else
       rc.Move 0
    End If

End With

Set frmBusTrn = Nothing


   On Error GoTo 0
   Exit Sub

cbIrA_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbIrA_Click de Formulario frmPtrans"

End Sub


Private Sub cbLista_click()
Dim tmpcodigo As Long
Dim tmpalmorig As Byte
Dim micargado As Boolean

'If rc.EditMode = adEditNone Then

'if rc.RecordCount <=

   On Error GoTo cbLista_click_Error

If rc.RecordCount > 0 Then

cargando_Grid = True

tmpcodigo = rc.fields("CODIGO")
tmpalmorig = rc.fields("CODALMORIG")

End If

With frmFlexPtrans
       
    Set .miosql = oSQL
    
    With .fg
             Set frmFlexPtrans.miRc = rc
             'Set frmFlexPtrans.selecciona_registro = micargado
    End With
    
    .Caption = "Peticiones de Transferencia ..."
    
'des - enlazar campos por si cambia el filtro del recordset
With ioCODIGO
    Set .DataSource = Nothing
        .DataField = ""
End With
      
'With cbCODALMORIG
'        .DataField = ""
 '   Set .DataSource = Nothing
'End With

With cbCODALMDEST
        .DataField = ""
    Set .DataSource = Nothing
End With
  
With ioFMODI
    Set .DataSource = Nothing
        .DataField = ""
End With

With ioDCTO
    Set .DataSource = Nothing
        .DataField = ""
End With

With ioNUMPED
    Set .DataSource = Nothing
        .DataField = ""
End With
    
    .Show 1
    
    DoEvents
    
    cargando_Grid = False
    
    If Not frmFlexPtrans.selecciona_registro Then
    
    If rc.EOF And rc.RecordCount > 0 Then
         rc.MovePrevious
    ElseIf rc.RecordCount > 0 Then
     '   rc.Move 0
    End If
    
    Else
    
        rc.Move 0
    End If
    
   ' rc.MoveFirst
   ' rc.Find "(CODIGO =" & tmpcodigo & ")" 'AND (CODALMORIG =" & tmpalmorig & ")", , adSearchForward
    
        
  '  Call cmdFirst_Click
    
'volver a enlazar los campos
With ioCODIGO
    Set .DataSource = rc
        .DataField = "CODIGO"
End With
      
With ioNUMPED
    Set .DataSource = rc
        .DataField = "NUMPED"
End With

With ioDCTO
    Set .DataSource = rc
        .DataField = "DCTO"
End With

'With cbCODALMORIG
 '       .DataField = "CODALMORIG"
'    Set .DataSource = rc
'End With

With cbCODALMDEST
        .DataField = "CODALMDEST"
    Set .DataSource = rc
End With
  
With ioFMODI
    Set .DataSource = rc
        .DataField = "FMODI"
End With
         
End With


Set frmFlexPtrans = Nothing

   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmPtrans"

End Sub

Private Sub cbcodalmdest_Validate(Cancel As Boolean)

If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub

If cbCODALMDEST.Text = "" Then Exit Sub

'comprobar q no sea origen y destino iguales
If (cbCODALMDEST.Text = AlmacenActual) Then

 lblstatus.Caption = "ORIGEN y DESTINO deben ser diferentes"
 cbCODALMDEST.SetFocus
 Cancel = True
 Exit Sub

End If

DoEvents

'actualizar y agregar un nuevo registro
'Call cbactualizar_Click
End Sub






'---------------------------------------------------------------------------------------
' Subrutina   : AnularTrans
' Fecha/Hora  : 06/12/2003 14:39
' Autor       : JCASTILLO
' Propósito   : Anular (CANCELAR) la transferencia. O borrar el registro si la
' transferencia esta en estado 0 (en creación)
'---------------------------------------------------------------------------------------
Private Sub cmAnularTrans_Click()
Dim tmpstrconn As String
       
 On Error GoTo AnularTrans

 If rc.RecordCount <= 0 Then Exit Sub
 If MsgBox("¿Desea anular la transferencia seleccionada? nº " & ioCODIGO.Caption & ".", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    cmAnularTrans.Visible = False
    
    Select Case rc.fields("ESTADO")
    
    Case 0
    
        'si esta en creacion, borrar el registro sin mas
        With locCnn
            .Execute "DELETE FROM DETTRANS WHERE CODIGO = " & CLng(ioCODIGO.Caption)
             DoEvents
            .Execute "DELETE FROM PTRANS WHERE CODIGO = " & CLng(ioCODIGO.Caption)
        End With
        
        rc.Requery
        Call cmdFirst_Click
    
    Case 1 'estado pendiente
    
        'si no esta editando o añadiendo, marcar como editando = true
        If rc.fields("CODALMORIG") <> AlmacenActual Then
            lblstatus.Caption = "Esta transferencia solo puede anularse desde el ORIGEN (" & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALMORIG"), locCnn)) & ")"
            Exit Sub
        
        'si pertenece al almacén actual ...
        Else
                            
                    rc.ActiveConnection = Nothing
                    Call des_enlaza_campos
        
                    With locCnn
                        tmpstrconn = .ConnectionString  'guardar el connection anterior por si acaso
                        If .State <> 0 Then .Close
                           .CursorLocation = adUseServer  'abrir con cursor para las transacciones
                           .Open strLocCnn
                          .BeginTrans
                    End With
                    
                    'Anular la transferencia, para eso se sumaran las unidades al origen (ya han sido descontadas al pasar la
                    'transferencia a PENDIENTE
                    Select Case anular_transferencia_pendiente(rc.fields("CODIGO"), rc.fields("CODALMORIG"), locCnn)
                
                    Case 0
                        
                        locCnn.CommitTrans
                        lblstatus.Caption = "La transferencia se ha anulado correctamente"
                        lblEstado.Caption = "CANCELADA"
                    
                    Case 1
                        
                        locCnn.RollbackTrans
                        lblstatus.Caption = "La transferencia NO ha podido anularse"
                                            
                    End Select
                               
                    With locCnn
                        If .State <> 0 Then .Close
                        .CursorLocation = adUseClient
                        .Open tmpstrconn
                    End With
                
                    Call enlaza_campos
                    Set rc.ActiveConnection = locCnn
        
        End If
        
        'If (Not mbEditFlag) And (Not mbAddNewFlag) Then mbEditFlag = True
       
        'rc.Fields("ESTADO") = 3
        'Call cbactualizar_Click
        rc.Requery
        Call cmdFirst_Click
    
   End Select
    
   
   tmpstrconn = ""
   
   On Error GoTo 0
   Exit Sub

AnularTrans:
                    
    With locCnn
            If .CursorLocation = adUseServer Then
            .RollbackTrans
            If .State <> 0 Then .Close
            .CursorLocation = adUseClient  'abrir con cursor para las transacciones
            .Open tmpstrconn
        End If
    End With

    tmpstrconn = ""

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmAnularTrans_Click de Formulario frmPtrans"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : des_enlaza_campos
' Fecha/Hora    : 11/12/2003 11:12
' Autor         : JCastillo
' Propósito     :  Desenlaza los controles para cerrar la conexión
'---------------------------------------------------------------------------------------
Private Sub des_enlaza_campos()

'des-enlazar controles
With ioCODIGO
  Set .DataSource = Nothing
        .DataField = ""
End With
      
'With cbCODALMORIG
'  Set .DataSource = Nothing
'        .DataField = ""
'End With

With cbCODALMDEST
  Set .DataSource = Nothing
        .DataField = ""
End With
  
With ioFMODI
  Set .DataSource = Nothing
        .DataField = ""
End With

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : enlaza_campos
' Fecha/Hora    : 11/12/2003 11:13
' Autor         : JCastillo
' Propósito     :   Vuelve a enlazar los campos al recordset
'---------------------------------------------------------------------------------------
Private Sub enlaza_campos()

With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
End With
      
'With cbCODALMORIG
'    .DataField = "CODALMORIG"
'    Set .DataSource = rc
'End With

With cbCODALMDEST
    .DataField = "CODALMDEST"
    Set .DataSource = rc
End With
  
On Error Resume Next
With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
End With

On Error GoTo 0

End Sub

Private Sub cmCambioGastos_Click()
Dim tmpgas As String

   On Error GoTo cmCambioGastos_Click_Error

 If mbEditFlag Or mbAddNewFlag Then Exit Sub

'obtener ioGASTOS
If ioGASTOS.Text = "" Then ioGASTOS.Text = "0"
tmpgas = InputBox("Cambiar Gastos", "Cambiar Gastos", ioGASTOS.Text)

'validaciones
If Trim(tmpgas) = "" Then Exit Sub
If Not IsNumeric(tmpgas) Then Exit Sub

'aplicar cambios
rc.fields("GASTOS") = Replace(tmpgas, ",", ".")
rc.UpdateBatch adAffectAll

DoEvents

rc.Move 0

   On Error GoTo 0
   Exit Sub

cmCambioGastos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmCambioGastos_Click de Formulario frmPtrans"
End Sub

Private Sub cmCambioPedido_Click()
Dim tmpped As String

   On Error GoTo cmCambioPedido_Click_Error

If mbEditFlag Or mbAddNewFlag Then Exit Sub

'obtener ioGASTOS
If ioNUMPED.Text = "" Then ioNUMPED.Text = "0"
tmpped = InputBox("Cambiar Número de Pedido", "Cambiar Número de Pedido", ioNUMPED.Text)

'validaciones
If Trim(tmpped) = "" Then Exit Sub
If Not IsNumeric(tmpped) Then Exit Sub

ioNUMPED.Text = tmpped
If comprueba_estado_pedido = True Then
    rc.fields("NUMPED") = 0
    rc.Update
    ioNUMPED.Text = "0"
    Exit Sub
End If

'aplicar cambios
rc.fields("NUMPED") = CDbl(tmpped)
rc.UpdateBatch adAffectAll

DoEvents

rc.Move 0

   On Error GoTo 0
   Exit Sub

cmCambioPedido_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmCambioPedido_Click de Formulario frmPtrans"
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cmdAceptarTrans_Click
' Fecha/Hora     : 09/12/2003 13:53
' Autor             : JCastillo
' Propósito        :  Botón de aceptar transferencia
'---------------------------------------------------------------------------------------
Private Sub cmdAceptarTrans_Click()

On Error GoTo cmdAceptarTrans_Click_Error

If rc.RecordCount <= 0 Then Exit Sub

lblstatus.Caption = "Aceptando transferencia ..."

With frmComPtrans
    .ID_TRANSF = rc.fields("CODIGO")
    .ALM_TRANSF = rc.fields("CODALMORIG")
    .Show 1
    
    'si es distinto de 1, no aceptar transferencia
    If .estado <> 1 Then
        
        lblstatus.Caption = "Imposible Aceptar Transferencia"
        Set frmComPtrans = Nothing
        Exit Sub
    
    End If
    
End With

Set frmComPtrans = Nothing

cmdAceptarTrans.Visible = False

Call des_enlaza_campos
    
'se intenta aceptar la transferencia
Select Case aceptar_transferencia(rc.fields("CODIGO"), rc.fields("CODALMORIG"), locCnn)

'salida por error
Case 1
    MsgBox "Se ha producido un error al aceptar la transferencia", vbExclamation
    
'salida por que no hay registros en el detalle
Case 2
    MsgBox "No hay ningun artículo en esa transferencia. Imposible aceptar.", vbInformation

'salida ok
Case 0
    MsgBox "Transferencia ACEPTADA correctamente", vbInformation
    
End Select

'volver a abrir recordset
If rc.State <> 0 Then rc.Close
rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic

'enlazar controles

Call enlaza_campos

rc.Requery
Call cmdFirst_Click

   On Error GoTo 0
   Exit Sub

cmdAceptarTrans_Click_Error:

    lblstatus.Caption = "Se ha producido un error al aceptar la transferencia"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdAceptarTrans_Click de Formulario frmPtrans"

End Sub

Private Sub cmdGenerarCodigo_Click()
'Dim m As Long

With frmGenCodSc
    .CodigoTrn = rc.fields("CODIGO")
    .AlmacenTrn = rc.fields("CODALMORIG")
    .Show
End With

'---------------------------------------------------------------------------------------
'                     9 digitos codigo de transferencia
'                     3 digitos codigo de almacen
'                     Es un codigo de seguridad para poder aceptar la transferencia
'                     incluso si no coinciden las prendas en la comprobación.
'---------------------------------------------------------------------------------------
' m = CodigoSeguridad_TRN(Format(rc.fields("CODIGO"), "000000000") & Format(rc.fields("CODALMORIG"), "000"))
 
' MsgBox "El Código de Seguridad generado es: " & m, vbInformation, titulo


End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmdVerMensajes_Click
' Fecha/Hora  : 08/12/2003 16:40
' Autor       : JCASTILLO
' Propósito   : Muestra el formulario frMsgPtrans para ver los mensajes de transferencias
'               asociados a la transferencia actual
'---------------------------------------------------------------------------------------
Private Sub cmdVerMensajes_Click()

   On Error GoTo cmdVerMensajes_Click_Error

If rc.RecordCount <= 0 Then Exit Sub

With frmMsgPtrans
    .miCODALMORIG = rc.fields("CODALMORIG")
    .TRNSF_ACTUAL = rc.fields("CODIGO")
    Me.WindowState = vbMinimized
    DoEvents
    .Show
    
    
End With

   On Error GoTo 0
   Exit Sub

cmdVerMensajes_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdVerMensajes_Click de Formulario frmPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmTerminarTrans_Click
' Fecha/Hora  : 06/12/2003 14:22
' Autor       : JCASTILLO
' Propósito   : Terminar transferencia: marcar el estado como PENDIENTE
'               y guardar el registro
'---------------------------------------------------------------------------------------
Private Sub cmTerminarTrans_Click()
Dim tmpstrconn As String
Dim rctrn As New ADODB.Recordset

   On Error GoTo cmTerminarTrans_Click_Error
  
    'si no esta editando o añadiendo, marcar como editando = true
    
   If rc.RecordCount <= 0 Then Exit Sub
    
   If (Not mbEditFlag) And (Not mbAddNewFlag) Then mbEditFlag = True
                
   If ioNUMPED.Text = "" Then
    ioNUMPED.Text = "0"
   End If
                
   If ioNUMPED.Text <> "0" Then
    If comprueba_pedido(ioNUMPED.Text) = 2 Then Exit Sub
   End If

    cmTerminarTrans.Visible = False
        
   ' rc.ActiveConnection = Nothing
   ' Call des_enlaza_campos
        
    'With locCnn
   '         tmpstrconn = .ConnectionString  'guardar el connection anterior por si acaso
   '         If .State <> 0 Then .Close
   '         .CursorLocation = adUseServer  'abrir con cursor para las transacciones
    '        .Open strLocCnn
   ' '        .BeginTrans
   ' End With
   
    
    'si hay pedido, que me devuelva todas las de ese pedido
   If ioNUMPED.Text > 0 Then
   
     rctrn.Open "SELECT CODIGO, CODALMORIG FROM PTRANS WHERE NUMPED = " & ioNUMPED.Text & " AND ESTADO = 0", locCnn, adOpenStatic, adLockReadOnly
   
   'si no hay pedido que solo me devuelva una transferencia
   Else
     
     rctrn.Open "SELECT CODIGO, CODALMORIG FROM PTRANS WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODALMORIG = " & rc.fields("CODALMORIG") & " AND ESTADO = 0", locCnn, adOpenStatic, adLockReadOnly
     
   End If
   
        
     Do Until rctrn.EOF
        
        '----------------------------------------------------------------------------------------------------
        'intentar iniciar la transferenca (marcar como pendiente y meter unidades en stock)
        Select Case iniciar_transferencia(rctrn.fields("CODIGO"), rctrn.fields("CODALMORIG"), locCnn)
    
        Case 0
          '  locCnn.CommitTrans
            lblstatus.Caption = "Transferencia terminada correctamente."
        
        Case 1
          '  locCnn.RollbackTrans
            lblstatus.Caption = "Se ha producido un error."
        
        Case 2
          '  locCnn.RollbackTrans
            lblstatus.Caption = "No hay artículos en esta transferencia"
        
        End Select
        '----------------------------------------------------------------------------------------------------
    
        rctrn.MoveNext
     
     Loop
     
     
     
     rctrn.Close
     Set rctrn = Nothing
            

            
   ' With locCnn
   '         If .State <> 0 Then .Close
   '         .CursorLocation = adUseClient
   '         .Open tmpstrconn
   ' End With
       
   ' Set rc.ActiveConnection = locCnn
     
   ' Call enlaza_campos
    
    'rc.Fields("ESTADO") = 1
    lblEstado.Caption = "PENDIENTE"
        
  '  If ioNUMPED.Text <> "0" Then
  '   'pasar el pedido a TRANSFERIDO
  '   locCnn.Execute "UPDATE CABPEDPRO SET ESTADO = 4 WHERE NUMERO = " & rc.fields("NUMPED") & " AND ALMORIG = " & rc.fields("CODALMORIG")
   ' End If
    
    Call cbactualizar_Click
    
    rc.Requery
    Call cmdFirst_Click
    
    tmpstrconn = ""
    
    'limpiar el .log temporal de sqlsqrver
    locCnn.Execute "BACKUP LOG LOCAL WITH TRUNCATE_ONLY"
    
   On Error GoTo 0
   Exit Sub

cmTerminarTrans_Click_Error:
    
    'With locCnn
    '        If .CursorLocation = adUseServer Then
    '        .RollbackTrans
    '        If .State <> 0 Then .Close
   ''         .CursorLocation = adUseClient  'abrir con cursor para las transacciones
    '        .Open tmpstrconn
   '         End If
    'End With
    
    'Set rc.ActiveConnection = locCnn
    
    tmpstrconn = ""

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmTerminarTrans_Click de Formulario frmPtrans"
End Sub



Private Sub cmCambioDCTO_Click()

Dim tmpdcto As String

'solo cambiar cuando no este editando o añadiendo
   On Error GoTo ioDCTO_Click_Error

If mbEditFlag Or mbAddNewFlag Then Exit Sub

'obtener dcto
If ioDCTO.Text = "" Then ioDCTO.Text = "0"
tmpdcto = InputBox("Cambiar Descuento", "Cambiar Descuento", ioDCTO.Text)

'validaciones
If Trim(tmpdcto) = "" Then Exit Sub
If Not IsNumeric(tmpdcto) Then Exit Sub

'aplicar cambios
rc.fields("DCTO") = CDbl(tmpdcto)
rc.UpdateBatch adAffectAll

DoEvents

rc.Move 0

   On Error GoTo 0
   Exit Sub

ioDCTO_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioDCTO_Click de Formulario frmPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : dtAgregar_Click
' Fecha/Hora  : 08/12/2003 11:32
' Autor       : JCASTILLO
' Propósito   : Agregar un nuevo artículo a la transferencia actual
'---------------------------------------------------------------------------------------
Private Sub dtAgregar_Click()

   On Error GoTo dtAgregar_Click_Error

   If rc.fields("ESTADO") > 0 Then
    lblstatus.Caption = "No se puede modificar la transferencia actual (ya no esta en creación)"
    Exit Sub
   End If
      
   
   lblstatus.Caption = ""
   
   With frmDetPtrans
   
    'ID = 0, para que entre añadiendo un nuevo registro
    .IR_A_ID = 0
    .CODIGO_TRANSF = rc.fields("CODIGO").Value
    .miCODALMORIG = rc.fields("CODALMORIG").Value
    .NUMERO_PEDIDO = rc.fields("NUMPED").Value
   
    'cargar textos de almacen origen y destino
    .lblORIGEN.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & AlmacenActual, locCnn))
    .lblDestino.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALMDEST.Text, locCnn))
    
    Me.WindowState = vbMinimized
    .Show
    'Me.WindowState = vbNormal
    
   End With

   'Call carga_grid_detalle(rc.Fields("CODIGO"), rc.Fields("CODALMORIG"))

   On Error GoTo 0
   Exit Sub

dtAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento dtAgregar_Click de Formulario frmPtrans"
   
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : fg_dblClick
' Fecha/Hora  : 25/11/2003 21:02
' Autor       : JCASTILLO
' Propósito   : Ir al registro seleccionado en el grid
'---------------------------------------------------------------------------------------
Private Sub fg_dblClick()
'
   On Error GoTo fg_dblClick_Error
   
   'si estado>0 no dejar modificar
   If rc.fields("ESTADO") > 1 Then
   
    lblstatus.Caption = "La Transferencia actual no se puede modificar"
    Exit Sub
    
   ElseIf rc.fields("ESTADO") = 1 Then
   
     'si no desea, salir
     If MsgBox("¿Desea modificar esta trasnferencia?, esta marcada como PENDIENTE", vbQuestion + vbYesNo) = vbNo Then Exit Sub
          
   End If
   
   lblstatus.Caption = ""
   'si no seleccionamos ninguna fila, salir
   If fg.TextMatrix(fg.Row, 1) = "" Or fg.TextMatrix(fg.Row, 1) = "ID" Then Exit Sub
   
   lblstatus.Caption = ""
   
   With frmDetPtrans
   
    .ESTADO_TRANSF = rc.fields("ESTADO")
    .IR_A_ID = fg.TextMatrix(fg.Row, 1)
    .CODIGO_TRANSF = rc.fields("CODIGO").Value
    .miCODALMORIG = rc.fields("CODALMORIG").Value
    .NUMERO_PEDIDO = rc.fields("NUMPED").Value
   
    'cargar textos de almacen origen y destino
    .lblORIGEN.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & AlmacenActual, locCnn))
    .lblDestino.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALMDEST.Text, locCnn))
    
    .Show
   End With

   Call carga_grid_detalle(rc.fields("CODIGO"), rc.fields("CODALMORIG").Value)
   
   

   On Error GoTo 0
   Exit Sub

fg_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_dblClick de Formulario frmPtrans"

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()

If Not prime Then
  
  If TipoPermiso = 1 Then cmdGenerarCodigo.Enabled = True

  If rc.RecordCount = 0 Then
        
        'If MsgBox("No se encuentran Transferencias. ¿Crear?", vbYesNo + vbQuestion, "Transferencias") = vbNo Then
        'Unload Me
        'Else
        'Call cbAgregar_Click
        'End If
        
  Else
        If añade_nueva Then
            Call cbAgregar_Click
        Else
            Call cmdFirst_Click
            Call cbCancelar_Click
        End If
        
  End If

prime = True
End If
    
End Sub

Private Sub Form_Load()


  If Dir("C:\TRANSFERENCIAS\", vbDirectory) = "" Then
    MkDir "C:\TRANSFERENCIAS"
  End If
  
  If Dir("C:\TRANSFERENCIAS\COMPTRN\", vbDirectory) = "" Then
    MkDir "C:\TRANSFERENCIAS\COMPTRN"
  End If
   
  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "PTRANS"
  oSQL.AddOrderClause "FMODI", True
  oSQL.AddOrderClause "CODALMORIG", False
  oSQL.AddOrderClause "CODIGO", True
  
  'estado = 0 -> en creación (no mostrar todavia en los puestos)
  'estado = 1 -> pendiente
  'estado = 2 -> aceptada
  'estado = 3 -> cancelada
  
  Select Case TipoPermiso
  
  Case 0 'dependiente comun (restringir solo a las transferencias en las
  'que participe el almacén actual
  
   
    'oSQL.AddSimpleWhereClause "CODALMORIG", AlmacenActual
    'oSQL.AddSimpleWhereClause "ESTADO", 0, , , LOGIC_AND
    'oSQL.AddSimpleWhereClause "ESTADO", 1, , , LOGIC_OR
    
    
    'que vean las de estado 0 y 1 para su almacen (cuando sea su alm. el origen)
    oSQL.AddComplexWhereClause "CODALMORIG = " & AlmacenActual & " AND (ESTADO = 0 OR ESTADO = 1)"
    'que vean las de estado 1 para su almacen como destino
    oSQL.AddComplexWhereClause "CODALMDEST = " & AlmacenActual & " AND ESTADO = 1", LOGIC_OR
    
    
    'oSQL.AddSimpleWhereClause "CODALMDEST", AlmacenActual, , , LOGIC_OR
    'oSQL.AddSimpleWhereClause "ESTADO", 1, , , LOGIC_AND
    
  Case 1 'supervisor
  
    oSQL.AddSimpleWhereClause "ESTADO", 0
    oSQL.AddSimpleWhereClause "ESTADO", 1, , , LOGIC_OR
  
  End Select

  'seleccionar las 50 primeras transferencias
  rc.Open "SELECT TOP 50 " & Right(oSQL.SQL, Len(oSQL.SQL) - 7), locCnn, adOpenStatic, adLockOptimistic
    
''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
End With
      
With cbCODALMDEST
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "CODALMDEST"
    .carga
    Set .DataSource = rc
End With

With ioDCTO
  Set .DataSource = rc
        .SoloNumeros = True
        .LongMaxima = 3
        .Alineacion = 1
        .DataField = "DCTO"
End With

With ioNUMPED
  Set .DataSource = rc
        .SoloNumeros = True
        .LongMaxima = 10
        .Alineacion = 1
        .DataField = "NUMPED"
End With

With ioGASTOS
    .dspFormat = "Currency"
    .Alineacion = 1
    .SoloNumeros = True
    .LongMaxima = 10
End With

fg.Cols = 0
  
With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
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
      
   ' Case vbKeyEnd
   '   cmdLast_Click
   ' Case vbKeyHome
   '   cmdFirst_Click
   ' Case vbKeyUp, vbKeyPageUp
   '   If Shift = vbCtrlMask Then
   '     cmdFirst_Click
   ' '  Else
   '     cmdPrevious_Click
    '  End If
      
   ' Case vbKeyDown, vbKeyPageDown
   '   If Shift = vbCtrlMask Then
   '     cmdLast_Click
    '  Else
    '    cmdNext_Click
   '   End If
      
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
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

If rc.State = 1 Then rc.Close
Set rc = Nothing

   With locCnn
    If .State = 1 Then .Close
   End With
   
Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmPtrans = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : ioDCTO_Click
' Fecha/Hora  : 29/01/2004 23:19
' Autor       : JCASTILLO
' Propósito   : Cambiar el descuento para esta transferencia
'---------------------------------------------------------------------------------------
Private Sub ioDCTO_Validate(Cancel As Boolean)
If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub

If ioDCTO.Text = "" Then ioDCTO.Text = "0"
If Not IsNumeric(ioDCTO.Text) Then Exit Sub

If ioDCTO.Text > 100 Then
    lblstatus.Caption = "DCTO Incorrecto (mayor de 100%)"
    ioDCTO.SetFocus
    ioDCTO.CancelarValidacion
    Cancel = True
    Exit Sub
End If
 
If cbCODALMDEST.Text = "" Then Exit Sub

'comprobar q no sea origen y destino iguales
If (cbCODALMDEST.Text = AlmacenActual) Then

 lblstatus.Caption = "ORIGEN y DESTINO deben ser diferentes"
 cbCODALMDEST.SetFocus
 'Cancel = True
 Exit Sub

End If

DoEvents

'actualizar y agregar un nuevo registro
Call cbactualizar_Click
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : comprueba_estado_pedido
' Fecha/Hora    : 09/03/2004 17:43
' Autor         : JCastillo
' Propósito     :  Comprueba el estado del pedido, devuelve FALSE si no hay ningun
'                      problema, y TRUE si hay algun error
'---------------------------------------------------------------------------------------
'
Private Function comprueba_estado_pedido() As Boolean
Dim tmpestado As Variant

   On Error GoTo comprueba_estado_pedido_Error

tmpestado = devuelve_campo("SELECT ESTADO FROM CABPEDPRO WHERE NUMERO = " & ioNUMPED.Text & " AND ALMORIG = " & AlmacenActual, locCnn)

If tmpestado = "@" Then
    MsgBox "El pedido no existe en la base de datos", vbInformation, titulo
    comprueba_estado_pedido = True
    Exit Function
End If

Select Case tmpestado
    
    Case Is < 3
    
        MsgBox "El pedido no se encuentra en histórico, imposible asignar a esta transferencia.", vbExclamation, titulo
        comprueba_estado_pedido = True
        Exit Function
    
    'Case 4
    '    MsgBox "El pedido ya ha sido transferido, imposible asignar a esta transferencia.", vbExclamation, titulo
    '    comprueba_estado_pedido = True
    '    Exit Function
                
End Select

   On Error GoTo 0
   Exit Function

comprueba_estado_pedido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_estado_pedido de Formulario frmPtrans"

End Function




'=======================================================================================================================================================================
Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo rc_MoveComplete_Error

  'Esto mostrará la posición de registro actual (para) este RecorDseT
  If rc.AbsolutePosition > 0 Then
  
    If cargando_Grid Then Exit Sub
    
    lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
        DoEvents
        
        fg.Clear
        
        Call carga_grid_detalle(rc.fields("CODIGO"), rc.fields("CODALMORIG"))
        If rc.fields("CODALMORIG") > 0 Then
            lblAlmOrig.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALMORIG")))
        Else
            lblAlmOrig.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & AlmacenActual))
        End If
        
        'si la transferencia ya esta pendiente o etc. ocultar el
        'boton de terminar transferencia y/o el de anular transferencia
        Select Case rc.fields("ESTADO").Value
        'si esta en creacion, mostrar ambos botones
        Case 0
            cmTerminarTrans.Visible = True
            cmAnularTrans.Visible = True
            cmdAceptarTrans.Visible = False
        'si esta pendiente ,ocultar el boton de terminar transferencia
        'y mostrar anular transferencia
        Case 1
            cmTerminarTrans.Visible = False
            cmAnularTrans.Visible = True
            
            'si es para el almacen actual, mostrar el botón de
            'aceptar transferencia
            If rc.fields("CODALMDEST") = AlmacenActual Then
               cmdAceptarTrans.Visible = True
            Else
               cmdAceptarTrans.Visible = False
            End If
            
        'si esta aceptada o cancelada ocultar el boton
        'de anular transferencia y terminar
        Case Is > 1
            cmTerminarTrans.Visible = False
            cmAnularTrans.Visible = False
            cmdAceptarTrans.Visible = False
        End Select
        
        'Mostrar la cantidad de mensajes asociados a la transferencia
        'actual
       'If CStr(rc.Fields("CODIGO")) <> "" Then lblMensajes.Caption = "Mensajes: " & devuelve_campo("SELECT COUNT(CODIGO) FROM PTRANSMSG WHERE CODIGO = " & rc.Fields("CODIGO"), locCnn)
       
       If CStr(rc.fields("CODIGO")) <> "" Then lblMensajes.Caption = "Mensajes: " & devuelve_campo("SELECT COUNT(CODIGO) FROM PTRANSMSG WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODALMORIG = " & rc.fields("CODALMORIG"), locCnn)
        'mostrar el estado de la transferencia
        Select Case rc.fields("ESTADO").Value
        
        Case 0
            lblEstado.Caption = "EN CREACION"
        Case 1
            lblEstado.Caption = "PENDIENTE"
        Case 2
            lblEstado.Caption = "ACEPTADA"
        Case 3
            lblEstado.Caption = "CANCELADA"
        
       End Select
       
        ioGASTOS.Text = rc.fields("GASTOS")

  Else
  
        ioGASTOS.Text = ""
  
  End If

   On Error GoTo 0
   Exit Sub

rc_MoveComplete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento rc_MoveComplete de Formulario frmPtrans"
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
  Dim tmpcodigo As Variant
  
   On Error GoTo cbAgregar_Click_Error

  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(CODIGO) + 1 from PTRANS where CODALMORIG = " & AlmacenActual)
    
    'cbCODALMORIG.Text = AlmacenActual
    rc.fields("CODALMORIG") = AlmacenActual
    'cbCODALMORIG.Locked = True
    'cbCODALMORIG.Enabled = False
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    
  'ioDescripcion.SetFocus
  End With

DoEvents

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'no funciona cbcodalorig.setfocus ¿?
cbCODALMDEST.Locked = False
DoEvents
cbCODALMDEST.SetFocus
'SendKeys "{UP}"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAgregar_Click de Formulario frmPtrans"

End Sub

''Private Sub cbeliminar_Click()
'    On Error GoTo DeleteErr
' With rc
'    '.Delete
'    '.MoveNext
'    .Fields("mbaja") = True
'   .Fields("FBAJA") = Date
'   If .EOF Then .MoveLast
' End With
'
' Call cbactualizar_Click
'
'
''Exit Sub
'DeleteErr:
'  MsgBox Err.Description
'End Sub

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
  
   On Error GoTo cbactualizar_Click_Error

  'validaciones
  If ioDCTO.Text = "" Then ioDCTO.Text = "0"
  If ioGASTOS.Text = "" Then ioGASTOS.Text = "0"
  
  'comprobación de numero de pedido
  If ioNUMPED.Text = "" Then
        lblstatus.Caption = "Debe especificar un numero de pedido."
        ioNUMPED.SetFocus
        ioNUMPED.CancelarValidacion
        Exit Sub
  End If
  
  If ioNUMPED.Text <> "0" Then
    'comprobar que el pedido no haya sido transferido
    If comprueba_estado_pedido = True Then Exit Sub
  End If
  
  If ioDCTO.Text > 100 Then
    lblstatus.Caption = "DCTO Incorrecto (mayor de 100%)"
    ioDCTO.SetFocus
    ioDCTO.CancelarValidacion
    Exit Sub
  End If
  
  'If cbCODALMORIG.Text = "" Then
  '  lblstatus.Caption = "Almacen de origen no puede estar en blanco"
  '  cbCODALMORIG.SetFocus
  '  Exit Sub
  'End If

  If cbCODALMDEST.Text = "" Then
    lblstatus.Caption = "Almacen de destino no puede estar en blanco"
    cbCODALMDEST.SetFocus
    Exit Sub
  End If
  
  If cbCODALMDEST.Text = AlmacenActual Then
    lblstatus.Caption = "Almacen de ORIGEN y DESTINO no pueden ser iguales"
    cbCODALMDEST.SetFocus
    Exit Sub
  End If
 
  rc.fields("GASTOS") = Replace(ioGASTOS.Text, ",", ".")
  rc.fields("CODUSR") = UsuarioActual
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
    Call dtAgregar_Click     'agregar en detalle
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  
  
  'If ioULTIMAS.Text <> "" Then carga_grid (CLng(ioULTIMAS.Text))
   
  DoEvents

  
  'Call cbAgregar_Click

   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario frmPtrans"

End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  If Not rc.BOF Then rc.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  If Not rc.EOF Then rc.MoveLast
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
  cbActualizar.Visible = Not bVal
  cbCerrar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cbLista.Visible = bVal
  dtAgregar.Visible = bVal
  cmdGenerarCodigo.Visible = bVal
  
    
'  cbCODALMORIG.Enabled = Not bVal
  cbCODALMDEST.Enabled = Not bVal
  ioDCTO.Locked = bVal
    
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid_detalle
' Fecha/Hora     : 05/12/2003 12:50
' Autor             : JCastillo
' Propósito       : Carga los registros de detalle de transferencias, correspondientes
'                       al registro actual de cabecera
'---------------------------------------------------------------------------------------
Public Sub carga_grid_detalle(numtrans As Double, codalmorig As Byte)
Dim tmprc As New ADODB.Recordset
Dim tmplinea As Long
Dim tmpcodcolor As Long
Dim t_articulo As Variant
Dim tmpprecom As Single
Dim total_importe As Double  '(precom * unidades)
Dim var As Byte
Dim tmpdcto As Double

   On Error GoTo carga_grid_detalle_Error

    If cargando_Grid Then Exit Sub
    
    lblstatus.Caption = "Cargando Rejilla, espere ..."
    DoEvents
    
    tmprc.Open "SELECT ID, CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES, FMODI FROM DETTRANS WHERE CODIGO = " & numtrans & " AND CODALM = " & codalmorig & " ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    
    With fg
        
        .Redraw = flexRDNone
        .Clear
        .AddItem "Cargando datos ..."
        .Cols = 13
        .ColFormat(7) = "Currency"
        .ColFormat(8) = "Currency"
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColHidden(0) = True
        .ColHidden(1) = True
        
        'poner títulos
        '.TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "PROV."
        .TextMatrix(0, 3) = "REF"
        .TextMatrix(0, 4) = "MODELO"
        .TextMatrix(0, 5) = "TALLA"
        .TextMatrix(0, 6) = "COLOR"
        .TextMatrix(0, 7) = "P.COM"
        .TextMatrix(0, 8) = "P.VEN"
        .TextMatrix(0, 9) = "UDS."
        .TextMatrix(0, 10) = "TEMP."
        .TextMatrix(0, 11) = "FECHA"
        .TextMatrix(0, 12) = "CBARRAS"
        
        .Rows = 1
    
    Do Until tmprc.EOF

        .Rows = .Rows + 1
        
       
        'sacar datos del artículo
        t_articulo = devuelve_matriz("SELECT MODELO, PREVEN, CODPROV, REF, PRECOM FROM MAARTIC WHERE CODIGO = " & tmprc.fields("CODART").Value & " AND TEMPOR = " & tmprc.fields("TEMPOR"), locCnn)
        
         'numero de linea
        'tmpprecom = Obtiene_Precom_Pedido(tmprc.Fields("CODART").Value, tmprc.Fields("TEMPOR").Value, tmprc.Fields("CODTALLA").Value, tmprc.Fields("CODCOL").Value, locCnn)
                
        'condición de error
        If Not IsArray(t_articulo) Then
        
            lblstatus.Caption = "Error al cargar el GRID"
            .TextMatrix(.Rows - 1, 4) = "** Revise el Código de Artículo **"
            'Exit Sub
        
        Else
               
        
        tmpprecom = t_articulo(4)
        
        .TextMatrix(.Rows - 1, 1) = tmprc.fields("ID")
        .TextMatrix(.Rows - 1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(2), locCnn))
        
        'referencia
        .TextMatrix(.Rows - 1, 3) = t_articulo(3)
        
        .TextMatrix(.Rows - 1, 4) = Format(tmprc.fields("CODART"), "00000") & " " & Trim(t_articulo(0))
        .TextMatrix(.Rows - 1, 5) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & tmprc.fields("CODTALLA").Value, locCnn))
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
       If Not IsNull(tmprc.fields("CODCOL")) And tmprc.fields("CODCOL") <> 0 Then
      
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & tmprc.fields("CODCOL"), locCnn)
            .TextMatrix(.Rows - 1, 6) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & tmprc.fields("CODCOL"), locCnn)
            .Col = 6
            .Row = .Rows - 1
            .CellBackColor = tmpcodcolor
            .Col = 2
        
       End If
       
        'PRECOM
        .TextMatrix(.Rows - 1, 7) = tmpprecom
        'PVP
        .TextMatrix(.Rows - 1, 8) = t_articulo(1)
                
        'UDS
        .TextMatrix(.Rows - 1, 9) = tmprc.fields("UNIDADES").Value
              
        .TextMatrix(.Rows - 1, 10) = devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & tmprc.fields("TEMPOR").Value, locCnn)
        
        .TextMatrix(.Rows - 1, 11) = tmprc.fields("FMODI").Value
        
        .TextMatrix(.Rows - 1, 12) = Conforma_CB(tmprc.fields("CODART"), tmprc.fields("TEMPOR"), tmprc.fields("CODTALLA"), tmprc.fields("CODCOL"))
        
        'acumulando ... importe * unidades
        total_importe = total_importe + (tmpprecom * tmprc.fields("UNIDADES").Value)
             
        
    End If
    tmprc.MoveNext
    
    Loop
    
   If rc.fields("DCTO") > 0 Then tmpdcto = rc.fields("DCTO").Value
  
    If tmprc.RecordCount > 0 Then
    
    .SubtotalPosition = flexSTAbove
    .subtotal flexSTSum, , 9, , vbBlue, vbWhite, True
    .TextMatrix(1, 8) = "Total Uds:"
    .TextMatrix(1, 3) = "Total (" & tmprc.RecordCount & ") Art."
    .TextMatrix(1, 4) = "Subtotal: " & Format(total_importe, "Currency") & ".  Total: " & Format(total_importe - ((total_importe * tmpdcto) / 100), "Currency")
    .TextMatrix(1, 1) = ""
    
    End If
    

    
    .AutoSize 1, .Cols - 1
    .Redraw = True
    '.Enabled = True
    End With
       
    
    If total_importe > 0 And total_importe <> rc.fields("TOTAL") Then
        rc.fields("TOTAL") = total_importe
        rc.UpdateBatch adAffectAll
    End If
    
    
    tmprc.Close
    Set tmprc = Nothing
    
    
    lblstatus.Caption = ""
    DoEvents
    
        
   On Error GoTo 0
   Exit Sub

carga_grid_detalle_Error:

    lblstatus.Caption = ""
    DoEvents

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_detalle de Formulario frmPtrans"
 
End Sub


Private Sub carga_cabecera_grid_diferencias()
              
     With frmDifTrn.fg
            .Clear
            .Rows = 1
            .Cols = 9
            .ColFormat(7) = "Currency"
            .ColAlignment(2) = flexAlignCenterCenter
            .ColAlignment(4) = flexAlignCenterCenter
            .TextMatrix(0, 1) = "Proveedor"
            .TextMatrix(0, 2) = "Ref."
            .TextMatrix(0, 3) = "Modelo"
            .TextMatrix(0, 4) = "Talla"
            .TextMatrix(0, 5) = "Color"
            .TextMatrix(0, 6) = "Uds."
            .TextMatrix(0, 7) = "Pre.Com."
            .TextMatrix(0, 8) = "CBarras"
     End With
     
     With frmDifTrn
        .NUMERO_PEDIDO = rc.fields("NUMPED")
        .CODIGO_TRANSF = rc.fields("CODIGO")
    End With
   
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : comprueba_pedido
' Fecha/Hora     : 10/03/2004 09:15
' Autor             : JCastillo
' Propósito       : Comprobar que no faltan unidades del pedido en las transferencias
'                       asignadas a ese pedido, antes de dar paso y ponerla como estado
'                       pendiente.
'                           1  Todo correcto
'                           2  Faltan Unidades
'                           3  No se encuentra pedido
'---------------------------------------------------------------------------------------
Private Function comprueba_pedido(gNumped) As Byte
Dim difvisible As Boolean
Dim t_articulo As Variant
Dim tmpprecom As Variant
Dim tmpcodcolor As Variant
Dim total As Currency

'Dim dif_cargadas As Boolean

Dim Puds As Variant
Dim Tuds As Variant

Dim rc As New ADODB.Recordset
Dim rct As New ADODB.Recordset

   On Error GoTo comprueba_unidades_pedido_Error

    'abrir los registros del pedido
    rc.Open "SELECT SUM(UNIDADES), CODART, TEMPOR, CODTALLA, CODCOL FROM DETPEDPRO WHERE NUMERO = " & gNumped & " AND ALMORIG = " & AlmacenActual & " GROUP BY CODART, TEMPOR, CODTALLA, CODCOL", locCnn, adOpenStatic, adLockReadOnly
      
    'si no hay registros en el pedido, salir
    If rc.RecordCount <= 0 Then
        comprueba_pedido = 3
        Exit Function
    End If
    
    Do Until rc.EOF
    
        'abrir los registros de detalle de transferencias (de todas las transferencias pendientes
        'que estan asignadas a este numero de pedido
        If rct.State = 1 Then rct.Close
        rct.Open "SELECT sum(UNIDADES) FROM DETTRANS WHERE (CAST(CODIGO AS VARCHAR(10)) + CAST(CODALM AS VARCHAR(3))) IN (SELECT CAST(CODIGO AS VARCHAR(10)) + CAST(CODALMORIG AS VARCHAR(3)) FROM PTRANS WHERE ESTADO <> 3 AND NUMPED = " & gNumped & " and CODALMORIG = " & AlmacenActual & ") AND CODART = " & rc.fields("CODART") & " AND TEMPOR = " & rc.fields("TEMPOR") & " AND CODTALLA = " & rc.fields("CODTALLA") & " AND CODCOL = " & rc.fields("CODCOL"), locCnn, adOpenStatic, adLockReadOnly
                
        Tuds = rct.fields(0)
        Puds = rc.fields(0)
        
        If IsNull(Puds) Then Puds = 0
        If IsNull(Tuds) Then Tuds = 0
        
        'si no hay registros, directamente guardar la diferencia:
        If ((rc.RecordCount <= 0) Or (Puds > Tuds)) Then
                
            comprueba_pedido = 2
            
            If Not difvisible Then
                difvisible = True
                        
            'If Not dif_cargadas Then
                Call carga_cabecera_grid_diferencias
            'dif_cargadas = True
            'End If
                With frmDifTrn
                    .Show
                End With
            End If
            
            t_articulo = devuelve_matriz("SELECT MODELO, PREVEN, CODPROV, REF, PRECOM FROM MAARTIC WHERE CODIGO = " & rc.fields("CODART").Value & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnn)
            
            With frmDifTrn.fg
                .AddItem "", 1
                
                 'numero de linea
                 'tmpprecom = Obtiene_Precom_Pedido(tmprc.Fields("CODART").Value, tmprc.Fields("TEMPOR").Value, tmprc.Fields("CODTALLA").Value, tmprc.Fields("CODCOL").Value, locCnn)
                 tmpprecom = t_articulo(4)
                 
                'proveedor
                .TextMatrix(1, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(2), locCnn))
        
                'referencia
                .TextMatrix(1, 2) = t_articulo(3)
    
                .TextMatrix(1, 3) = Format(rc.fields("CODART"), "00000") & " " & Trim(t_articulo(0))
                .TextMatrix(1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA").Value, locCnn))
        
                'obtener el texto del color y su codigo de color (para colorear
                'la celda del grid)
                If Not IsNull(rc.fields("CODCOL")) And rc.fields("CODCOL") <> 0 Then
      
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"), locCnn)
                .TextMatrix(1, 5) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"), locCnn)
                .Col = 5
                .Row = 1
                .CellBackColor = tmpcodcolor
                .Col = 2
                
                End If
       
                'UNIDADES
                If rc.RecordCount <= 0 Then
                    .TextMatrix(1, 6) = Puds
                    total = total + (Puds * tmpprecom)
                Else
                    .TextMatrix(1, 6) = Puds - Tuds
                    total = total + ((Puds - Tuds) * tmpprecom)
                End If
       
                'PRECOM
                .TextMatrix(1, 7) = tmpprecom
                
                'CODIGO DE BARRAS
                .TextMatrix(1, 8) = Conforma_CB(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"))
                                    
            End With
                
        End If
        
        rc.MoveNext
    
    Loop
    
    If difvisible Then
    
    With frmDifTrn.fg
            .SubtotalPosition = flexSTAbove
            .subtotal flexSTSum, , 6, , vbBlue, vbWhite, True
            .TextMatrix(1, 6) = "Uds: " & .TextMatrix(1, 6)
            .TextMatrix(1, 3) = "Total: " & Format(total, "Currency")
            .TextMatrix(1, 1) = ""
            .AutoSize 1, .Cols - 1
    End With
    
    End If
    
    If comprueba_pedido = 0 Then comprueba_pedido = 1
  
   On Error GoTo 0
   Exit Function

comprueba_unidades_pedido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_unidades_pedido de Formulario frmDetPtrans"
End Function





