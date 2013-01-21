VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMntArt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11025
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
   ScaleHeight     =   6915
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   3675
      Left            =   15
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "FrmMntArt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioTARIFA"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioPRECOM"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ioPREVEN"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbSECCION"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioREF"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ioABREVIA"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbSUBFAM"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbFAMILIA"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ioMODELO"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "vsModificaciones"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "FrmMntArt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsTarifas"
      Tab(1).Control(1)=   "ioDCTO"
      Tab(1).Control(2)=   "Label23"
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(4)=   "lblPRECOM"
      Tab(1).Control(5)=   "lblPrecomFin"
      Tab(1).Control(6)=   "Label15"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Hoja 3"
      TabPicture(2)   =   "FrmMntArt.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ioSTOCK"
      Tab(2).Control(1)=   "Label20"
      Tab(2).Control(2)=   "Label19"
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(6)=   "Label9"
      Tab(2).Control(7)=   "ioIVACOM"
      Tab(2).Control(8)=   "ioFOTO"
      Tab(2).Control(9)=   "ioPEDIR"
      Tab(2).Control(10)=   "ioSTOCKMAX"
      Tab(2).Control(11)=   "ioSTOCKMIN"
      Tab(2).Control(12)=   "lblExisteCom"
      Tab(2).Control(13)=   "cmComentario"
      Tab(2).Control(14)=   "Dialogo"
      Tab(2).Control(15)=   "cbTIPOIVA"
      Tab(2).Control(16)=   "cbCODPROV"
      Tab(2).Control(17)=   "cmBorrarFoto"
      Tab(2).ControlCount=   18
      Begin VSFlex8Ctl.VSFlexGrid vsModificaciones 
         Height          =   2520
         Left            =   6510
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1125
         Width           =   4455
         _cx             =   7858
         _cy             =   4445
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMntArt.frx":0054
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
         DataMode        =   0
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
      Begin PCGestion.miText ioMODELO 
         Height          =   525
         Left            =   1350
         TabIndex        =   5
         Top             =   2280
         Width           =   5190
         _ExtentX        =   9155
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
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   540
         Left            =   1335
         TabIndex        =   3
         Top             =   1095
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   953
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
         Height          =   495
         Left            =   1335
         TabIndex        =   4
         Top             =   1695
         Width           =   5175
         _ExtentX        =   9128
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
      Begin PCGestion.miText ioABREVIA 
         Height          =   525
         Left            =   4125
         TabIndex        =   1
         Top             =   525
         Width           =   1695
         _ExtentX        =   2990
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
      Begin PCGestion.miText ioREF 
         Height          =   525
         Left            =   1350
         TabIndex        =   0
         Top             =   510
         Width           =   1965
         _ExtentX        =   3466
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
      Begin PCGestion.miCombo cbSECCION 
         Height          =   495
         Left            =   6630
         TabIndex        =   2
         Top             =   525
         Width           =   4320
         _ExtentX        =   7620
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
      Begin PCGestion.ucGrdBttn cmBorrarFoto 
         Height          =   315
         Left            =   -64365
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3300
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
         Image           =   "FrmMntArt.frx":00FF
      End
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   540
         Left            =   -73530
         TabIndex        =   15
         Top             =   2340
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   953
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
      Begin PCGestion.miCombo cbTIPOIVA 
         Height          =   510
         Left            =   -73545
         TabIndex        =   13
         Top             =   1725
         Width           =   3330
         _ExtentX        =   5794
         _ExtentY        =   900
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
      Begin VSFlex8Ctl.VSFlexGrid vsTarifas 
         Height          =   2280
         Left            =   -74985
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1335
         Width           =   10905
         _cx             =   19235
         _cy             =   4022
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMntArt.frx":011B
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
         DataMode        =   0
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
      Begin MSComDlg.CommonDialog Dialogo 
         Left            =   -70305
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin PCGestion.ucGrdBttn cmComentario 
         Height          =   375
         Left            =   -71025
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   661
         Caption         =   "Modificar Comentario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "FrmMntArt.frx":0214
      End
      Begin PCGestion.bsGradientLabel lblExisteCom 
         Height          =   405
         Left            =   -73515
         Top             =   2985
         Visible         =   0   'False
         Width           =   2430
         _ExtentX        =   4921
         _ExtentY        =   714
         Caption         =   "¡ Existe Comentario!"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   0
         Colour1         =   8454143
         Colour2         =   49152
         CaptionAlignment=   1
      End
      Begin PCGestion.miText ioPREVEN 
         Height          =   525
         Left            =   5115
         TabIndex        =   7
         Top             =   2925
         Width           =   1410
         _ExtentX        =   2487
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
      Begin PCGestion.miText ioDCTO 
         Height          =   525
         Left            =   -68895
         TabIndex        =   8
         Top             =   570
         Width           =   1020
         _ExtentX        =   1799
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
      Begin PCGestion.miText ioSTOCKMIN 
         Height          =   525
         Left            =   -73515
         TabIndex        =   10
         Top             =   1065
         Width           =   1170
         _ExtentX        =   2064
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
      Begin PCGestion.miText ioSTOCKMAX 
         Height          =   525
         Left            =   -71385
         TabIndex        =   11
         Top             =   1065
         Width           =   1170
         _ExtentX        =   2064
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
      Begin PCGestion.miText ioPEDIR 
         Height          =   525
         Left            =   -68880
         TabIndex        =   12
         Top             =   1065
         Width           =   1170
         _ExtentX        =   2064
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
      Begin PCGestion.miText ioPRECOM 
         Height          =   525
         Left            =   1350
         TabIndex        =   6
         Top             =   2910
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
      Begin VB.PictureBox ioFOTO 
         Height          =   3075
         Left            =   -67545
         ScaleHeight     =   3015
         ScaleWidth      =   3375
         TabIndex        =   65
         Top             =   525
         Width           =   3435
      End
      Begin PCGestion.miText ioIVACOM 
         Height          =   525
         Left            =   -68880
         TabIndex        =   14
         Top             =   1725
         Width           =   1170
         _ExtentX        =   2064
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA COMPRA"
         Height          =   270
         Left            =   -70200
         TabIndex        =   67
         Top             =   1785
         Width           =   1290
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD A PEDIR"
         Height          =   645
         Left            =   -70005
         TabIndex        =   66
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO COMPRA"
         Height          =   645
         Left            =   -74040
         TabIndex        =   64
         Top             =   495
         Width           =   900
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DESCUENTO"
         Height          =   345
         Left            =   -70290
         TabIndex        =   63
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label lblPRECOM 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   -73080
         TabIndex        =   62
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblPrecomFin 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   -66465
         TabIndex        =   61
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO COMPRA FINAL"
         Height          =   885
         Left            =   -67710
         TabIndex        =   60
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK MINIMO"
         Height          =   645
         Left            =   -74445
         TabIndex        =   58
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK MAXIMO"
         Height          =   645
         Left            =   -72315
         TabIndex        =   57
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR"
         Height          =   330
         Left            =   -74985
         TabIndex        =   56
         Top             =   2430
         Width           =   1410
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA VENTA"
         Height          =   330
         Left            =   -74730
         TabIndex        =   55
         Top             =   1800
         Width           =   1155
      End
      Begin MSForms.CheckBox ioSTOCK 
         Height          =   450
         Left            =   -73515
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   1965
         VariousPropertyBits=   746588183
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3466;794"
         Value           =   "0"
         Caption         =   "TRATAR STOCK"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox ioTARIFA 
         Height          =   750
         Left            =   2940
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2805
         Width           =   1080
         VariousPropertyBits=   746596371
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1905;1323"
         Value           =   "0"
         Caption         =   "USAR TARIFA"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO VENTA"
         Height          =   645
         Left            =   4215
         TabIndex        =   47
         Top             =   2850
         Width           =   885
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO COMPRA"
         Height          =   645
         Left            =   390
         TabIndex        =   46
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECCION"
         Height          =   300
         Left            =   5745
         TabIndex        =   44
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIA"
         Height          =   360
         Left            =   120
         TabIndex        =   43
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ABREVIA"
         Height          =   360
         Left            =   3225
         TabIndex        =   42
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAMILIA"
         Height          =   300
         Left            =   135
         TabIndex        =   41
         Top             =   1770
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   330
         Left            =   465
         TabIndex        =   40
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         Height          =   360
         Left            =   405
         TabIndex        =   39
         Top             =   2340
         Width           =   915
      End
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5415
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
      MICON           =   "FrmMntArt.frx":0230
      PICN            =   "FrmMntArt.frx":024C
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
      Left            =   0
      Top             =   5025
      Width           =   10995
      _ExtentX        =   16722
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5415
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
      MICON           =   "FrmMntArt.frx":0F1E
      PICN            =   "FrmMntArt.frx":0F3A
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
      Left            =   5070
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5430
      Width           =   1110
      _ExtentX        =   1958
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
      MICON           =   "FrmMntArt.frx":1C70
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
      Left            =   8865
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5430
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
      MICON           =   "FrmMntArt.frx":1C8C
      PICN            =   "FrmMntArt.frx":1CA8
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
      Left            =   9930
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5430
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
      MICON           =   "FrmMntArt.frx":297A
      PICN            =   "FrmMntArt.frx":2996
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
      Left            =   45
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":36CC
      PICN            =   "FrmMntArt.frx":36E8
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
      Left            =   1140
      TabIndex        =   18
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":43C2
      PICN            =   "FrmMntArt.frx":43DE
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
      Left            =   2370
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":4CB8
      PICN            =   "FrmMntArt.frx":4CD4
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
      Left            =   7875
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":5532
      PICN            =   "FrmMntArt.frx":554E
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
      Left            =   8835
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":5E28
      PICN            =   "FrmMntArt.frx":5E44
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
      Left            =   9930
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6090
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
      MICON           =   "FrmMntArt.frx":6A16
      PICN            =   "FrmMntArt.frx":6A32
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo Temporada 
      Height          =   480
      Left            =   6300
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4530
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   847
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
   Begin PCGestion.bsGradientLabel lblTemActual 
      Height          =   405
      Left            =   2085
      Top             =   4560
      Width           =   2430
      _ExtentX        =   4921
      _ExtentY        =   714
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
      Colour1         =   16744576
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.ucGrdBttn cmCambTem 
      Height          =   420
      Left            =   10605
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4560
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   741
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmMntArt.frx":770C
   End
   Begin PCGestion.bsGradientLabel lblExisteCom2 
      Height          =   375
      Left            =   5385
      Top             =   15
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      Caption         =   "Comentario"
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
      Colour1         =   8454143
      Colour2         =   49152
      CaptionAlignment=   1
   End
   Begin VB.Label lblArticulo 
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
      Left            =   1875
      TabIndex        =   51
      Top             =   30
      Width           =   3450
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUAL"
      Height          =   330
      Left            =   1140
      TabIndex        =   49
      Top             =   4620
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPORADAS"
      Height          =   330
      Left            =   4890
      TabIndex        =   48
      Top             =   4620
      Width           =   1365
   End
   Begin MSForms.CheckBox ioHIST 
      Height          =   435
      Left            =   105
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4140
      Width           =   840
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1482;767"
      Value           =   "0"
      Caption         =   "HIST"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   1290
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4140
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
      Left            =   9750
      TabIndex        =   26
      Top             =   4155
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
      Left            =   6330
      TabIndex        =   25
      Top             =   4155
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
      Left            =   8550
      TabIndex        =   24
      Top             =   30
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
      Left            =   930
      TabIndex        =   23
      Top             =   30
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   45
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Alta"
      Height          =   315
      Left            =   8400
      TabIndex        =   21
      Top             =   4185
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   270
      Left            =   6450
      TabIndex        =   20
      Top             =   30
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Baja"
      Height          =   330
      Left            =   5055
      TabIndex        =   19
      Top             =   4200
      Width           =   1200
   End
End
Attribute VB_Name = "FrmMntArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmMntArt
' DateTime  : 06/11/2003 10:10
' Author    : Administrador
' Purpose   : Mantenimiento de Artículos
'
' Si entra como add_en_detalle = TRUE, quiere decirse que las adicciones de
' registros es decir, registros nuevos, deben añadirse al recordset de detalle
' asi como en modificiaciones de articulos, si ese articulo esta también en el
' recordset de detalle, pues modificar. El funcionamiento del formulario seria
' igual, salvo por que ira grabando los datos correspondientes en rc_detalle.
'---------------------------------------------------------------------------------------
Option Explicit

Dim WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim add_ref As Boolean
Dim prime As Boolean
Dim oSQL As New clsSmartSQL

'para contar las lineas del vsmodificaciones
Dim tmplinea As Long
Dim Iva_Descontado As Boolean


'para guardar los valores de los micombo en el momento del gotfocus
'y comprobar en el validate con el valor actual para saber si se ha
'cambiado o no, y si ha cambiado llamar a la rutina para actualizar
'los otros combos.
Dim tmpseccion As String
Dim tmpfamilia As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public add_en_detalle As Boolean
Public rc_detalle As New ADODB.Recordset
Public NumeroPedido As Long
Public ProveedPedido As Long
Public TemporadaWrk As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cbCODPROV_Validate(Cancel As Boolean)
Tab1.Tab = 0
End Sub

Private Sub cbFamilia_GotFocus()
tmpfamilia = cbFAMILIA.Text
End Sub

Private Sub cbSECCION_GotFocus()
tmpseccion = cbSECCION.Text
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'---------------------------------------------------------------------------------------
' Procedure : añade_en_detalle
' DateTime  : 06/11/2003 10:35
' Author    : Administrador
' Purpose   : Añadir en el rc_detalle el registro correspondiente al ultimo
'                 movimiento.
'---------------------------------------------------------------------------------------
Private Sub añade_en_detalle()
Dim tmpcodigo As Variant

   On Error GoTo añade_en_detalle_Error

'si esta editando preguntar si desea añadir también el articulo editado
If mbAddNewFlag = False And mbEditFlag = True Then

    If MsgBox("¿Desea añadir también este artículo:" & ioCODIGO.Caption & "  " & ioMODELO.Text & " al pedido?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

'si es un articulo de nueva creación, añadir en el momento
ElseIf (mbAddNewFlag) Or add_ref Then

With rc_detalle

    .AddNew
    
    tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETPEDPRO where NUMERO = " & NumeroPedido)
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("LINEA") = tmpcodigo
    
    .fields("NUMERO") = NumeroPedido
    .fields("CODART") = CSng(ioCODIGO.Caption)
    .fields("TEMPOR") = TemporadaWrk
    .fields("ALMORIG") = AlmacenActual
    .fields("PRECOM") = CSng(ioPRECOM.Text)
    .fields("CODTALLA") = 0 '
    .fields("CODCOL") = 0
    
    'meter las unidades por defecto
     If ioPEDIR.Text <> "" Then
        .fields("UNIDADES") = CSng(ioPEDIR.Text)
     Else
         .fields("UNIDADES") = 0
     End If
     
     'meter los descuentos
     'If ioDCTO1.Text <> "" Then .Fields("DCTO1") = ioDCTO1.Text
     If ioDCTO.Text <> "" Then
        .fields("DCTO") = ioDCTO.Text
     Else
        .fields("DCTO") = 0
     End If
     
     'meter los % para iva y re
     If cbTIPOIVA.Text <> "" Then
        .fields("IVA").Value = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & cbTIPOIVA.Text)
        .fields("RE").Value = devuelve_campo("SELECT RE FROM IVA WHERE CODIGO = " & cbTIPOIVA.Text)
     Else
       .fields("IVA").Value = 0
       .fields("RE").Value = 0
     End If
     
     .UpdateBatch
     
      frmPedProv.Linea_Creada = tmpcodigo
      Set tmpcodigo = Nothing
    
     Call frmPedProv.refresca_grid_externo(True)
     
End With

End If

   On Error GoTo 0
   Exit Sub

añade_en_detalle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure añade_en_detalle of Formulario FrmMntArt"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cbSECCION_Validate(Cancel As Boolean)

If Trim(cbSECCION.Text <> "" And cbSECCION.Locked = False) Then

    'si son iguales, salir
    If tmpseccion = cbSECCION.Text Then
        tmpseccion = ""
        Exit Sub
    End If

    With cbFAMILIA
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE CODSEC = " & cbSECCION.Text & " AND MBAJA = 0 ORDER BY CODIGO"
        .LenCodigo = 5
        .carga
    End With
    
    
    cbSUBFAM.borra_combo
    cbSUBFAM.Text = ""

'buscar como sección NULO
ElseIf cbSECCION.Text = "" And cbFAMILIA.Text = "" Then

    'si son iguales, salir
    'If tmpseccion = cbSECCION.Text Then
    '    tmpseccion = ""
    '    Exit Sub
    'End If
    
     With cbFAMILIA
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE CODSEC IS NULL AND MBAJA = 0 ORDER BY CODIGO"
        .LenCodigo = 5
        .carga
    End With
    
    
End If

    tmpseccion = ""
End Sub



Private Sub cbfamilia_Validate(Cancel As Boolean)

If Trim(cbFAMILIA.Text <> "" And cbFAMILIA.Locked = False) Then

    'si son iguales, salir
    If tmpfamilia = cbFAMILIA.Text Then
        tmpfamilia = ""
        Exit Sub
    End If
    
With cbSUBFAM
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE CODFAM = " & cbFAMILIA.Text & " AND MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .carga
End With
End If

tmpfamilia = ""
End Sub







'---------------------------------------------------------------------------------------
' Subrutina   : cbTIPOIVA_Validate
' Fecha/Hora  : 27/11/2003 21:36
' Autor       : JCASTILLO
' Propósito   : Si estamos llamando al formulario desde pedidos u otro,
'               realizar la inserción continua de registros hasta que se pulse sobre el
'               botón de terminar edición
'---------------------------------------------------------------------------------------
Private Sub cbTIPOIVA_Validate(Cancel As Boolean)

If add_en_detalle Then

    Call cbactualizar_Click
   ' DoEvents
    
    'Tab1.Tab = 0
    'Tab1.SetFocus
    
   ' Call cbAgregar_Click
    
    'cbSECCION.SetFocus
  '  SendKeys "{TAB}"
    
End If

End Sub

Private Sub cmBorrarFoto_Click()
If mbEditFlag = False And mbAddNewFlag = False Then Exit Sub

If MsgBox("¿Desea quitar la imagen?", vbQuestion + vbYesNo, titulo) = vbYes Then

    rc.fields(ioFOTO.DataField).Value = Null
    ioFOTO.Picture = Nothing

End If
End Sub

Private Sub cmCambTem_Click()

On Error GoTo Temporada_Validate_Error

If Temporada.Text <> "" Then
    
    Call des_enlaza_campos
    
    oSQL.ClearWhereClause
    oSQL.AddSimpleWhereClause "TEMPOR", CByte(Temporada.Text)
    oSQL.AddSimpleWhereClause "HIST", "0"
    oSQL.AddSimpleWhereClause "MBAJA", "0"
    
    rc.Close
    rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    
    'si no hay registros, borrar temporada
    If rc.EOF Then
    
        If add_en_detalle Then
         Temporada.Text = TemporadaWrk
        Else
         Temporada.Text = TemporadaActual
        End If
        
        oSQL.ClearWhereClause
        oSQL.AddSimpleWhereClause "HIST", "0"
        oSQL.AddSimpleWhereClause "MBAJA", "0"
        oSQL.AddSimpleWhereClause "TEMPOR", CByte(Temporada.Text)
        
        rc.Close
        rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
                
        DoEvents
        
        On Error Resume Next
        rc.Requery
        On Error GoTo Temporada_Validate_Error
        
        MsgBox "No se encuentran artículos para esa temporada", vbExclamation
    
    End If
    
        lblTemActual.Caption = devuelve_campo("SELECT AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & Temporada.Text)
    
    Call enlaza_campos
    DoEvents
    
End If

   On Error GoTo 0
   Exit Sub

Temporada_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Temporada_Validate of Formulario FrmMntArt"

End Sub

Private Sub cmComentario_Click()

FrmInicio.Editor.carga "Comentario de Artículos [" & ioMODELO.Text & "]", rc.fields("COMEN"), ""

End Sub

Private Sub Form_Activate()
If Not prime Then
  
  If add_en_detalle Then
      
    Call cbAgregar_Click
    
  Else
  
    If rc.RecordCount = 0 Then
        
            If MsgBox("No se encuentran Artículos. ¿Crear?", vbYesNo + vbQuestion, "Artículos") = vbNo Then
                Unload Me
            Else
            Call cbAgregar_Click
            End If
              
    Else
       
            Call cmdFirst_Click
            Call cbCancelar_Click

        
    End If
    
  End If
    
    
prime = True
End If

End Sub







Private Sub Form_Load()
  
  
   On Error GoTo Form_Load_Error

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm '(Screen.Height - Height) \ 2
  
  'ORIGINAL:
  'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
  
  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
   
  Set rc = New Recordset
  oSQL.AddTable "MAARTIC"
  oSQL.AddOrderClause "CODIGO"
  oSQL.AddSimpleWhereClause "MBAJA", "0"
  oSQL.AddSimpleWhereClause "HIST", "0"
  
  If add_en_detalle Then
    oSQL.AddSimpleWhereClause "TEMPOR", TemporadaWrk
  Else
    oSQL.AddSimpleWhereClause "TEMPOR", TemporadaActual
  End If

  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    
  'rc.Move 0

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With lblPRECOM
  Set .DataSource = rc
        .DataField = "PRECOM"
  End With
  
    
  If TipoServer = 1 Then
  
  With ioFOTO
  Set .DataSource = rc
        .DataField = "FOTO"
  End With
  
  End If
  
 
With ioMODELO
    Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "MODELO"
        .LongMaxima = 30
End With
   
With cbSECCION
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SECCIONES WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    '.DataField = "SECCION"
    .carga
    .CodigoWidth = 800
    'Set .DataSource = rc
End With
 
With cbFAMILIA
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .DataField = "FAMILIA"
    .carga
    DoEvents
    Set .DataSource = rc
End With

With cbSUBFAM
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .DataField = ""
    .carga
    DoEvents
   ' Set .DataSource = rc
End With

With cbTIPOIVA
    .ConexionString = locCnn
    .LenCodigo = 3
    
    If TipoServer = 1 Then
        .SQLString = "SELECT CODIGO, CAST(IVA AS char(" & .LenCodigo & ")) + ' - ' + DESCRIPCION FROM IVA WHERE MBAJA = 0 ORDER BY CODIGO"
    Else
        .SQLString = "SELECT CODIGO, CONCAT(CONCAT(CAST(IVA AS char),' - '), DESCRIPCION) FROM IVA WHERE MBAJA = 0 ORDER BY CODIGO"
    End If
    
    .DataField = "TIPOIVA"
    .carga
    DoEvents
    .CodigoWidth = 500
    Set .DataSource = rc
End With

With cbCODPROV
    .ConexionString = locCnn
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .DataField = "CODPROV"
    DoEvents
    .carga
     Set .DataSource = rc
End With

With ioABREVIA
    Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "ABREVIA"
        .LongMaxima = 20
End With

With ioREF
    Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "REF"
        .LongMaxima = 15
End With

With ioSTOCK
    Set .DataSource = rc
        .DataField = "STOCK"
End With

With ioTARIFA
    Set .DataSource = rc
        .DataField = "TARIFA"
End With

'With ioPRECOM
'.Alineacion = 1
'.dspFormat = "Currency"
'.SoloNumeros = True
  '  Set .DataSource = rc
   '     .DataField = "PRECOM"
       ' .displayformat = "00000.00 €"
       ' .Format = "####0.##"
'End With



With ioPRECOM
 'Set .DataSource = rc
.Alineacion = 1
'.dspFormat = "Currency"
.SoloNumeros = True
   'Set .DataSource = rc
        ' .DataField = "PRECOM"
   ' .displayformat = "00000.00 €"
   .dspFormat = "Currency"
   '.Mask = "#####.##"
    
       ' .Format = "####0.##"
       .PermitirBlanco = False
End With

With ioPREVEN
    .Alineacion = 1
.dspFormat = "Currency"
.SoloNumeros = True
  '  Set .DataSource = rc
   '     .DataField = "PREVEN"
        '.displayformat = "00000.00 €"
       ' .Format = "#####.##"
End With

With ioSTOCKMIN
.Alineacion = 1
.SoloNumeros = True
   '  Set .DataSource = rc
   ' .DataField = "STOCKMIN"
   ' .Format = "######.#"
End With

With ioSTOCKMAX
.Alineacion = 1
.SoloNumeros = True
  '   Set .DataSource = rc
  '  .DataField = "STOCKMAX"
  '  .displayformat = "000000.0"
  '  .Format = "######.#"
End With

With ioPEDIR
.Alineacion = 1
   ' Set .DataSource = rc
   '  .DataField = "PEDIR"
      '  .displayformat = "000000.0"
       ' .Format = "######.#"
End With

With ioDCTO
.Alineacion = 1
.dspFormat = "00.00"
.SoloNumeros = True
   '     .DataField = "DCTO2"
  '  Set .DataSource = rc
      '  .displayformat = "00.00 %"
       ' .Format = "##.##"
End With

'no enlazar a datos
With Temporada
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .carga
    .TabStop = False
End With



  
  '''''''QUITAR''''''''''''''''''
  With ioHIST
  Set .DataSource = rc
        .DataField = "HIST"
        .Locked = False
  End With
  '''''''QUITAR''''''''''''''''''
  
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
  
  With ioIVACOM
    Set .DataSource = rc
    .DataField = "IVACOM"
    .SoloNumeros = True
    .LongMaxima = 3
    .Alineacion = 1
  End With
              
  'obtener la descripción de la temporada actual
  If add_en_detalle Then
    lblTemActual.Caption = devuelve_campo("SELECT AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & TemporadaWrk)
    ioIVACOM.Visible = False
  Else
    lblTemActual.Caption = devuelve_campo("SELECT AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & TemporadaActual)
    ioIVACOM.Visible = True
  End If
  
  DoEvents
 
  If rc.RecordCount > 0 Then
  If rc.EditMode <> adEditNone Then
    rc.CancelUpdate
    End If
  End If
  
  
  Call llena_grid
  
  
  'si estamos desde pedidos, poner todos los botones como TabStop = false
  'para que funcione bien la inserción en serie.
  If add_en_detalle Then
  
  'cbActualizar.TabStop = False
  
  End If
  
              
  mbDataChanged = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario FrmMntArt"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Porcen_IVA As Byte
 
  If (mbEditFlag Or mbAddNewFlag) And KeyCode <> vbKeyF10 Then Exit Sub

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
      
            If Not add_en_detalle Then Call cbLista_click
      
      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
        Case vbKeyF8
        Call cmdLast_Click
        
       Case vbKeyF10
            
            KeyCode = 0
            
            If Screen.ActiveControl.Name = "ioPRECOM" Then
            'para PRECOM
            
            If Not Iva_Descontado Then
            
                
            
                'descontar el IVA del PRECIO
                If ioPRECOM.Text <> "" And cbTIPOIVA.Text <> "" Then
            
                    Porcen_IVA = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & cbTIPOIVA.Text)
                    ioPRECOM.Text = Format(ioPRECOM.Text - ((ioPRECOM.Text * Porcen_IVA / 100)), "Currency")
                    lblstatus.Caption = "Se ha descontado el " & Porcen_IVA & " % de IVA a el precio de compra"
                    Iva_Descontado = True
                    
                    Exit Sub
                                   
                Else
            
                    lblstatus.Caption = "Debe establecer un precio y un IVA para realizar la operación"
                    Exit Sub
                        
               End If
            
            End If
            
            
            End If
      
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

'If rc.EditMode <> adEditNone Then rc.CancelUpdate

'If rc.State <> 0 Then rc.Close
'Set rc = Nothing

Set oSQL = Nothing

  ' With locCnn
  '  If .State <> 0 Then .Close
  ' End With

'If Me.MDIChild = True Then frmMenuTactil.Show

Set FrmMntArt = Nothing



End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub des_enlaza_campos()
  
  With ioCODIGO
  Set .DataSource = Nothing
      .DataField = ""
  End With
  
  With lblPRECOM
Set .DataSource = Nothing
      .DataField = ""
  End With
  
    With ioIVACOM
  Set .DataSource = Nothing
      .DataField = ""
  End With
    
  With ioFOTO
Set .DataSource = Nothing
      .DataField = ""
  End With
   
With ioMODELO
Set .DataSource = Nothing
      .DataField = ""
End With
   
With cbSECCION
Set .DataSource = Nothing
      .DataField = ""
End With
 
With cbFAMILIA
Set .DataSource = Nothing
      .DataField = ""
End With

With cbSUBFAM
Set .DataSource = Nothing
      .DataField = ""
End With

With cbTIPOIVA
Set .DataSource = Nothing
      .DataField = ""
End With

With cbCODPROV
Set .DataSource = Nothing
      .DataField = ""
End With

With ioABREVIA
Set .DataSource = Nothing
      .DataField = ""
End With

With ioREF
Set .DataSource = Nothing
      .DataField = ""
End With

With ioSTOCK
Set .DataSource = Nothing
      .DataField = ""
End With

With ioTARIFA
Set .DataSource = Nothing
      .DataField = ""
End With

  
  '''''''QUITAR''''''''''''''''''
With ioHIST
Set .DataSource = Nothing
      .DataField = ""
End With
  '''''''QUITAR''''''''''''''''''
  
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


Private Sub enlaza_campos()
   
  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With lblPRECOM
  Set .DataSource = rc
      .DataField = "PRECOM"
  End With
  
  With ioIVACOM
  Set .DataSource = rc
      .DataField = "IVACOM"
  End With
  
    
  With ioFOTO
  Set .DataSource = rc
        .DataField = "FOTO"
  End With
   
With ioMODELO
    Set .DataSource = rc
        .DataField = "MODELO"
End With
   
With cbSECCION
        .DataField = "SECCION"
    Set .DataSource = rc
End With
 
With cbFAMILIA
        .DataField = "FAMILIA"
    Set .DataSource = rc
End With

With cbSUBFAM
        .DataField = "SUBFAM"
    Set .DataSource = rc
End With

With cbTIPOIVA
        .DataField = "TIPOIVA"
    Set .DataSource = rc
End With

With cbCODPROV
    .DataField = "CODPROV"
     Set .DataSource = rc
End With

With ioABREVIA
    Set .DataSource = rc
        .DataField = "ABREVIA"
End With

With ioREF
    Set .DataSource = rc
        .DataField = "REF"
End With

With ioSTOCK
    Set .DataSource = rc
        .DataField = "STOCK"
End With

With ioTARIFA
    Set .DataSource = rc
        .DataField = "TARIFA"
End With

  
  '''''''QUITAR''''''''''''''''''
With ioHIST
Set .DataSource = rc
        .DataField = "HIST"
End With
  '''''''QUITAR''''''''''''''''''
  
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

Private Sub cbLista_click()
Dim tmpbook As Variant

If Not add_en_detalle Then

    With frmFlexArt
    
    Set .miosql = oSQL
    
    
    With .fg
            .ColFormat(1) = "00000"
             Set frmFlexArt.miRc = rc
    End With
    
    End With
    
    

    Call des_enlaza_campos
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MOSTRAR
    frmFlexArt.Show 1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call enlaza_campos
    
    tmpbook = rc.Bookmark
     
    rc.Bookmark = tmpbook
    Set tmpbook = Nothing
    
End If

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub ioDCTO_GotFocus()

If Tab1.Tab <> 1 Then Tab1.Tab = 1

End Sub

Private Sub ioDCTO_Validate(Cancel As Boolean)
   On Error GoTo ioDCTO_Validate_Error

With ioDCTO
If Trim(.Text) <> "" Then
    If CDbl(Replace(.Text, ".", ",")) >= 100 Then
    lblstatus.Caption = "No se permite un Descuento mayor o igual al 100%"
    .Text = 0
    .CancelarValidacion
    Cancel = True
    Else
        lblstatus.Caption = ""
    End If
End If
End With

Call calcula_descuentos
DoEvents
Call llena_grid

DoEvents
Tab1.Tab = 2

   On Error GoTo 0
   Exit Sub

ioDCTO_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDCTO_Validate of Formulario FrmMntArt"
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

Private Sub ioMODELO_Validate(Cancel As Boolean)
If ioMODELO.Text <> "" Then ioMODELO.Text = UCase(ioMODELO.Text)
End Sub

Private Sub ioPRECOM_GotFocus()

lblstatus.Caption = "Pulse F10 para descontar el IVA (para precios IVA incluido)"

End Sub

Private Sub ioPRECOM_LostFocus()

'si la tarifa esta a true, pasar desde aqui a la siguiente pagina
If ioTARIFA.Value = True Then
     Tab1.Tab = 1
End If

End Sub

Private Sub ioPRECOM_Validate(Cancel As Boolean)

'validación de precio de compra
If ioPRECOM.Text <> "" Then
    
    If IsNumeric(ioPRECOM.Text) Then
        
        'si es mayor de 1000€ avisar al usuario (anormalmente alto)
        'si es menor de 1 €, (anormalmente bajo)
        If (CDbl(ioPRECOM.Text) >= 1000) Or (CDbl(ioPRECOM.Text) <= 1) Then
            'preguntar
            If MsgBox("El precio de compra introducido es de: " & ioPRECOM.Text & " . ¿Esta usted seguro?", vbQuestion + vbYesNo) = vbNo Then
                Cancel = True
                ioPRECOM.SetFocus
                DoEvents
                Exit Sub
            End If
        
        End If
        
    End If

End If


DoEvents
Call llena_grid

End Sub

Private Sub ioPREVEN_GotFocus()

If Tab1.Tab > 0 Then Tab1.Tab = 0

End Sub






Private Sub ioPREVEN_Validate(Cancel As Boolean)

If ioPREVEN.Text <> "" Then
    
    If IsNumeric(ioPREVEN.Text) Then
        
        'si es mayor de 1000€ avisar al usuario (anormalmente alto)
        'si es menor de 1 €, (anormalmente bajo)
        If (CDbl(ioPREVEN.Text) >= 1000) Or (CDbl(ioPREVEN.Text) <= 1) Then
            'preguntar
            If MsgBox("El precio de compra introducido es de: " & ioPREVEN.Text & " . ¿Esta usted seguro?", vbQuestion + vbYesNo) = vbNo Then
                Cancel = True
                ioPREVEN.SetFocus
                Exit Sub
            End If
        
        End If
        
    End If

End If
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : ioREF_Validate
' Fecha/Hora  : 18/01/2004 20:25
' Autor       : JCASTILLO
' Propósito   : Comprobar que no exista la referencia antes de meter el artículo
'---------------------------------------------------------------------------------------
Private Sub ioREF_Validate(Cancel As Boolean)
Dim tmpart As Variant

   On Error GoTo ioREF_Validate_Error

If Not mbAddNewFlag Then Exit Sub

If ioREF.Text = "" Then Exit Sub

lblstatus.Caption = "Comprobando referencia (espere por favor) ..."

'comprobar si existe esa referencia ...
'buscar solo en la temporada de trabajo
If add_en_detalle Then
    tmpart = devuelve_matriz("SELECT CODIGO, TEMPOR FROM MAARTIC WHERE REF = '" & ioREF.Text & "' AND MBAJA = 0 AND TEMPOR = " & TemporadaWrk & " AND CODPROV = " & cbCODPROV.Text, locCnn)
Else
    tmpart = devuelve_matriz("SELECT CODIGO, TEMPOR FROM MAARTIC WHERE REF = '" & ioREF.Text & "' AND MBAJA = 0 AND TEMPOR = " & TemporadaActual, locCnn)
End If


If IsArray(tmpart) Then

    lblstatus.Caption = "La referencia ya existe"
    
    'si hemos entrado por el pedido ...
    If add_en_detalle Then
        
        'si no quiere insertar, salir
        If MsgBox("La referencia YA EXISTE, ¿Desea insertar en pedido?." & Chr(13) & _
               "Codigo: " & tmpart(0) & " Temporada: " & tmpart(1), vbQuestion + vbYesNo, titulo) = vbNo Then
           'cancelar cambios y salir
           Call cbCancelar_Click
           Exit Sub
        End If
    
    'si hemos entrado por añadir normal
    Else
    
        'si existe, pero no estamos en pedido, dar la posibilidad de ir al
        'artículo de esa referencia
        If MsgBox("La referencia YA EXISTE, ¿Desea ir al artículo?" & Chr(13) & _
               "Codigo: " & tmpart(0) & " Temporada: " & tmpart(1), vbQuestion + vbYesNo, titulo) = vbYes Then
                       
            'cancelar cambios
            Call cbCancelar_Click
            
            'mover campos de rc a pantalla
            mbAddNewFlag = False
            
            'ir al artículo ...
            Call des_enlaza_campos
    
            oSQL.ClearWhereClause
            oSQL.AddSimpleWhereClause "CODIGO", tmpart(0)
            oSQL.AddSimpleWhereClause "TEMPOR", tmpart(1)
    
            If rc.State = 1 Then
                rc.CancelUpdate
                rc.Close
            End If
        
            rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    
            Call enlaza_campos
            DoEvents
            
            Call mov_rc_pan
            
            DoEvents
            
            
            
            Exit Sub
        
        'no quiere ir al articulo
        Else
        
            'cancelar cambios
            Call cbCancelar_Click
            Exit Sub
        
        End If
          

    
    End If
    
        'ir al artículo ...
        Call des_enlaza_campos
    
        oSQL.ClearWhereClause
        oSQL.AddSimpleWhereClause "CODIGO", tmpart(0)
        oSQL.AddSimpleWhereClause "TEMPOR", tmpart(1)
    
        If rc.State = 1 Then
            rc.CancelUpdate
            rc.Close
        End If
        
        rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    
        Call enlaza_campos
        DoEvents
        
        'si hemos entrado por el pedido, añadir el artículo al
        'pedido actual
        If add_en_detalle Then
                
               add_ref = True
            
               'mover campos de rc a pantalla
               Call mov_rc_pan
               
               DoEvents
               
               'para q añada el artículo directamente
               Call añade_en_detalle
               
               add_ref = False
               
               Call cbCancelar_Click
               
              'abrir el formulario detpedpro, y posicionarse para actualizar
              'el registro introducido
               Call Posicionarse_En_DetPedPro
               
               
 
        End If

End If

lblstatus.Caption = "Referencia OK"

Set tmpart = Nothing


   On Error GoTo 0
   Exit Sub

ioREF_Validate_Error:

    add_ref = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioREF_Validate de Formulario FrmMntArt"

End Sub

'Private Sub ioPREVEN_Validate(Cancel As Boolean)
'DoEvents
'Tab1.Tab = 1
'End Sub

Private Sub ioSTOCKMIN_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub

Private Sub ioTARIFA_Click()

If ioTARIFA.Locked Then Exit Sub
'If ioTARIFA.Enabled Then Exit Sub

'validar tarifa, bloqueando precio de venta si elegimos por tarifa.
If ioTARIFA.Value = True Then

    With ioPREVEN
        .Text = 0
        .Locked = True
        .TabStop = False
    End With

    'pasar a la siguiente pagina
   ' Tab1.Tab = 1

Else

    With ioPREVEN
        .Locked = False
        .TabStop = True
    End With

End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : calcula_descuentos
' DateTime  : 05/11/2003 21:01
' Author    : Administrador
' Purpose   : Calcular el precio de compra final una vez efectuados los
'             descuentos, y mostrar por pantalla
'---------------------------------------------------------------------------------------
'
Private Sub calcula_descuentos()
Dim tmpprecom As Double


'si estan los descuentos vacios poner el precio directamente
   On Error GoTo calcula_descuentos_Error
 On Error Resume Next

If rc.BOF Then Exit Sub

If (ioDCTO.Text = "" Or ioDCTO.Text = "0") Then
    lblPrecomFin.Caption = lblPRECOM.Caption

Else
    
    If IsNull(rc.fields("PRECOM").Value) Then Exit Sub
    'obtener precio de compra
    tmpprecom = rc.fields("PRECOM").Value
   
        'sino aplicar el segundo descuento
    If ioDCTO.Text <> "" And ioDCTO.Text <> "0" Then
    
        tmpprecom = tmpprecom - ((tmpprecom * CSng(ioDCTO.Text)) / 100)
        lblPrecomFin.Caption = Format(tmpprecom, "currency")
    
    End If

End If

   On Error GoTo 0
   Exit Sub

calcula_descuentos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcula_descuentos of Formulario FrmMntArt"

End Sub



Private Sub lblprecom_Change()

lblPRECOM.Caption = Format(lblPRECOM.Caption, "currency")
Call calcula_descuentos

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  On Error GoTo rc_MoveComplete_Error

  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
 
  If Not IsNull(rc.fields("MODELO")) Then
    lblArticulo.Caption = rc.fields("MODELO")
  Else
    lblArticulo.Caption = ""
  End If
  
  If Not IsNull(rc.fields("COMEN")) Then
    lblExisteCom.Visible = True
    lblExisteCom2.Visible = True
  Else
    lblExisteCom.Visible = False
    lblExisteCom2.Visible = False
  End If
  
  Call mov_rc_pan
  
  End If
  
  Call llena_grid

   On Error GoTo 0
   Exit Sub

rc_MoveComplete_Error:

    If Err.Number = 3021 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rc_MoveComplete of Formulario FrmMntArt"
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
    
  On Error GoTo AddErr
  
  If add_en_detalle Then
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from MAARTIC where TEMPOR =" & TemporadaWrk)
  Else
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from MAARTIC where TEMPOR =" & TemporadaActual)
  End If

    'locCnn.BeginTrans

  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    'DoEvents
    
    mbAddNewFlag = True
    
    .AddNew
    
  ioPRECOM.Text = ""
  ioPREVEN.Text = ""
  ioDCTO.Text = ""
  ioSTOCKMIN.Text = ""
  ioSTOCKMAX.Text = ""
  ioPEDIR.Text = ""
  cbSUBFAM.Text = ""
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    If add_en_detalle Then
        .fields("TEMPOR") = TemporadaWrk
    Else
        .fields("TEMPOR") = TemporadaActual
    End If
    
    .fields("CODIGO") = tmpcodigo
    
    
    'ioTARIFA.Value = True
    cbTIPOIVA.Text = 4  '0 %
    ioSTOCK.Value = True

    tmpcodigo = ""
    
    lblstatus.Caption = "Agregar registro"
    
    SetButtons False
    cbActualizar.Visible = True
    
    
    If add_en_detalle Then
        'que entre directamente agregando registro
        cbCODPROV.Text = ProveedPedido
        cbCODPROV.Locked = True
    End If
    
    
  Tab1.SetFocus
  ioREF.SetFocus
  End With

    
  Exit Sub
AddErr:
mbAddNewFlag = False
  If Err.Number = -2147217842 Then Exit Sub
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
    
  ioREF.SetFocus
  
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
Dim Color As Long
Dim forecol As Long
  On Error GoTo UpdateErr

'If Not mbEditFlag Or mbAddNewFlag Then Exit Sub
    
    If ioPRECOM.Text = "" Then
        rc.fields("PRECOM") = 0
        ioPRECOM.Text = 0
    End If
    
    If ioPREVEN.Text = "" Then
        rc.fields("PREVEN") = 0
        ioPREVEN.Text = 0
    End If
    
    If ioDCTO.Text = "" Then
        rc.fields("DCTO") = 0
        ioDCTO.Text = 0
    End If
    
    If ioIVACOM.Text = "" Then
        rc.fields("IVACOM") = 0
        ioIVACOM.Text = 0
    End If
    
    If ioSTOCKMIN.Text = "" Then
        rc.fields("STOCKMIN") = 0
        ioSTOCKMIN.Text = 0
    End If
    
    If ioSTOCKMAX.Text = "" Then
        rc.fields("STOCKMAX") = 0
        ioSTOCKMAX.Text = 0
    End If
    
    If ioPEDIR.Text = "" Then
        rc.fields("PEDIR") = 0
        ioPEDIR.Text = 0
    End If
        
    'validaciones
  '  If cbSECCION.Text = "" Then
   '     lblStatus.Caption = "SECCION no puede estar en blanco"
  '      cbSECCION.SetFocus
  '      Exit Sub
  '  End If
    
    If cbFAMILIA.Text = "" Then
        lblstatus.Caption = "FAMILIA no puede estar en blanco"
        cbFAMILIA.SetFocus
        Exit Sub
    End If
    
    If ioMODELO.Text = "" Then
        lblstatus.Caption = "MODELO no puede estar en blanco"
        ioMODELO.SetFocus
        Exit Sub
    End If
    
    If cbSUBFAM.Text = "" Then
        cbSUBFAM.Text = 0
    End If
    
    If cbCODPROV.Text = "" Then
        lblstatus.Caption = "PROVEEDOR no puede estar en blanco"
        cbCODPROV.SetFocus
        Exit Sub
    End If
    
    If cbTIPOIVA.Text = "" Then
        lblstatus.Caption = "TIPO de IVA no puede estar en blanco"
        cbTIPOIVA.SetFocus
        Exit Sub
    End If
    
    If CDbl(ioPRECOM.Text) = "0" Then
        lblstatus.Caption = "PRECIO DE COMPRA no puede ser cero"
        ioPRECOM.SetFocus
        Exit Sub
    End If
        
    'tiene q establecer alguno de los 2, a tarifa o precio fijo
    If ioTARIFA.Value = False And CDbl(ioPREVEN.Text) = "0" Then
       lblstatus.Caption = "Debe elegir Precio de venta A TARIFA o FIJO"
       Exit Sub
    End If

    'tiene q establecer alguno de los 2, a tarifa o precio fijo
    If ioTARIFA.Value = True And CDbl(ioPREVEN.Text) <> "0" Then
       lblstatus.Caption = "Debe elegir Precio de venta A TARIFA o FIJO (no ambos)"
       Exit Sub
    End If
  
 ' rc.Fields("SUBFAM") = cbSUBFAM.Text
  
  'MsgBox cbSUBFAM.Text
  
    Call mov_pan_rc
  'If rc.EditMode = adEditNone Then Exit Sub
  
  
  rc.Update  'Batch adAffectAll
  
  'rc.Update 'Batch adAffectAll
  
  
  
  If mbAddNewFlag Then
    'añadir el registro en el adoDetalleRs si viene del pedido
    If add_en_detalle Then Call añade_en_detalle
    rc.MoveLast              'va al nuevo registro
  End If
  
  If mbEditFlag Then
    Color = &HCCFF66
    forecol = vbBlack
  End If
  
  If mbAddNewFlag Then
    Color = vbBlue
    forecol = vbWhite
  End If
  
  tmplinea = tmplinea + 1
  
  With vsModificaciones
  
    .TextMatrix(tmplinea, 1) = ioCODIGO.Caption  'codigo
    
    .Col = 1
    .CellBackColor = Color  'establecer el color de fondo de la celda
    .Cell(flexcpForeColor) = forecol 'establecer el fore color
    
    .TextMatrix(tmplinea, 2) = ioMODELO.Text  'modelo
     
     .Col = 2
     .CellBackColor = Color
     .Cell(flexcpForeColor) = forecol
     
    .TextMatrix(tmplinea, 3) = ioPRECOM.Text  'precio compra
    
    .Col = 3
    .CellBackColor = Color
    .Cell(flexcpForeColor) = forecol
    
    .TextMatrix(tmplinea, 4) = Now 'ultima modificacion
    
    .Col = 4
    .CellBackColor = Color
    .Cell(flexcpForeColor) = forecol
    
    
    
    .AutoSize 1, .Cols - 1
    DoEvents
  End With
  
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  
  lblstatus.Caption = ""
  Tab1.Tab = 0
  Iva_Descontado = False
  
  'abrir el formulario detpedpro, y posicionarse para actualizar
  'el registro introducido
  Call Posicionarse_En_DetPedPro
   
  
  Exit Sub
UpdateErr:
'error de no transaccion activa
 ' If (Err.Number = -2147217887) Then Exit Sub
  Debug.Print Err.Number
  MsgBox Err.Description, vbInformation, "Atención"
End Sub


Private Sub Posicionarse_En_DetPedPro()
 
   On Error GoTo Posicionarse_En_DetPedPro_Error

 'posicionarse y editar la ultima linea insertada
   If add_en_detalle Then
    
    With frmPedProv
    
            If Not .adoPrimaryRS.RecordCount = 0 Then
                If Not .adoDetalleRS.BOF And .Linea_Creada > 0 Then
        
            'abrir para editar la ultima linea añadida
                .adoDetalleRS.MoveFirst
                .adoDetalleRS.Find "LINEA = " & .Linea_Creada
               
                   ' llamar a frmDPedPro
                    frmDPedPro.NumeroPedido = frmPedProv.ioNUMERO.Caption
                    If Not .adoPrimaryRS.EOF Then
                        Set frmDPedPro.rc = .adoDetalleRS
                    End If
                    
                    frmDPedPro.Show
                
                .Linea_Creada = 0
                'Unload Me
                
            End If
        End If
        
    End With
   
   End If

   On Error GoTo 0
   Exit Sub

Posicionarse_En_DetPedPro_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Posicionarse_En_DetPedPro de Formulario FrmMntArt"
  
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
  
  ioTARIFA.Enabled = Not bVal
  cbAgregar.Visible = bVal
  cbEdicion.Visible = bVal
 
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
  
  Temporada.Enabled = bVal
  cmCambTem.Enabled = bVal
  
  
  
  'si venimos del pedido, no mostrar el boton de lista
  If add_en_detalle Then
    cbLista.Visible = False
  Else
    cbLista.Visible = bVal
  End If
  
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  cbSECCION.Locked = bVal
  cbFAMILIA.Locked = bVal
  cbSUBFAM.Locked = bVal
  cbTIPOIVA.Locked = bVal
  
  'si venimos del pedido, el codigo de proveedor sera siempre
  'el mismo (el del pedido) y el combo debe aparecer siempre
  'como locked
  If Not add_en_detalle Then cbCODPROV.Locked = bVal
  
  cmComentario.Enabled = Not bVal
  
 
  
End Sub



'---------------------------------------------------------------------------------------
' Procedure : llena_grid
' DateTime : 05/11/2003 17:14
' Author     : Administrador
' Purpose   : Llenar el grid con las tarifas y el precio de venta segun tarifa
'---------------------------------------------------------------------------------------
Private Sub llena_grid()
Dim tmpfila As Long
Dim tmprecom As Single
Dim tmptipoiva As Single


  On Error GoTo llena_grid_Error

  Dim tmprc As New ADODB.Recordset
  
  'si tiene descuento, coger el precio de lblPrecomFin
  If lblPRECOM.Caption <> lblPrecomFin.Caption And lblPrecomFin.Caption <> "" Then
   tmprecom = CSng(lblPrecomFin.Caption)
  Else
  
  If IsNumeric(ioPRECOM.Text) Then If ioPRECOM.Text <> "" Then tmprecom = ioPRECOM.Text
  End If
  
  If cbTIPOIVA.Text <> "" Then tmptipoiva = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & CLng(cbTIPOIVA.Text))

  
  tmprc.Open "select  CODIGO,  DESCRIPCION, PORCEN from tarifas order by activa desc, codigo", locCnn, adOpenDynamic, adLockReadOnly
 
 tmpfila = 1
  With vsTarifas
  
  .Clear
  
  .TextMatrix(0, 1) = "Tarifa"
  .TextMatrix(0, 2) = "Nombre"
  .TextMatrix(0, 3) = "Precio Comp"
  .TextMatrix(0, 4) = "Margen"
  .TextMatrix(0, 5) = "Precio Venta"
  .TextMatrix(0, 6) = "Precio + IVA"
  
  Do Until tmprc.EOF
  
    .TextMatrix(tmpfila, 1) = tmprc.fields(0)   'codigo
    .TextMatrix(tmpfila, 2) = tmprc.fields(1)   'descripción
    .TextMatrix(tmpfila, 3) = tmprecom          'precio de compra
    .TextMatrix(tmpfila, 4) = Format(tmprc.fields(2).Value, "000.00") & " %"  'porcen
    
    'precio de venta y precio con iva
    If (tmprecom > 0) And (tmprc.fields(2).Value > 0) Then
        .TextMatrix(tmpfila, 5) = tmprecom + (tmprecom * tmprc.fields(2) / 100)
    
        'iva
        If (tmptipoiva > 0) Then .TextMatrix(tmpfila, 6) = .TextMatrix(tmpfila, 5) + (.TextMatrix(tmpfila, 5) * tmptipoiva / 100)
    
    End If
   
    tmprc.MoveNext
    tmpfila = tmpfila + 1
    
  Loop
    
    .AutoSize 1, .Cols - 1
  End With
    
  tmprc.Close
  Set tmprc = Nothing
  
   On Error GoTo 0
   Exit Sub

llena_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure llena_grid of Formulario FrmMntArt"

End Sub


Private Sub mov_pan_rc()

With rc

    If cbSECCION.Text <> "" Then
    .fields("SECCION") = cbSECCION.Text
    Else
    .fields("SECCION") = Null
    End If
    
    .fields("PRECOM") = ioPRECOM.Text
    .fields("PREVEN") = ioPREVEN.Text
    .fields("STOCKMAX") = ioSTOCKMAX.Text
    .fields("STOCKMIN") = ioSTOCKMIN.Text
    .fields("PEDIR") = ioPEDIR.Text
    .fields("DCTO") = ioDCTO.Text
    
    If cbSUBFAM.Text <> "" Then
        .fields("SUBFAM") = cbSUBFAM.Text
    Else
        .fields("SUBFAM") = Null
    End If
    
End With

End Sub

Private Sub mov_rc_pan()

If (Not mbAddNewFlag) Or (add_ref) Then

With rc

    If Not IsNull(.fields("SECCION")) Then
        cbSECCION.Text = .fields("SECCION")
    Else
        cbSECCION.Text = ""
    End If

    ioPRECOM.Text = .fields("PRECOM")
    ioPREVEN.Text = .fields("PREVEN")
    ioSTOCKMAX.Text = .fields("STOCKMAX")
    ioSTOCKMIN.Text = .fields("STOCKMIN")
    ioPEDIR.Text = .fields("PEDIR")
    ioDCTO.Text = .fields("DCTO")
    
    If Not IsNull(.fields("SUBFAM")) Then
        cbSUBFAM.Text = .fields("SUBFAM")
    Else
        cbSUBFAM.Text = ""
    End If
    
End With

End If

End Sub




'---------------------------------------------------------------------------------------
' Subrutina   : vsTarifas_Click
' Fecha/Hora  : 30/11/2003 21:10
' Autor       : JCASTILLO
' Propósito   : Llama al formulario de Ajuste de precio de venta (para ayudar en el
'               ajuste de precio psicologico. Usa el % margen de la tarifa seleccionada
'---------------------------------------------------------------------------------------
Private Sub vsTarifas_dblClick()

   On Error GoTo vsTarifas_Click_Error

           'si no esta editando o añadiendo, salir
           If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub
                      
           'si esta a blanco salir
           If vsTarifas.TextMatrix(vsTarifas.Row, 4) = "" Then Exit Sub
                    
           
           'para PREVEN
                        
            With frmAjusPreVen
                 
                .AJ_Cancelado = False
                .AJ_Margen_Tarifa = CDbl(Left(vsTarifas.TextMatrix(vsTarifas.Row, 4), Len(vsTarifas.TextMatrix(vsTarifas.Row, 4)) - 1))
                .AJ_Precio_Compra = ioPRECOM.Valor
                .AJ_TipoIVA = cbTIPOIVA.Text
                'mostrar como modal
                .Show 1
                DoEvents
                
                'si no ha salido por cancelar
                If Not .AJ_Cancelado Then
                
                    DoEvents
                    'quitar a tarifa
                    ioTARIFA.Value = False
                    rc.fields("TARIFA") = False
                    'fijar precio de venta
                    ioPREVEN.Text = .AJ_PreVenta_Final
                    rc.fields("PREVEN") = .AJ_PreVenta_Final
                        
                    DoEvents
                                
                End If
                
                Tab1.Tab = 2
                ioSTOCKMIN.SetFocus
                
            End With
            

   On Error GoTo 0
   Exit Sub

vsTarifas_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento vsTarifas_Click de Formulario FrmMntArt"

End Sub
