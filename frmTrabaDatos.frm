VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTrabaDatos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajar con Datos ..."
   ClientHeight    =   4170
   ClientLeft      =   4260
   ClientTop       =   3150
   ClientWidth     =   9600
   ControlBox      =   0   'False
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
   ScaleHeight     =   4170
   ScaleWidth      =   9600
   Begin TabDlg.SSTab Tab1 
      Height          =   2610
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   4604
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Pedidos"
      TabPicture(0)   =   "frmTrabaDatos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ioNPFIN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ioNPINI"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbTrabajarPedido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmPasarPendientePedido"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbDeshacerPedido"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioNUMPED"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Ventas"
      TabPicture(1)   =   "frmTrabaDatos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ioTICKET"
      Tab(1).Control(1)=   "cmSeleccionaVenta"
      Tab(1).Control(2)=   "cmBorrarVenta"
      Tab(1).Control(3)=   "cmPasarPendienteVenta"
      Tab(1).Control(4)=   "ioFecha"
      Tab(1).Control(5)=   "cbCAJAS"
      Tab(1).Control(6)=   "cbCambiarCaja"
      Tab(1).Control(7)=   "Shape2"
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(10)=   "Label16"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Transferencias"
      TabPicture(2)   =   "frmTrabaDatos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "miText2"
      Tab(2).Control(1)=   "cmSeleccionarTrans"
      Tab(2).Control(2)=   "cmBorrarTransferencia"
      Tab(2).Control(3)=   "cmPasarPendienteTrans"
      Tab(2).Control(4)=   "Label2"
      Tab(2).ControlCount=   5
      Begin PCGestion.miText ioTICKET 
         Height          =   495
         Left            =   -74040
         TabIndex        =   3
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
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
      Begin PCGestion.chameleonButton cmSeleccionaVenta 
         Height          =   870
         Left            =   -72435
         TabIndex        =   11
         Top             =   375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1535
         BTYPE           =   3
         TX              =   "&Seleccionar Venta"
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
         MICON           =   "frmTrabaDatos.frx":0054
         PICN            =   "frmTrabaDatos.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cmBorrarVenta 
         Height          =   525
         Left            =   -71970
         TabIndex        =   12
         Top             =   1980
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "&Borrar"
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
         MICON           =   "frmTrabaDatos.frx":0D4A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cmPasarPendienteVenta 
         Height          =   525
         Left            =   -70005
         TabIndex        =   13
         Top             =   1980
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "&Pasar a Pendiente"
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
         MICON           =   "frmTrabaDatos.frx":0D66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miText ioNUMPED 
         Height          =   495
         Left            =   1470
         TabIndex        =   0
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
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
      Begin PCGestion.chameleonButton cbDeshacerPedido 
         Height          =   675
         Left            =   3795
         TabIndex        =   15
         Top             =   1065
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Deshacer Pedido Completo"
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
         MICON           =   "frmTrabaDatos.frx":0D82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cmPasarPendientePedido 
         Height          =   675
         Left            =   1800
         TabIndex        =   16
         Top             =   1065
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Pasar a Pendiente"
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
         MICON           =   "frmTrabaDatos.frx":0D9E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miText miText2 
         Height          =   495
         Left            =   -72930
         TabIndex        =   6
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
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
      Begin PCGestion.chameleonButton cmSeleccionarTrans 
         Height          =   1170
         Left            =   -71535
         TabIndex        =   17
         Top             =   510
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   2064
         BTYPE           =   3
         TX              =   "&Seleccionar Transferencia"
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
         MICON           =   "frmTrabaDatos.frx":0DBA
         PICN            =   "frmTrabaDatos.frx":0DD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cmBorrarTransferencia 
         Height          =   525
         Left            =   -73605
         TabIndex        =   18
         Top             =   1875
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "&Borrar"
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
         MICON           =   "frmTrabaDatos.frx":1AB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cmPasarPendienteTrans 
         Height          =   525
         Left            =   -71640
         TabIndex        =   19
         Top             =   1875
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "&Pasar a Pendiente"
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
         MICON           =   "frmTrabaDatos.frx":1ACC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton cbTrabajarPedido 
         Height          =   675
         Left            =   1800
         TabIndex        =   21
         Top             =   1755
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "&Trabajar con Pedido"
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
         MICON           =   "frmTrabaDatos.frx":1AE8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miText ioNPINI 
         Height          =   495
         Left            =   3915
         TabIndex        =   1
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
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
      Begin PCGestion.miText ioNPFIN 
         Height          =   495
         Left            =   5865
         TabIndex        =   2
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
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
      Begin PCGestion.miText ioFecha 
         Height          =   495
         Left            =   -72060
         TabIndex        =   4
         Top             =   1365
         Width           =   1395
         _ExtentX        =   2461
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
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   525
         Left            =   -69750
         TabIndex        =   5
         Top             =   1365
         Width           =   4185
         _ExtentX        =   7303
         _ExtentY        =   926
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
      Begin PCGestion.chameleonButton cbCambiarCaja 
         Height          =   870
         Left            =   -68460
         TabIndex        =   26
         Top             =   450
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1535
         BTYPE           =   3
         TX              =   "&Cambiar de Caja"
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
         MICON           =   "frmTrabaDatos.frx":1B04
         PICN            =   "frmTrabaDatos.frx":1B20
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00FF0000&
         Height          =   1545
         Left            =   -70635
         Top             =   375
         Width           =   5160
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA ORIGEN"
         Height          =   600
         Left            =   -70620
         TabIndex        =   25
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
         Height          =   330
         Left            =   -72960
         TabIndex        =   24
         Top             =   1440
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00FF0000&
         Height          =   570
         Left            =   3090
         Top             =   420
         Width           =   4185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Final"
         Height          =   300
         Left            =   5085
         TabIndex        =   23
         Top             =   555
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial"
         Height          =   300
         Left            =   3105
         TabIndex        =   22
         Top             =   555
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   330
         Left            =   -73815
         TabIndex        =   20
         Top             =   555
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Pedido"
         Height          =   585
         Left            =   195
         TabIndex        =   14
         Top             =   375
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET"
         Height          =   330
         Left            =   -74970
         TabIndex        =   10
         Top             =   540
         Width           =   855
      End
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   780
      Left            =   3848
      TabIndex        =   7
      Top             =   3345
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1376
      BTYPE           =   9
      TX              =   "&Aceptar"
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
      MICON           =   "frmTrabaDatos.frx":27FA
      PICN            =   "frmTrabaDatos.frx":2816
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
      Height          =   780
      Left            =   4823
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3345
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1376
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
      MICON           =   "frmTrabaDatos.frx":34F0
      PICN            =   "frmTrabaDatos.frx":350C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   690
      Left            =   15
      Top             =   2625
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   1217
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
End
Attribute VB_Name = "frmTrabaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmTrabaDatos
' Fecha/Hora : 23/02/2004 13:00
' Autor      : JCastillo
' Propósito  : Trabajar con datos de PEDIDOS, VENTAS y TRANSFERENCIAs, para
'              borrar datos de la base de datos y cambiar su estado
'---------------------------------------------------------------------------------------

Option Explicit


Private Function comprueba_ticket() As Boolean
Dim tmps As String

If Trim(ioTICKET.Text) = "" Then
    MsgBox "Debe introducir un código de Ticket, o seleccionar una venta", vbInformation, titulo
    comprueba_ticket = True
    Exit Function
End If

If Trim(Len(ioTICKET.Text)) <= 3 Then
    MsgBox "Código de Ticket Inválido", vbExclamation, titulo
    comprueba_ticket = True
    Exit Function
End If

tmps = Right(Trim(ioTICKET.Text), 3)

If Not IsNumeric(tmps) Then
    MsgBox "Código de Ticket Inválido", vbExclamation, titulo
    comprueba_ticket = True
    Exit Function
End If

If CLng(tmps) <> CajaActual Then
    MsgBox "La venta no corresponde a la caja actual, imposible deshacer.", vbExclamation, titulo
    comprueba_ticket = True
    Exit Function
End If

End Function



Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub cbDeshacerPedido_Click()
Dim rcped As ADODB.Recordset
Dim entrans As Boolean
Dim tmpest As Variant
Dim pinicial As Long
Dim pfinal As Long
Dim var As Long
Dim rcab As ADODB.Recordset
Dim tmpvar As Variant

  On Error GoTo cbDeshacerPedido_Click_Error

'filtros
If (Trim(ioNUMPED.Text = "") And Trim(ioNPINI.Text = "") And Trim(ioNPFIN.Text = "")) Then Exit Sub
If Trim(ioNUMPED.Text <> "") And Not IsNumeric(ioNUMPED.Text) Then Exit Sub
If Trim(ioNPINI.Text <> "") And Not IsNumeric(ioNPINI.Text) Then Exit Sub
If Trim(ioNPFIN.Text <> "") And Not IsNumeric(ioNPFIN.Text) Then Exit Sub

If (Trim(ioNUMPED.Text) = "") And (Trim(ioNPINI.Text <> "") Or Trim(ioNPFIN.Text <> "")) Then Exit Sub

'ver el estado que tiene el pedido
tmpest = devuelve_campo("SELECT ESTADO FROM CABPEDPRO WHERE NUMERO = " & ioNUMPED.Text & " AND ALMORIG = " & AlmacenActual, locCnn)

If tmpest = "@" Then
    MsgBox "No se encuentra pedido en la base de datos", vbInformation, titulo
    Exit Sub
End If

 'cerrar la conexion y abrirla para hacer la transaccion
  With locCnn
    If .State = 1 Then .Close
        .CursorLocation = adUseServer
        .Open strLocCnn
        .BeginTrans
        entrans = True
  End With

        Set rcped = New ADODB.Recordset
        Set rcab = New ADODB.Recordset
        
        
        'si son varios pedidos, recorrer la serie
        If ioNUMPED.Text = "" Then
            pinicial = ioNPINI.Text
            pfinal = ioNPFIN.Text
        Else
            pinicial = ioNUMPED.Text
            pfinal = ioNUMPED.Text
        End If
        
        'recorrer el rango
        '---------------------------------------------------------------------
        For var = pinicial To pfinal
        
            rcab.Open "SELECT CODPTRN FROM CABPEDPRO WHERE NUMERO = " & var & " AND ALMORIG = " & AlmacenActual, locCnn, adOpenStatic, adLockReadOnly
            
            'si no tiene transferencia asociada
            If rcab.fields(0) = 0 Then
                        
                rcped.Open "SELECT * from DETPEDPRO where NUMERO = " & var & " AND ALMORIG = " & AlmacenActual, locCnn, adOpenDynamic, adLockOptimistic
                 
                Do Until rcped.EOF
                
                    If rcped.fields("METIDO") = True Then
                        Call stock(rcped.fields("CODART"), rcped.fields("TEMPOR"), rcped.fields("CODTALLA"), rcped.fields("CODCOL"), AlmacenActual, rcped.fields("UNIDADES"), False, locCnn)
                        rcped.fields("METIDO") = 0
                        rcped.Update
                    End If
                    
                    rcped.MoveNext
            
                Loop
                rcped.Close
                
            
            'si tiene transferencia, comprobar que no este con estado > 1
            'si asi fuera, no se puede deshacer (ya ha sido aceptada o cancelada)
            Else
            
                tmpvar = devuelve_campo("SELECT ESTADO FROM PTRANS WHERE CODIGO = " & rcab.fields(0) & " AND CODALMORIG = " & AlmacenActual, locCnn)
                
                If tmpvar = "@" Then
                    lblstatus.Caption = "Ha ocurrido un error al obtener el estado de la transferencia"
                    Exit For
                End If
                
                If tmpvar > 1 Then
                    lblstatus.Caption = "La transferencia de ese pedido ya no esta PENDIENTE. Imposible deshacer."
                    Exit For
                End If
                
                'poner la transferencia como cancelada
                Call anular_transferencia_pendiente(rcab.fields(0), AlmacenActual, locCnn)
                
                
                'locCnn.Execute "UPDATE PTRANS SET ESTADO = 3 WHERE NUMERO = " & rcab.fields(0) & " AND CODALMORIG = " & AlmacenActual
            
            
            End If
            rcab.Close
            
        Next var
        '---------------------------------------------------------------------
                 
        Set rcped = Nothing
                                       
  With locCnn
        .CursorLocation = adUseServer
        
        If tmpvar <> "@" Then
            .CommitTrans
        Else
            .RollbackTrans
        End If
        
        .Close
        entrans = False
        .CursorLocation = adUseClient
        .Open strLocCnn
  End With
    
  MsgBox "El pedido se ha desecho satisfactoriamente", vbInformation, titulo
    
   On Error GoTo 0
   Exit Sub

cbDeshacerPedido_Click_Error:
    
    If entrans Then
        locCnn.RollbackTrans
        locCnn.Close
    End If
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbDeshacerPedido_Click de Formulario frmTrabaDatos"

End Sub

Private Sub cbTrabajarPedido_Click()
Dim tmpest As Variant

'filtros
If (Trim(ioNUMPED.Text = "") And Trim(ioNPINI.Text = "") And Trim(ioNPFIN.Text = "")) Then Exit Sub
If Trim(ioNUMPED.Text <> "") And Not IsNumeric(ioNUMPED.Text) Then Exit Sub
If Trim(ioNPINI.Text <> "") And Not IsNumeric(ioNPINI.Text) Then Exit Sub
If Trim(ioNPFIN.Text <> "") And Not IsNumeric(ioNPFIN.Text) Then Exit Sub

If (Trim(ioNUMPED.Text) = "") And (Trim(ioNPINI.Text <> "") Or Trim(ioNPFIN.Text <> "")) Then Exit Sub

'ver el estado que tiene el pedido
tmpest = devuelve_campo("SELECT ESTADO FROM CABPEDPRO WHERE NUMERO = " & ioNUMPED.Text & " AND ALMORIG = " & AlmacenActual, locCnn)

If tmpest = "@" Then
    MsgBox "No se encuentra pedido en la base de datos", vbInformation, titulo
    Exit Sub
End If

With frmPedProv
    .trabajar_con_pedido = True
    .codigo_almacen = AlmacenActual
    .NUMERO_PEDIDO = ioNUMPED.Text
    .Show
End With

End Sub

Private Sub cmBorrarVenta_Click()

   On Error GoTo cmBorrarVenta_Click_Error

If comprueba_ticket = True Then Exit Sub

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

lblstatus.Caption = "Deshaciendo venta..."
DoEvents

'deshace la venta actual
Call Deshacer_Venta(Left(Trim(ioTICKET.Text), Len(Trim(ioTICKET.Text)) - 3), Right(Trim(ioTICKET.Text), 3), locCnn)

lblstatus.Caption = "Borrando registros..."
DoEvents

'borrar la venta ...
locCnn.Execute "DELETE FROM DETVENTA WHERE CODVEN = " & Left(ioTICKET.Text, Len(ioTICKET.Text) - 3) & " AND CODCAJA = " & Right(Trim(ioTICKET.Text), 3)
DoEvents
locCnn.Execute "DELETE FROM CABVENTA WHERE CODIGO = " & Left(ioTICKET.Text, Len(ioTICKET.Text) - 3) & " AND CODCAJA = " & Right(Trim(ioTICKET.Text), 3)


lblstatus.Caption = "La venta " & ioTICKET.Text & " ha sido eliminada"
DoEvents

ioTICKET.Text = ""
ioTICKET.SetFocus

   On Error GoTo 0
   Exit Sub

cmBorrarVenta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmBorrarVenta_Click de Formulario frmTrabaDatos"

End Sub

Private Sub cmPasarPendientePedido_Click()
Dim tmpest As Variant

   On Error GoTo cmPasarPendientePedido_Click_Error

'filtros
If (Trim(ioNUMPED.Text = "") And Trim(ioNPINI.Text = "") And Trim(ioNPFIN.Text = "")) Then Exit Sub
If Trim(ioNUMPED.Text <> "") And Not IsNumeric(ioNUMPED.Text) Then Exit Sub
If Trim(ioNPINI.Text <> "") And Not IsNumeric(ioNPINI.Text) Then Exit Sub
If Trim(ioNPFIN.Text <> "") And Not IsNumeric(ioNPFIN.Text) Then Exit Sub

If (Trim(ioNUMPED.Text) = "") And (Trim(ioNPINI.Text <> "") Or Trim(ioNPFIN.Text <> "")) Then Exit Sub


With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

'ver el estado que tiene el pedido
tmpest = devuelve_campo("SELECT ESTADO FROM CABPEDPRO WHERE NUMERO = " & ioNUMPED.Text & " AND ALMORIG = " & AlmacenActual, locCnn)

If tmpest = "@" Then
    MsgBox "No se encuentra pedido en la base de datos", vbInformation, titulo
    Exit Sub
End If

locCnn.Execute "UPDATE CABPEDPRO SET ESTADO = 1 WHERE NUMERO = " & ioNUMPED.Text & " AND ALMORIG = " & AlmacenActual

MsgBox "Se ha pasado el pedido " & ioNUMPED.Text & " a pendiente", vbInformation, titulo

ioNUMPED.SetFocus

   On Error GoTo 0
   Exit Sub

cmPasarPendientePedido_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmPasarPendientePedido_Click de Formulario frmTrabaDatos"

End Sub

Private Sub cmPasarPendienteVenta_Click()
Dim rc As ADODB.Recordset
Dim entrans As Boolean
Dim cn As ADODB.Connection
Dim var As Long

   On Error GoTo cmPasarPendienteVenta_Click_Error

Set cn = New ADODB.Connection

With cn
    .CursorLocation = adUseServer
    .Open strLocCnn
    .BeginTrans
End With

entrans = True
    

If ioFECHA.Text = "" Then

    If comprueba_ticket = True Then

        If entrans Then cn.RollbackTrans
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    
        Exit Sub
        
    End If
    
    lblstatus.Caption = "Deshaciendo venta ..."
    DoEvents
    
    'deshace la venta actual
    Call Deshacer_Venta(Left(Trim(ioTICKET.Text), Len(Trim(ioTICKET.Text)) - 3), Right(Trim(ioTICKET.Text), 3), cn)

    lblstatus.Caption = "La venta se ha pasado a pendiente de manera correcta ..."
    DoEvents

    ioTICKET.SetFocus

Else
    
    'preguntar si quiere deshacer las de todo el dia
    If MsgBox("Va a deshacer las ventas del dia: " & ioFECHA.Text & ". ¿Esta seguro?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
    
    Set rc = New ADODB.Recordset
    
    rc.Open "Select CODIGO FROM CABVENTA where ESTADO = 1 AND CODCAJA =" & CajaActual & " AND FHORA = '" & Format(ioFECHA.Text, "yyyymmdd") & "'", locCnn, adOpenStatic, adLockReadOnly
    
    'si no hay registros
    If rc.BOF And rc.EOF Then
    
        MsgBox "No se encuentran ventas para ese dia", vbExclamation, titulo
        
        rc.Close
        Set rc = Nothing
        
        If entrans Then cn.RollbackTrans

        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    
        Exit Sub
        
    End If
    
    'deshacer todas las ventas ...
    Do Until rc.EOF
    
        lblstatus.Caption = "Deshaciendo venta (" & rc.fields(0) & ")..."
        DoEvents
        'deshace la venta actual
        Call Deshacer_Venta(rc.fields(0), CajaActual, cn)
        
        rc.MoveNext
    
    Loop
       
End If


If entrans Then
    cn.CommitTrans
End If

    cn.Close
    Set cn = Nothing

   On Error GoTo 0
   Exit Sub

cmPasarPendienteVenta_Click_Error:

    If entrans Then
        cn.RollbackTrans
        entrans = False
    End If

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmPasarPendienteVenta_Click de Formulario frmTrabaDatos"

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : cbCambiarCaja_Click
' Fecha/Hora  : 25/10/2004 22:25
' Autor       : JCASTILLO
' Propósito   :
'
'---------------------------------------------------------------------------------------
Private Sub cbCambiarCaja_Click()
Dim cn As ADODB.Connection
Dim entrans As Boolean

   On Error GoTo cbCambiarCaja_Click_Error

        If ioFECHA.Text = "" Then
        
            lblstatus.Caption = "Se necesita una fecha"
            Exit Sub
            
        End If
        
        If cbCAJAS.Text = "" Then
        
            lblstatus.Caption = "Se necesita una caja de destino"
            Exit Sub
                
        End If
        
        'pasar las ventas a pendiente ...
        'Call cmPasarPendienteVenta_Click
               
        Set cn = New ADODB.Connection
        
        With cn
            .ConnectionString = strLocCnn
            .CursorLocation = adUseServer
            .Open
            .BeginTrans
        End With
        
        entrans = True
        
        lblstatus.Caption = "Procesando ..."
                
        'si devuelve 1, es q hubo error
        If recodifica_ventas(CDate(ioFECHA.Text), cn) = 1 Then
            cn.RollbackTrans
            entrans = False
             lblstatus.Caption = "Ha ocurrido un error al procesar los datos."
            Exit Sub
        End If
               
        'aceptar cambios ...
        If entrans Then
            cn.CommitTrans
        End If
        
        lblstatus.Caption = "La operación se ha finalizado correctamente."
            
   Exit Sub

cbCambiarCaja_Click_Error:

    If entrans Then
        cn.RollbackTrans
    End If

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCambiarCaja_Click de Formulario frmTrabaDatos"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : recodifica_ventas
' Fecha/Hora  : 26/10/2004 20:52
' Autor       : JCASTILLO
' Propósito   : Recodifica y cambia a la caja actual, las ventas de una fecha dada
' Devuelve  1 si hubo error (para hacer el rollback y cancelar)
'---------------------------------------------------------------------------------------
Private Function recodifica_ventas(fecha As Date, conexion As ADODB.Connection) As Byte
Dim rc As ADODB.Recordset
Dim rcd As ADODB.Recordset
Dim tmpcodigo As Long
Dim caja_origen As Byte
Dim entrans As Boolean

   On Error GoTo recodifica_ventas_Error


' 1º, añadir un registro en cabventa con la nueva cabecera de la venta
' (copiando los datos de la antigua)
' 2º, cambiar el codigo de los registros de detventa a la nueva cabecera
' 3º, borrar le registro de cabecera viejo

        Set rc = New ADODB.Recordset
        Set rcd = New ADODB.Recordset
        rc.Open "SELECT * FROM CABVENTA WHERE FHORA = '" & Format(fecha, "yyyymmdd") & "' AND (ESTADO = 0) ORDER BY CODIGO", conexion, adOpenDynamic, adLockOptimistic
                
        If rc.EOF And rc.BOF Then
            rc.Close
            Set rc = Nothing
            Exit Function
        End If
        
        With rc
            
            'obtener nuevo codigo disponible
            tmpcodigo = devuelve_codigo(conexion)
            
            Do Until .EOF
                
                If caja_origen = 0 Then caja_origen = .fields("CODCAJA")
                
                
                If IsNull(.fields("CODCLI")) Then .fields("CODCLI") = 0
                If IsNull(.fields("CAJACLI")) Then .fields("CAJACLI") = 0
                If IsNull(.fields("COMEN")) Then .fields("COMEN") = ""
                
                'insertar el nuevo registro de cabecera ...
                Call inserta_cabventa(tmpcodigo, .fields("CODPER"), CajaActual, .fields("CODCLI"), .fields("CAJACLI"), .fields("CAJADES"), .fields("SUBTOT"), .fields("IVATOT"), .fields("RETOT"), .fields("IMP_PRIMERA"), .fields("IMP_SEGUNDA"), .fields("COMISION"), 1, .fields("FCOBRO"), .fields("COMEN"), .fields("FHORA"), .fields("FMODI"), conexion)
                DoEvents
                
                If rcd.State = 1 Then rcd.Close
                rcd.Open "SELECT * FROM DETVENTA WHERE CODCAJA = " & caja_origen & " AND CODVEN = " & .fields("CODIGO") & " ORDER BY LINEA", conexion, adOpenDynamic, adLockOptimistic
                
                'si hay registros ...
                If Not (rcd.EOF And rcd.BOF) Then
                    
                    'actualizar todos los registros del detalle
                    Do Until rcd.EOF
                            
                            'quitar de stock para el articulo actual
                            stock rcd.fields("CODART"), rcd.fields("TEMPOR"), rcd.fields("CODTALLA"), rcd.fields("CODCOL"), AlmacenActual, rcd.fields("UNIDADES"), False, conexion
                            rcd.fields("CODVEN") = tmpcodigo
                            rcd.fields("CODCAJA") = CajaActual
                            rcd.Update
                            DoEvents
                            rcd.MoveNext
                    Loop
                    
                End If
                
                'recodificar los arreglos para la venta
                
                'si devuelve 1 (condicion de error) salir y deshacer la operación.
                If recodifica_arreglos(.fields("CODIGO"), tmpcodigo, .fields("CODCAJA"), conexion) = 1 Then
                
                    recodifica_ventas = 1
                    
                    If rc.State = 1 Then rc.Close
                    Set rc = Nothing
                    If rcd.State = 1 Then rcd.Close
                    Set rc = Nothing
                    
                    Exit Function
                
                End If
                
                
                'recodificar las devoluciones para la venta
                'si devuelve 1 (condicion de error) salir y deshacer la operación.
                If recodifica_devoluciones(.fields("CODIGO"), tmpcodigo, .fields("CODCAJA"), conexion) = 1 Then
                
                    recodifica_ventas = 1
                    
                    If rc.State = 1 Then rc.Close
                    Set rc = Nothing
                    If rcd.State = 1 Then rcd.Close
                    Set rc = Nothing
                    
                    Exit Function
                
                End If
                
             

                
                
                'actualizar el registro de cabecera
                '.fields("CODIGO") = tmpcodigo
                '.fields("CODCAJA") = CajaActual
                '.fields("ESTADO") = 1
                'poner como aceptada, porque ya se han descontado las
                'unidades del stock
                
                'BORRAR EL REGISTRO ACTUAL (se ha creado el nuevo ya en la cabecera)
                .Delete
                
                DoEvents
                
                If Not .EOF Then .MoveNext
                
                'incrementar el nuevo codigo disponible
                tmpcodigo = tmpcodigo + 1
        
            Loop
            
        End With
        
        
        'recodificar todas las devoluciones sueltas de esa fecha (las q no estan asignadas a
        'ninguna venta)
        If recodifica_devoluciones(0, 0, caja_origen, conexion, CDate(ioFECHA.Text)) = 1 Then
            recodifica_ventas = 1
            
           If rc.State = 1 Then rc.Close
           Set rc = Nothing
           If rcd.State = 1 Then rcd.Close
           Set rc = Nothing
                        
           Exit Function
        End If


        If rc.State = 1 Then rc.Close
        Set rc = Nothing
        If rcd.State = 1 Then rcd.Close
        Set rc = Nothing
        
        recodifica_ventas = 0

   On Error GoTo 0
   Exit Function

recodifica_ventas_Error:

    recodifica_ventas = 1
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento recodifica_ventas de Formulario frmTrabaDatos"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : recodifica_arreglos
' Fecha/Hora    : 27/10/2004 10:25
' Autor         : JCastillo
' Propósito     :   Devuelve 1 si hubo algun error
'---------------------------------------------------------------------------------------
'
Private Function recodifica_arreglos(CODIGO_VENTA As Long, nuevo_codigo_venta, codigo_caja As Byte, conexion As ADODB.Connection) As Byte
Dim rc As ADODB.Recordset
Dim tmpvar As Variant

   On Error GoTo recodifica_arreglos_Error
   
   Set rc = New ADODB.Recordset

        'abrir todos los arreglos para esa venta ...
        rc.Open "SELECT * FROM ARREGLOS WHERE MBAJA = 0 AND CODVEN = " & CODIGO_VENTA & " AND CODCAJ = " & codigo_caja & " ORDER BY ID", conexion, adOpenDynamic, adLockOptimistic
        
        If rc.EOF And rc.BOF Then
            recodifica_arreglos = 0
            Exit Function
        End If
        
        'sacar el ultimo id de arreglos para la caja actual
        tmpvar = devuelve_campo("SELECT MAX(ID) FROM ARREGLOS WHERE CODCAJ = " & CajaActual, conexion)
    
        'incrementar código
        If tmpvar = "@" Then
            tmpvar = 1
        Else
            If IsNumeric(tmpvar) Then
                tmpvar = tmpvar + 1
            Else
                tmpvar = 1
            End If
        End If
        
        
        Do Until rc.EOF
        
            rc.fields("ID") = tmpvar
            rc.fields("CODCAJ") = CajaActual
            rc.fields("CODVEN") = nuevo_codigo_venta
            rc.fields("ESTADO") = 2
            rc.Update
            DoEvents
            
            rc.MoveNext
            tmpvar = tmpvar + 1
            
        Loop
        
   rc.Close
   Set rc = Nothing
   
   On Error GoTo 0
   Exit Function

recodifica_arreglos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento recodifica_arreglos de Formulario frmTrabaDatos"
End Function


'---------------------------------------------------------------------------------------
' Procedimiento : recodifica_devoluciones
' Fecha/Hora    : 27/10/2004 10:39
' Autor         : JCastillo
' Propósito     : Devuelve 1 si hubo algun error
'                     si entra codigo_venta = 0, se intentarán recodificar todas las ventas
'                     de la fecha dada, que tengan el codigo de venta a 0
'---------------------------------------------------------------------------------------
'
Private Function recodifica_devoluciones(CODIGO_VENTA As Long, nuevo_codigo_venta, codigo_caja As Byte, conexion As ADODB.Connection, Optional fecha As Date) As Byte
Dim rc As ADODB.Recordset
Dim tmpvar As Variant
Dim tmpvales As Variant
Dim ffinal As Date


   On Error GoTo recodifica_devoluciones_Error
   
   Set rc = New ADODB.Recordset

        'una sola venta
        If CODIGO_VENTA > 0 Then
            'abrir todos los arreglos para esa venta ...
            rc.Open "SELECT * FROM DEVOL WHERE MBAJA = 0 AND CODVEN = " & CODIGO_VENTA & " AND CODCAJA = " & codigo_caja & " ORDER BY CODIGO", conexion, adOpenDynamic, adLockOptimistic
        'todas las devoluciones de la fecha
        Else
           ffinal = DateAdd("d", 1, fecha)
           rc.Open "SELECT * FROM DEVOL WHERE MBAJA = 0 AND CODVEN = 0 AND FMODI >= '" & Format(fecha, "yyyymmdd") & "'" & " AND FMODI <= '" & Format(ffinal, "yyyymmdd") & "' AND CODCAJA = " & codigo_caja & " ORDER BY CODIGO", conexion, adOpenDynamic, adLockOptimistic
        End If
        
        
        If rc.EOF And rc.BOF Then
            recodifica_devoluciones = 0
            Exit Function
        End If
        
        'sacar el ultimo id de arreglos para la caja actual
        tmpvar = devuelve_campo("SELECT MAX(CODIGO) FROM DEVOL WHERE CODCAJA = " & CajaActual, conexion)
    
        'incrementar código
        If tmpvar = "@" Then
            tmpvar = 1
        Else
            If IsNumeric(tmpvar) Then
                tmpvar = tmpvar + 1
            Else
                tmpvar = 1
            End If
        End If
        
        'sacar el ultimo id de arreglos para la caja actual
        tmpvales = devuelve_campo("SELECT MAX(CODIGO) FROM VALES WHERE CODCAJA = " & CajaActual, conexion)
        
                'incrementar código
        If tmpvales = "@" Then
            tmpvales = 1
        Else
            If IsNumeric(tmpvales) Then
                tmpvales = tmpvales + 1
            Else
                tmpvales = 1
            End If
        End If
    
        Do Until rc.EOF
            
            'ACTUALIZAR EL VALE SI TIENE LA DEVOLUCION
            '(cambiar codigo, codcaja y poner el nuevo codigo de venta)
            If CODIGO_VENTA > 0 Then
            conexion.Execute "UPDATE VALES SET CODVEN = " & nuevo_codigo_venta & ", CODIGO = " & tmpvales & ", CODCAJA = " & CajaActual & " WHERE CODVEN =" & CODIGO_VENTA & " AND CODCAJA =" & codigo_caja
            End If
            
            'descontar la unidad de la devolución para el almacén actual
            Call stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, 1, False, conexion)
            
            rc.fields("CODIGO") = tmpvar
            rc.fields("CODCAJA") = CajaActual
            rc.fields("CODVEN") = nuevo_codigo_venta
            rc.Update
            DoEvents
            
            rc.MoveNext
            tmpvar = tmpvar + 1
            tmpvales = tmpvales + 1
            
                      
        Loop
        
        
   rc.Close
   Set rc = Nothing

   On Error GoTo 0
   Exit Function

recodifica_devoluciones_Error:

    recodifica_devoluciones = 1
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento recodifica_devoluciones de Formulario frmTrabaDatos"
End Function


'---------------------------------------------------------------------------------------
' Subrutina   : devuelve_codigo
' Fecha/Hora  : 26/10/2004 20:45
' Autor       : JCASTILLO
' Propósito   :
'
'---------------------------------------------------------------------------------------
Private Function devuelve_codigo(conexion As ADODB.Connection) As Long
Dim tmpvar As Variant
'devuelve el ultimo codigo para la venta

   On Error GoTo devuelve_codigo_Error

    tmpvar = devuelve_campo("SELECT MAX(CODIGO) FROM CABVENTA WHERE CODCAJA = " & CajaActual, conexion)
    
    'incrementar código
    If tmpvar = "@" Then
        tmpvar = 1
    Else
        If IsNumeric(tmpvar) Then
            tmpvar = tmpvar + 1
        Else
            tmpvar = 1
        End If
    End If
    
    devuelve_codigo = CLng(tmpvar)

   On Error GoTo 0
   Exit Function

devuelve_codigo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento devuelve_codigo de Formulario frmTrabaDatos"

End Function


'---------------------------------------------------------------------------------------
' Procedimiento : inserta_cabventa
' Fecha/Hora    : 27/10/2004 16:33
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
Private Sub inserta_cabventa(codigo As Long, codper As Integer, codcaja As Byte, codcli As Variant, cajacli As Variant, cajades As Boolean, subtot As Double, _
ivatot As Double, retot As Double, imp_primera As Double, imp_segunda As Double, comision As Double, estado As Byte, _
fcobro As Byte, comen As String, fhora As Date, fmodi As Date, conexion As ADODB.Connection)
Dim rc As New ADODB.Recordset

   On Error GoTo inserta_cabventa_Error
   
   With rc
    .Open "SELECT * FROM CABVENTA", conexion, adOpenDynamic, adLockOptimistic
    .AddNew
    
    .fields("CODIGO") = codigo
    .fields("CODPER") = codper
    .fields("CODCAJA") = codcaja
    .fields("CODCLI") = codcli
    .fields("CAJACLI") = cajacli
    .fields("CAJADES") = cajades
    .fields("SUBTOT") = subtot
    .fields("IVATOT") = ivatot
    .fields("RETOT") = retot
    .fields("imp_primera") = imp_primera
    .fields("imp_segunda") = imp_segunda
    .fields("comision") = comision
    .fields("estado") = estado
    .fields("fcobro") = fcobro
    .fields("comen") = comen
    .fields("fhora") = fhora
    .fields("fmodi") = fmodi
    
    .Update
    
   End With
    
    rc.Close
    Set rc = Nothing
    
   On Error GoTo 0
   Exit Sub

inserta_cabventa_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento inserta_cabventa de Formulario frmTrabaDatos"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : devuelve_codigo_linea
' Fecha/Hora  : 26/10/2004 21:12
' Autor       : JCASTILLO
' Propósito   :
'
'---------------------------------------------------------------------------------------
'Private Function devuelve_codigo_linea(codigo_venta As Long) As Long
'Dim tmpvar As Variant
'devuelve el ultimo codigo de linea para el detalle de la venta


 '  On Error GoTo devuelve_codigo_linea_Error

  '  tmpvar = devuelve_campo("SELECT MAX(LINEA) FROM DETVENTA WHERE CODCAJA = " & CajaActual & " AND CODVEN = " & codigo_venta, locCnn)
    
    'incrementar código
   ' If tmpvar = "@" Then
  '      tmpvar = 1
  '  Else
'        If IsNumeric(tmpvar) Then
 '           tmpvar = tmpvar + 1
 '       Else
  '          tmpvar = 1
  '      End If
  '  End If
    
 '   devuelve_codigo_linea = CLng(tmpvar)


  ' On Error GoTo 0
  ' Exit Function

'devuelve_codigo_linea_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento devuelve_codigo_linea de Formulario frmTrabaDatos"

'End Function



Private Sub Form_Load()

   On Error GoTo Form_Load_Error

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm

  With locCnn
  If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
  End If
  End With
  
  With ioNUMPED
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 13
  End With
  
  With ioFECHA
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
  End With
  
  With ioTICKET
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 13
  End With
  
   'Cargar el micombo cajas
  With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
    DoEvents
  End With


   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario frmTrabaDatos"
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmTrabaDatos = Nothing
End Sub





Private Sub ioFECHA_Validate(Cancel As Boolean)

    If ioFECHA.Text <> "" Then
        ioTICKET.Text = ""
    End If

End Sub

Private Sub ioNPINI_Validate(Cancel As Boolean)

'limpiar el numero si va a ser un rango
If ioNUMPED.Text <> "" And ioNPINI.Text <> "" Then ioNUMPED.Text = ""

End Sub

Private Sub ioNPfin_Validate(Cancel As Boolean)

'limpiar el numero si va a ser un rango
If ioNUMPED.Text <> "" And ioNPFIN.Text <> "" Then ioNUMPED.Text = ""

End Sub


Private Sub ioNUMPED_Validate(Cancel As Boolean)

'limpiar
If ioNUMPED.Text <> "" Then
    ioNPFIN.Text = ""
    ioNPINI.Text = ""
End If

End Sub
