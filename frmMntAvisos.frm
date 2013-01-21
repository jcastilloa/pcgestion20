VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMntAvisos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9615
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   0
      Top             =   4665
      Width           =   9615
      _extentx        =   16960
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntAvisos.frx":0000
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid lstAvisos 
      Height          =   4650
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9615
      _cx             =   16960
      _cy             =   8202
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
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
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMntAvisos.frx":002E
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
End
Attribute VB_Name = "frmMntAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmMntAvisos
' Fecha/Hora  : 11/07/2004 21:41
' Autor       : JCASTILLO
' Propósito   : Avisos del programa (pagos a proveedores para este mes, etc)
'---------------------------------------------------------------------------------------
Option Explicit

Dim prime As Boolean
Dim usarbeep As Boolean

Private Sub Form_Activate()
Const segundos = 10
Dim numseg As Long

If Not prime Then

For numseg = 1 To segundos
            Espera 1
            If usarbeep Then Beep
            lblstatus.Caption = "Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
Next numseg

prime = True
Unload Me

End If

End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, Separacion_MDIForm '(Screen.Height - Height) \ 2

With lstAvisos
    .Cols = 10
    .Rows = 1
    .TextMatrix(0, 1) = "Aviso"
    .TextMatrix(0, 2) = "Tipo"
    .TextMatrix(0, 3) = "Prov."
    .TextMatrix(0, 4) = "Cuota"
    .TextMatrix(0, 5) = "Total"
    .TextMatrix(0, 6) = "Pagado"
    .TextMatrix(0, 7) = "Pendi."
    .TextMatrix(0, 8) = "Meses"
    .TextMatrix(0, 9) = "Comen."
    
    .ColFormat(4) = "Currency"
    .ColFormat(5) = "Currency"
    .ColFormat(6) = "Currency"
    .ColFormat(7) = "Currency"
End With

Call consulta_pagos

With lstAvisos
    .AutoSize 1, .Cols - 1
End With


If lstAvisos.Rows = 1 Then
'    lstAvisos.AddItem "No hay avisos pendientes"
Else
    usarbeep = True 'que pite si hay avisos
End If

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : consulta_pagos
' Fecha/Hora  : 11/07/2004 21:55
' Autor       : JCASTILLO
' Propósito   : consultar si hay pagos, cuya ultimo MES de pago + MESES = mes actual
'               es decir, que le toque pagar este mes
'---------------------------------------------------------------------------------------
Private Sub consulta_pagos()
Dim rc As New ADODB.Recordset
Dim rcdet As New ADODB.Recordset
Dim fecha As Date
Dim fechaact As Date
Dim DspProv As String
Dim DspTipo As String

   On Error GoTo consulta_pagos_Error

        'coger los pagos estado = 1 (parcial) y meses > 0 (período de pago)
        'para ir comprobando
        rc.Open "SELECT CODIGO, CODCAJA, IMPORTE, PAGADO, MESES, CUOTA, TIPOPAGO, CODPROV, DESCRIPCION, FMODI FROM PAGOS WHERE (ESTADO = 1) AND (MESES >0)", locCnn, adOpenStatic, adLockReadOnly

        If (rc.BOF And rc.EOF) Then Exit Sub
        
        Do Until rc.EOF
            'coger de los pagos        estado = 1 (parcial), el ultimo mes de
            'pago del detalle.
            If rcdet.State = 1 Then rcdet.Close
            rcdet.Open "SELECT TOP 1 FMODI FROM DETPAGOS WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & rc.fields("CODCAJA") & " ORDER BY FMODI DESC", locCnn, adOpenStatic, adLockReadOnly
            
            'si nunca se han echo pagos, coger como fecha inicial la fecha del pago
            If (rcdet.EOF And rcdet.BOF) Then
                'añadir los meses del campo período a la fecha inicial del pago
                fecha = DateAdd("m", rc.fields("MESES"), rc.fields("FMODI"))
            Else
                'añadir los meses del campo período
                fecha = DateAdd("m", rc.fields("MESES"), rcdet.fields("FMODI"))
            End If
            
            'poner a dia uno de ese mes para poder comparar bien
            fecha = CDate("1/" & Month(fecha) & "/" & Year(fecha))
            
            'lo mismo para la fecha actual
            fechaact = CDate("1/" & Month(Date) & "/" & Year(Date))
            
            'si el mes coincide, añadir una nueva linea en el Lst
            If fecha <= fechaact Then
            
                'sacar descripciones
                If rc.fields("TIPOPAGO") > 0 Then
                    DspTipo = Trim(devuelve_campo("SELECT DESCRIPCION FROM MAPAG WHERE CODIGO = " & rc.fields("CODPROV"), locCnn))
                End If
                
                If rc.fields("CODPROV") > 0 Then
                    DspProv = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & rc.fields("CODPROV"), locCnn))
                End If
    
                With lstAvisos
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = "Pendiente"
                    .TextMatrix(.Rows - 1, 2) = DspTipo
                    .TextMatrix(.Rows - 1, 3) = DspProv
                    .TextMatrix(.Rows - 1, 4) = rc.fields("CUOTA")
                    .TextMatrix(.Rows - 1, 5) = rc.fields("IMPORTE")
                    .TextMatrix(.Rows - 1, 6) = rc.fields("PAGADO")
                    .TextMatrix(.Rows - 1, 7) = rc.fields("IMPORTE") - rc.fields("PAGADO")
                    .TextMatrix(.Rows - 1, 8) = rc.fields("MESES")
                    .TextMatrix(.Rows - 1, 9) = rc.fields("DESCRIPCION")
                End With
                
                          
            End If
           
        
        rc.MoveNext
        
        Loop

    rc.Close
    Set rc = Nothing
    rcdet.Close
    Set rcdet = Nothing
    DspProv = ""
    DspTipo = ""

   On Error GoTo 0
   Exit Sub
     

consulta_pagos_Error:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento consulta_pagos de Formulario frmMntAvisos"
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set frmMntAvisos = Nothing

End Sub
