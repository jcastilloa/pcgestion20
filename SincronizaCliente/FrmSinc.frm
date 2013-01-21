VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSinc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sincroniza Cliente"
   ClientHeight    =   3660
   ClientLeft      =   4005
   ClientTop       =   3555
   ClientWidth     =   7470
   ClipControls    =   0   'False
   Icon            =   "FrmSinc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7470
   Begin VB.TextBox ioFFIN 
      Height          =   360
      Left            =   2108
      TabIndex        =   5
      Top             =   2040
      Width           =   1230
   End
   Begin VB.TextBox ioFINI 
      Height          =   360
      Left            =   2108
      TabIndex        =   4
      Top             =   1515
      Width           =   1230
   End
   Begin VB.ListBox lstTablas 
      Height          =   960
      Left            =   3983
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   1500
      Width           =   2250
   End
   Begin VB.TextBox ioDBCliente 
      Height          =   360
      Left            =   5235
      TabIndex        =   1
      Top             =   465
      Width           =   2130
   End
   Begin VB.TextBox ioDBServer 
      Height          =   360
      Left            =   5235
      TabIndex        =   3
      Top             =   990
      Width           =   2130
   End
   Begin VB.TextBox ioServer 
      Height          =   360
      Left            =   1515
      TabIndex        =   2
      Top             =   990
      Width           =   2130
   End
   Begin VB.TextBox ioCliente 
      Height          =   360
      Left            =   1515
      TabIndex        =   0
      Top             =   465
      Width           =   2130
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Comenzar Operación"
      Default         =   -1  'True
      Height          =   540
      Left            =   2745
      TabIndex        =   6
      Top             =   2580
      Width           =   2100
   End
   Begin MSForms.Frame Frame1 
      Height          =   645
      Left            =   2670
      OleObjectBlob   =   "FrmSinc.frx":06D2
      TabIndex        =   15
      Top             =   2535
      Width           =   2250
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.Fin"
      Height          =   240
      Left            =   1238
      TabIndex        =   17
      Top             =   2100
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.Inicio"
      Height          =   225
      Left            =   1283
      TabIndex        =   16
      Top             =   1590
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incluir:"
      Height          =   255
      Left            =   3285
      TabIndex        =   14
      Top             =   1590
      Width           =   780
   End
   Begin VB.Label lblCaja 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   15
      Width           =   7455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE cliente"
      Height          =   225
      Left            =   3825
      TabIndex        =   10
      Top             =   540
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE server"
      Height          =   240
      Left            =   3825
      TabIndex        =   9
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IP/HOSTNAME server"
      Height          =   405
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP/HOSTNAME cliente"
      Height          =   405
      Left            =   90
      TabIndex        =   7
      Top             =   465
      Width           =   1365
   End
End
Attribute VB_Name = "FrmSinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : FrmSinc
' Fecha/Hora : 04/08/2004 16:40
' Autor      : JCastillo
' Propósito  :   Sincronizar los datos del cliente y del server, cuando se ha ROTO el enlace
'                    y el cliente ha comenzado a funcionar de manera asincrona con el server (no evia ni recibe)
'---------------------------------------------------------------------------------------

'por CODCAJA
'CABVENTA
'DETVENTA
'VALES
'DEVOL
'ARREGLOS
'MOVCAJA
'CIERREDIA
'CIERREMES
'CIERREANO
'PAGOS
'DETPAGOS
'CABPRESTA
'DETPRESTA

'por CODALM
'STOCK
'PTRANS  (estados)

Option Explicit

Dim cno As New ADODB.Connection
Dim cnd As New ADODB.Connection
Dim cliStr As String
Dim srvStr As String

Dim sCnn    As String

Const strCnnMdb = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source="
Const logfile = "log.mdb"


'para recoger las variables maestras ...
Dim CajaActual As Byte
Dim AlmacenActual As Byte

Dim entrans As Boolean
'---------------------------------------------------------------------------------------
' Procedimiento : cmd_Click
' Fecha/Hora    : 05/08/2004 09:29
' Autor         : JCastillo
' Propósito     :  Proceso principal
'---------------------------------------------------------------------------------------
'
Private Sub cmd_Click()
Dim rco As New ADODB.Recordset
Dim rcd As New ADODB.Recordset

Dim enFechas As Boolean

   'On Error GoTo cmd_Click_Error
   
If MsgBox("Se va a proceder al envío de datos al servidor, ¿desea continuar?", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

With ioCliente
    If .Text = "" Then
        lblStatus.Caption = "No se permite IP Cliente en blanco"
        Exit Sub
    End If
End With

With ioServer
    If .Text = "" Then
        lblStatus.Caption = "No se permite IP Server en blanco"
        Exit Sub
    End If
End With

With ioDBCliente
    If .Text = "" Then
        lblStatus.Caption = "No se permite DB Cliente en blanco"
        Exit Sub
    End If
End With

With ioDBServer
    If .Text = "" Then
        lblStatus.Caption = "No se permite DB Server en blanco"
        Exit Sub
    End If
End With

If Dir(App.Path & "\Config.pcg") = "" Then
    MsgBox "No se encuentra la configuración de PCGestión (Config.pcg) en el directorio actual del programa. Imposible continuar.", vbExclamation, "Atención"
    Exit Sub
End If

If ioFINI.Text <> "" And ioFFIN.Text <> "" Then enFechas = True

cnd.BeginTrans
entrans = True

'------------------------------------------------------------------------------------------------
If comprueba_marcado("VENTAS") Then

    Call add_log("Comenzado Ventas. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
    
    If enFechas Then
        '(CAST(CODIGO AS VARCHAR(10)) + CAST(CODCAJA AS CHAR(3)))
        cnd.Execute "DELETE FROM DETVENTA WHERE (CAST(CODVEN AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABVENTA WHERE FHORA >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FHORA <= '" & Format(ioFFIN.Text, "yyyymmdd") & "' AND ESTADO < 2)"
        cnd.Execute "DELETE FROM CABVENTA WHERE FHORA >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FHORA <= '" & Format(ioFFIN.Text, "yyyymmdd") & "' AND ESTADO < 2"
        rco.Open "SELECT * FROM CABVENTA WHERE FHORA >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FHORA <= '" & Format(ioFFIN.Text, "yyyymmdd") & "' AND ESTADO < 2", cno, adOpenStatic, adLockReadOnly
    Else
        cnd.Execute "DELETE FROM DETVENTA WHERE CODCAJA = " & CajaActual
        cnd.Execute "DELETE FROM CABVENTA WHERE CODCAJA = " & CajaActual
        rco.Open "SELECT * FROM CABVENTA WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
    End If

    'CABVENTA
    DoEvents
    rcd.Open "CABVENTA", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
    '------------------------------------------------------------------------------------------------
    
    'DETVENTA
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
    
    If enFechas Then
        rco.Open "SELECT * FROM DETVENTA WHERE (CAST(CODVEN AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABVENTA WHERE FHORA >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FHORA <= '" & Format(ioFFIN.Text, "yyyymmdd") & "' AND ESTADO < 2)", cno, adOpenStatic, adLockReadOnly
    Else
        rco.Open "SELECT * FROM DETVENTA WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
    End If
    
    DoEvents
    rcd.Open "DETVENTA", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------
 
    Call add_log("Finalizado Ventas. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)

End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("VALES") Then

   Call add_log("Comenzado Vales. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
     
    'VALES
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close

    If enFechas Then
        rco.Open "SELECT * FROM VALES WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "') OR (FACEP >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FACEP <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM VALES WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "') OR (FACEP >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FACEP <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM VALES WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM VALES WHERE CODCAJA = " & CajaActual
    End If

    DoEvents
    rcd.Open "VALES", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------

    Call add_log("Finalizado Vales. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    

End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("DEVOLUCIONES") Then

    Call add_log("Comenzado Devoluciones. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    'DEVOL
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
    
    If enFechas Then
        rco.Open "SELECT * FROM DEVOL WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM DEVOL WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM DEVOL WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM DEVOL WHERE CODCAJA = " & CajaActual
    End If
    
    DoEvents
    rcd.Open "DEVOL", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------

    Call add_log("Finalizado Devoluciones. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    

End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("ARREGLOS") Then

    Call add_log("Comenzado Arreglos. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    'ARREGLOS
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
    
    If enFechas Then
        rco.Open "SELECT * FROM ARREGLOS WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM ARREGLOS WHERE (FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM ARREGLOS WHERE CODCAJ = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM ARREGLOS WHERE CODCAJ = " & CajaActual
    End If
    
    DoEvents
    rcd.Open "ARREGLOS", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True, True)
'------------------------------------------------------------------------------------------------
    Call add_log("Finalizado Arreglos. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    

End If


'------------------------------------------------------------------------------------------------
If comprueba_marcado("MOV. CAJA") Then
    'MOVCAJA
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close

    Call add_log("Comenzado MOV. CAJA. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    If enFechas Then
        rco.Open "SELECT * FROM MOVCAJA WHERE (FCIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FCIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM MOVCAJA WHERE (FCIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FCIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM MOVCAJA WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM MOVCAJA WHERE CODCAJA = " & CajaActual
    End If

    DoEvents
    rcd.Open "MOVCAJA", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------
    
    Call add_log("Finalizado MOV. CAJA. Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("CIERRE DIARIO") Then

    'CIERREDIA
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
    
    Call add_log("Comenzado CIERRE DIARIO Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    If enFechas Then
        rco.Open "SELECT * FROM CIERREDIA WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREDIA WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM CIERREDIA WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREDIA WHERE CODCAJA = " & CajaActual
    End If
    
    
    DoEvents
    rcd.Open "CIERREDIA", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------
    Call add_log("Finalizado CIERRE DIARIO Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    

End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("CIERRE MENSUAL") Then

    'CIERREMES
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close

    Call add_log("Comenzado CIERRE MENSUAL Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
    
    If enFechas Then
        rco.Open "SELECT * FROM CIERREMES WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREMES WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM CIERREMES WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREMES WHERE CODCAJA = " & CajaActual
    End If
    
    DoEvents
    rcd.Open "CIERREMES", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------
    Call add_log("Finalizado CIERRE MENSUAL Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)

End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("CIERRE ANUAL") Then

    'CIERREANO
    If rco.State = 1 Then rco.Close
    If rcd.State = 1 Then rcd.Close
        
    Call add_log("Comenzado CIERRE ANUAL Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)
      
    If enFechas Then
        rco.Open "SELECT * FROM CIERREANO WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREANO WHERE (FECIERRE >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    Else
        rco.Open "SELECT * FROM CIERREANO WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
        cnd.Execute "DELETE FROM CIERREANO WHERE CODCAJA = " & CajaActual
    End If
    
    DoEvents
    rcd.Open "CIERREANO", cnd, adOpenStatic, adLockOptimistic
    Call copia_rc(rco, rcd, CajaActual, True)
    '------------------------------------------------------------------------------------------------
    Call add_log("Finalizado CIERRE ANUAL Entre fechas(" & CStr(enFechas) & ") Intervalo: " & ioFINI.Text & "-" & ioFFIN.Text)


End If

'------------------------------------------------------------------------------------------------
If comprueba_marcado("PAGOS") Then
   
'PAGOS
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

If enFechas Then
    cnd.Execute "DELETE FROM DETPAGOS WHERE (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM PAGOS WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    cnd.Execute "DELETE FROM PAGOS WHERE CODCAJA = " & CajaActual & " AND FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    rco.Open "SELECT * FROM PAGOS WHERE CODCAJA = " & CajaActual & " AND FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
Else
    cnd.Execute "DELETE FROM DETPAGOS WHERE CODCAJA = " & CajaActual
    cnd.Execute "DELETE FROM PAGOS WHERE CODCAJA = " & CajaActual
    rco.Open "SELECT * FROM PAGOS WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
End If
    
DoEvents
rcd.Open "PAGOS", cnd, adOpenStatic, adLockOptimistic
Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------

'DETPAGOS
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

    If enFechas Then
        rco.Open "SELECT * FROM DETPAGOS WHERE (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM PAGOS WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
    Else
        rco.Open "SELECT * FROM DETPAGOS WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
    End If

DoEvents
rcd.Open "DETPAGOS", cnd, adOpenStatic, adLockOptimistic
Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------

End If


If comprueba_marcado("PRUEBAS MERC.") Then

'CABPRESTA
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

If enFechas Then
    cnd.Execute "DELETE FROM DETPRESTA WHERE CODCAJA = " & CajaActual & " AND (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABPRESTA WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    cnd.Execute "DELETE FROM CABPRESTA WHERE CODCAJA = " & CajaActual & " AND FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    rco.Open "SELECT * FROM CABPRESTA  WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
Else
    cnd.Execute "DELETE FROM DETPRESTA WHERE CODCAJA = " & CajaActual
    cnd.Execute "DELETE FROM CABPRESTA WHERE CODCAJA = " & CajaActual
    rco.Open "SELECT * FROM CABPRESTA  WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
End If

rcd.Open "CABPRESTA", cnd, adOpenStatic, adLockOptimistic

DoEvents
Call copia_rc(rco, rcd, CajaActual, True)

'DETPRESTA
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

If enFechas Then
    rco.Open "SELECT * FROM DETPRESTA WHERE CODCAJA = " & CajaActual & " AND (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABPRESTA WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
Else
    rco.Open "SELECT * FROM DETPRESTA wHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
End If

DoEvents
rcd.Open "DETPRESTA", cnd, adOpenStatic, adLockOptimistic
Call copia_rc(rco, rcd, CajaActual, True)
'------------------------------------------------------------------------------------------------
End If


'DEUDAS CLIENTES
'------------------------------------------------------------------------------------------------
If comprueba_marcado("DEUDAS CLIENT.") Then

'CABDEUDCLI
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

If enFechas Then
    cnd.Execute "DELETE FROM DETDEUDCLI WHERE CODCAJA = " & CajaActual & " AND (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABDEUDCLI WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    cnd.Execute "DELETE FROM CABDEUDCLI WHERE CODCAJA = " & CajaActual & " AND FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')"
    rco.Open "SELECT * FROM CABDEUDCLI  WHERE CODCAJA = " & CajaActual & " AND FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
Else
    cnd.Execute "DELETE FROM DETDEUDCLI WHERE CODCAJA = " & CajaActual
    cnd.Execute "DELETE FROM CABDEUDCLI WHERE CODCAJA = " & CajaActual
    rco.Open "SELECT * FROM CABDEUDCLI  WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
End If

DoEvents

rcd.Open "CABDEUDCLI", cnd, adOpenStatic, adLockOptimistic

Call copia_rc(rco, rcd, CajaActual, True)

'DETDEUDCLI
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

If enFechas Then
    rco.Open "SELECT * FROM DETDEUDCLI  WHERE CODCAJA = " & CajaActual & " AND (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) IN (SELECT (CAST(CODIGO AS CHAR(10)) + CAST(CODCAJA AS CHAR(3))) FROM CABDEUDCLI WHERE FMODI >= '" & Format(ioFINI.Text, "yyyymmdd") & "' AND FMODI <= '" & Format(ioFFIN.Text, "yyyymmdd") & "')", cno, adOpenStatic, adLockReadOnly
Else
    rco.Open "SELECT * FROM DETDEUDCLI  WHERE CODCAJA = " & CajaActual, cno, adOpenStatic, adLockReadOnly
End If

DoEvents
rcd.Open "DETDEUDCLI", cnd, adOpenStatic, adLockOptimistic

Call copia_rc(rco, rcd, CajaActual, True)

End If
'------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
If comprueba_marcado("STOCK") Then

'STOCK (siempre actualizar todo el stock)
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close
rco.Open "SELECT * FROM STOCK WHERE CODALM = " & AlmacenActual & " AND STOCK <> 0", cno, adOpenStatic, adLockReadOnly
cnd.Execute "DELETE FROM STOCK WHERE CODALM = " & AlmacenActual
DoEvents
rcd.Open "STOCK", cnd, adOpenStatic, adLockOptimistic
Call copia_rc(rco, rcd, AlmacenActual, False)
'------------------------------------------------------------------------------------------------

End If


If comprueba_marcado("TRANSFERENCIAS") Then

'PTRANS
If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close
rco.Open "SELECT * FROM PTRANS WHERE CODALMORIG = " & AlmacenActual, cno, adOpenStatic, adLockReadOnly
rcd.Open "PTRANS", cnd, adOpenStatic, adLockOptimistic
'Call copia_rc(rco, rcd, AlmacenActual, False, False, True)
'------------------------------------------------------------------------------------------------

End If

cnd.CommitTrans
entrans = False

If rco.State = 1 Then rco.Close
If rcd.State = 1 Then rcd.Close

lblStatus.Caption = "Se ha terminado el proceso correctamente"

   On Error GoTo 0
   Exit Sub

cmd_Click_Error:

    Call add_log("Error en el proceso: " & Err.Number & ". " & Err.Description)
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmd_Click de Formulario FrmSinc"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : copia_rc
' Fecha/Hora    : 05/08/2004 09:28
' Autor         : JCastillo
' Propósito     :   Copiar un recordset en otro
'---------------------------------------------------------------------------------------
'
Private Sub copia_rc(origen As ADODB.Recordset, destino As ADODB.Recordset, codigo As Byte, escaja As Boolean, Optional esarreglo As Boolean, Optional esptrans As Boolean)
Dim campo As String
Dim var As Integer
Dim tmpstr As String

   On Error GoTo copia_rc_Error

    lblStatus.Caption = "Procesando " & destino.Source & " ..."
    lblStatus.Refresh
    
    tmpstr = lblStatus.Caption
    
    Do Until origen.EOF
    
        destino.AddNew
    
        For var = 0 To origen.Fields.Count - 1
    
            campo = origen.Fields(var).Name
        
            If UCase(campo) <> "ROWGUID" Then
                destino.Fields(campo).Value = origen.Fields(var).Value
            End If
    
        Next var
    
        destino.Update
        DoEvents
        
        origen.MoveNext
        lblStatus.Caption = tmpstr & " (" & origen.AbsolutePosition & "/" & origen.RecordCount & ")"
    
    Loop
    
    lblStatus.Caption = "Se ha Procesado " & destino.Source & " correctamente."
    lblStatus.Refresh
    
    
   On Error GoTo 0
   Exit Sub

copia_rc_Error:

    lblStatus.Caption = "Error al procesar " & destino.Source
    lblStatus.Refresh
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento copia_rc de Formulario FrmSinc. En tabla " & destino.Source

End Sub



Private Sub Form_Load()
Dim rco As New ADODB.Recordset

'cargar el connection para la configuracion
On Error GoTo Form_Load_Error

If Dir(App.Path & "\" & logfile) = "" Then Call CreateDatabase
DoEvents

cliStr = strCnnMdb & App.Path & "\Config.pcg"

rco.Open "SELECT * FROM PUESTCNF", cliStr, adOpenStatic, adLockReadOnly

CajaActual = rco.Fields("CODCAJA")
AlmacenActual = rco.Fields("CODALM")

rco.Close
Set rco = Nothing

Call carga_tablas

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario FrmSinc"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If cnd.State = 1 Then
        If entrans Then
            cnd.RollbackTrans
        End If
        cnd.Close
    End If
    Set cnd = Nothing
    If cno.State = 1 Then cno.Close
    Set cno = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : ioDBCliente_Validate
' Fecha/Hora    : 05/08/2004 09:28
' Autor         : JCastillo
' Propósito     :   Validar datos del cliente y abrir una conexion
'---------------------------------------------------------------------------------------
'
Private Sub ioDBCliente_lostfocus()
Dim rc As New ADODB.Recordset
       
   On Error GoTo ioDBCliente_Validate_Error

    lblStatus.Caption = ""
    If (ioCliente.Text = "") Then
                lblStatus.Caption = "No se permite IP en blanco"
                ioCliente.SetFocus
    End If
    
    If (ioDBCliente.Text = "") Then
                lblStatus.Caption = "No se permite IP o DATABASE en blanco"
                ioDBCliente.SetFocus
    End If
        
    'cargar cadenas de conexion
    cliStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & ioDBCliente.Text & ";Data Source=" & ioCliente.Text
    If cno.State = 1 Then cno.Close
    cno.Open cliStr
    
    lblStatus.Caption = "Se ha realizado la conexion al cliente."
    
    'Call carga_temp(lstTempor)
    rc.Open "SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & CajaActual, cno, adOpenStatic, adLockOptimistic
    lblCaja.Caption = Trim(rc.Fields(0))
    lblCaja.Refresh
            
    rc.Close
    Set rc = Nothing
           
   On Error GoTo 0
   Exit Sub

ioDBCliente_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioDBCliente_Validate de Formulario FrmSinc"
      
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : ioDBServer_Validate
' Fecha/Hora    : 05/08/2004 09:28
' Autor         : JCastillo
' Propósito     :   Validar datos del servidor y abrir una conexion
'---------------------------------------------------------------------------------------
'
Private Sub ioDBServer_lostfocus()
       
   On Error GoTo ioDBServer_Validate_Error

    lblStatus.Caption = ""
    If (ioServer.Text = "") Then
                lblStatus.Caption = "No se permite IP SERVER en blanco"
                ioCliente.SetFocus
    End If
    
    If (ioDBServer.Text = "") Then
                lblStatus.Caption = "No se permite IP o DATABASE SERVER en blanco"
                ioDBCliente.SetFocus
    End If
        
    'cargar cadenas de conexion
    srvStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & ioDBServer.Text & ";Data Source=" & ioServer.Text
    If cnd.State = 1 Then cnd.Close
    cnd.Open srvStr
     
    lblStatus.Caption = "Se ha realizado la conexion al servidor."

   On Error GoTo 0
   Exit Sub

ioDBServer_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioDBServer_Validate de Formulario FrmSinc"
              
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : carga_temp
' Fecha/Hora    : 05/08/2004 09:28
' Autor         : JCastillo
' Propósito     :   Cargar las temporadas en el listbox
'---------------------------------------------------------------------------------------
Private Sub carga_temp(lst As ListBox)
Dim rc As New ADODB.Recordset

   On Error GoTo carga_temp_Error

        rc.Open "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC", cno, adOpenStatic, adLockReadOnly
        
        lst.Clear
        
        Do Until rc.EOF
        
            lst.AddItem (Format(rc.Fields("IDTEM"), "000") & " -" & rc.Fields("ABREVIA"))
        
            rc.MoveNext
        Loop
        
        rc.Close
        Set rc = Nothing
                
   On Error GoTo 0
   Exit Sub

carga_temp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_temp de Formulario FrmSinc"
End Sub


Private Sub carga_tablas()
Dim var As Long
   
   On Error GoTo carga_tablas_Error

    With lstTablas
        .Clear
        .AddItem "VENTAS"
        .AddItem "VALES"
        .AddItem "ARREGLOS"
        .AddItem "DEVOLUCIONES"
        .AddItem "MOV. CAJA"
        .AddItem "STOCK"
        .AddItem "PAGOS"
        .AddItem "PRUEBAS MERC."
        .AddItem "DEUDAS CLIENT."
        .AddItem "CIERRE DIARIO"
        .AddItem "CIERRE MENSUAL"
        .AddItem "CIERRE ANUAL"
        .AddItem "TRANSFERENCIAS"
        
        For var = 0 To .ListCount - 1
        .Selected(var) = True
        Next var
        
    End With
 
   On Error GoTo 0
   Exit Sub

carga_tablas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_tablas de Formulario FrmSinc"
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : comprueba_marcado
' Fecha/Hora    : 05/08/2004 11:01
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
Private Function comprueba_marcado(nombre As String) As Boolean
Dim var As Long

   On Error GoTo comprueba_marcado_Error
       
        For var = 0 To lstTablas.ListCount - 1
                
            If lstTablas.List(var) = nombre Then
             comprueba_marcado = lstTablas.Selected(var)
            End If
                
        Next var
    
   On Error GoTo 0
   Exit Function

comprueba_marcado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_marcado de Formulario FrmSinc"
End Function


Private Sub CreateDatabase()
Dim Cat     As New ADOX.Catalog
Dim Tbl(5) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim i As Long


   On Error GoTo CreateDatabase_Error

Cat.Create strCnnMdb & App.Path & "\" & logfile

  '----------* Table Definition of LOG *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "LOG"
    .Columns.Append "FECHA", adDate
      .Columns("FECHA").Properties("Nullable").Value = False
      .Columns("FECHA").Properties("Default").Value = "now()"
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    .Columns.Append "LOG", adVarWChar, 200
      .Columns("LOG").Properties("Nullable").Value = False
  End With
  '----------* Index Definitions of LOG *----------
  ReDim Idx(1)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Set Idx(1) = New ADOX.Index
    Idx(1).Name = "FECHA"
    Idx(1).IndexNulls = adIndexNullsAllow
      Idx(1).Columns.Append "FECHA"
  For i = 0 To UBound(Idx)
    Tbl(0).Indexes.Append Idx(i)
  Next i

  Cat.Tables.Append Tbl(0)

  Set Cat = Nothing
  Exit Sub

CreateDatabase_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CreateDatabase de Formulario FrmSinc"

   On Error GoTo 0
   Exit Sub


End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : add_log
' Fecha/Hora    : 05/08/2004 16:39
' Autor         : JCastillo
' Propósito     :  Añade un registro en la db de log
'---------------------------------------------------------------------------------------
'
Private Sub add_log(cadena As String)
Dim cn As New ADODB.Connection

   On Error GoTo add_log_Error

        cn.Open strCnnMdb & App.Path & "\" & logfile
        cn.Execute "INSERT INTO LOG (LOG) VALUES('" & cadena & "')"
        DoEvents
        cn.Close
        Set cn = Nothing

   On Error GoTo 0
   Exit Sub

add_log_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento add_log de Formulario FrmSinc"
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : conforma_log
' Fecha/Hora    : 05/08/2004 17:17
' Autor         : JCastillo
' Propósito     :   crea un string con los nombres de loq se va a actualizar
'---------------------------------------------------------------------------------------
'
Private Function conforma_log() As String
Dim var As Long
Dim tmpcad As String

       
   On Error GoTo conforma_log_Error

        For var = 0 To lstTablas.ListCount - 1
     
            If lstTablas.Selected(var) = True Then
             tmpcad = tmpcad & "-" & lstTablas.List(var)
            End If
                
        Next var
        
        conforma_log = tmpcad

   On Error GoTo 0
   Exit Function

conforma_log_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento conforma_log de Formulario FrmSinc"
    
End Function



'---------------------------------------------------------------------------------------
' Procedimiento : ioFINI_Validate
' Fecha/Hora    : 06/08/2004 09:51
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
Private Sub ioFINI_Validate(Cancel As Boolean)

   On Error GoTo ioFINI_Validate_Error

    If ioFINI.Text = "" Then Exit Sub
    
    If Not IsDate(ioFINI.Text) Then
        lblStatus.Caption = "Fecha inicio incorrecta: " & ioFINI.Text
        ioFINI.Text = ""
        Cancel = True
    Else
        ioFINI.Text = Format(ioFINI.Text, "dd/mm/yyyy")
    End If

   On Error GoTo 0
   Exit Sub

ioFINI_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioFINI_Validate de Formulario FrmSinc"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : ioFFIN_Validate
' Fecha/Hora    : 06/08/2004 09:51
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
Private Sub ioFFIN_Validate(Cancel As Boolean)

   On Error GoTo ioFFIN_Validate_Error

    If ioFFIN.Text = "" Then Exit Sub
    
    If Not IsDate(ioFFIN.Text) Then
        lblStatus.Caption = "Fecha fin incorrecta: " & ioFFIN.Text
        ioFFIN.Text = ""
        Cancel = True
    Else
        ioFFIN.Text = Format(ioFFIN.Text, "dd/mm/yyyy")
    End If

   On Error GoTo 0
   Exit Sub

ioFFIN_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioFFIN_Validate de Formulario FrmSinc"

End Sub
