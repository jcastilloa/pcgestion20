VERSION 5.00
Begin VB.Form frmBackupSQL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia de Seguridad"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
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
   ScaleHeight     =   5580
   ScaleWidth      =   7665
   Begin VB.DriveListBox Drive1 
      Height          =   420
      Left            =   4215
      TabIndex        =   7
      Top             =   405
      Width           =   3405
   End
   Begin VB.DirListBox Dir1 
      Height          =   3060
      Left            =   4200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   870
      Width           =   3435
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usar Compresión"
      Height          =   510
      Left            =   45
      TabIndex        =   5
      Top             =   1290
      Value           =   1  'Checked
      Width           =   3150
   End
   Begin VB.OptionButton optaSQL 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Autentificación SQLServer"
      Height          =   600
      Left            =   45
      TabIndex        =   3
      Top             =   660
      Width           =   3165
   End
   Begin VB.OptionButton optaWin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Autentificación de Windows"
      Height          =   600
      Left            =   45
      TabIndex        =   2
      Top             =   15
      Value           =   -1  'True
      Width           =   3165
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   4687
      TabIndex        =   0
      Top             =   4770
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
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
      MICON           =   "frmBackupSQL.frx":0000
      PICN            =   "frmBackupSQL.frx":001C
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
      Left            =   5647
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4770
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
      MICON           =   "frmBackupSQL.frx":0CF6
      PICN            =   "frmBackupSQL.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCrearCopia 
      Height          =   795
      Left            =   2752
      TabIndex        =   4
      Top             =   4770
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "C&opia de Seguridad"
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
      MICON           =   "frmBackupSQL.frx":15EC
      PICN            =   "frmBackupSQL.frx":1608
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   375
      Left            =   15
      Top             =   4350
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cmEstablecer 
      Height          =   345
      Left            =   4185
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3945
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "&Establecer"
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
      MICON           =   "frmBackupSQL.frx":22E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbDestinos 
      Height          =   795
      Left            =   1087
      TabIndex        =   10
      Top             =   4770
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "Trabajar con Destinos de la copia"
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
      MICON           =   "frmBackupSQL.frx":22FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label ioFMODI 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      Caption         =   "Guardar en"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4200
      TabIndex        =   8
      Top             =   30
      Width           =   3435
   End
End
Attribute VB_Name = "frmBackupSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmBackupSQL
' Fecha/Hora : 23/03/2004 12:03
' Autor         : JCastillo
' Propósito   :  Genera una copia de seguridad de la base de datos
'---------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------------------------------

Dim nombre_fichero As String '"c:\localpru"
Dim dir_destino As String

Private Sub cbAceptar_Click()

Unload Me

End Sub

Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub cbCrearCopia_Click()
Dim salida As Boolean
Dim nombref As String

   On Error GoTo cbCrearCopia_Click_Error
   
   
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With

lblStatus.Caption = "Realizando copia de seguridad ..."
DoEvents
   
salida = DB_Backup("(local)", "LOCAL", "sa", "admin", nombre_fichero, "Backup", "PC Gestión [" & Format(Date, "dd/mm/yyyy") & "]. Copia realizada por " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)))

'si no tiene la barra ponersela
If Right(dir_destino, 1) <> "\" Then dir_destino = dir_destino & "/"

nombref = dir_destino & "PCGBackup" & Format(Date, "dd-mm-yy") & ".baz"
Call CompressFile(nombre_fichero, nombref)

DoEvents
Kill nombre_fichero

'copia OK
If salida = True Then
    lblStatus.Caption = "La copia se ha realizado correctamente (" & nombref & ")"
'error
Else
    lblStatus.Caption = "No se ha podido realizar la copia"
End If

nombref = ""


   On Error GoTo 0
   Exit Sub

cbCrearCopia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCrearCopia_Click de Formulario frmBackupSQL"

End Sub


Private Sub cbDestinos_Click()
frmMntBack.Show
End Sub

Private Sub Dir1_Change()

   On Error GoTo Dir1_Change_Error

    dir_destino = Dir1.Path

   On Error GoTo 0
   Exit Sub

Dir1_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Dir1_Change de Formulario frmBackupSQL"
End Sub

Private Sub Drive1_Change()

   On Error GoTo Drive1_Change_Error

   Dir1.Path = Drive1.Drive

   On Error GoTo 0
   Exit Sub

Drive1_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Drive1_Change de Formulario frmBackupSQL"

End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) \ 2, Separacion_MDIForm
    
    'crear la base de datos (si no existe)
    Call CreateDB_BACKCNF
        
    dir_destino = "C:\"
    Dir1.Path = dir_destino
    Drive1.Drive = dir_destino
    nombre_fichero = GetTempFileName
    
End Sub


Private Sub CreateDB_BACKCNF()

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Dim i As Long

   On Error GoTo CreateDB_BACKCNF_Error
   
Const dbname = "\Backcnf.pcg"

If Dir(App.Path & dbname) <> "" Then Exit Sub

sCnn = strCnnMdb & App.Path & dbname

Cat.Create sCnn

  '----------* Table Definition of CONFIG *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "CONFIG"
    .Columns.Append "ANO", adSmallInt
      .Columns("ANO").Properties("Default").Value = "0"
    .Columns.Append "DEFECTO", adBoolean
      .Columns("DEFECTO").Properties("Default").Value = "False"
    .Columns.Append "DESCRIPCION", adVarWChar, 200
    .Columns.Append "ID", adInteger
      .Columns("ID").Properties("AutoIncrement").Value = True
      .Columns("ID").Properties("Nullable").Value = False
    .Columns.Append "MES", adUnsignedTinyInt
      .Columns("MES").Properties("Default").Value = "0"
    .Columns.Append "NUMCOPIAS", adInteger
      .Columns("NUMCOPIAS").Properties("Default").Value = "0"
    .Columns.Append "RUTA", adVarWChar, 200
    .Columns.Append "TEMPOR", adUnsignedTinyInt
      .Columns("TEMPOR").Properties("Default").Value = "0"
    .Columns.Append "TOTALCOPIAS", adInteger
      .Columns("TOTALCOPIAS").Properties("Default").Value = "0"
    .Columns.Append "ULFECHA", adDate
      .Columns("ULFECHA").Properties("Default").Value = "Date()"
  End With
  '----------* Index Definitions of CONFIG *----------
  ReDim Idx(1)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "ID"
  Set Idx(1) = New ADOX.Index
    Idx(1).Name = "ID"
    Idx(1).IndexNulls = adIndexNullsAllow
      Idx(1).Columns.Append "ID"
  For i = 0 To UBound(Idx)
    Tbl(0).Indexes.Append Idx(i)
  Next i

  Cat.tables.Append Tbl(0)

  Set Cat = Nothing
  frmMntBack.Show
  
  Exit Sub
  
  
  


   On Error GoTo 0
   Exit Sub

CreateDB_BACKCNF_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CreateDB_BACKCNF de Módulo BackupSQLS"

End Sub


