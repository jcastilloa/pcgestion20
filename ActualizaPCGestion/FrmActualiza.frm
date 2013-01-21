VERSION 5.00
Begin VB.Form FrmActualiza 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmActualiza.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListView1 
      BackColor       =   &H00E0E0E0&
      Height          =   2160
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   9330
   End
   Begin VB.CommandButton cbCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4717
      TabIndex        =   1
      Top             =   2565
      Width           =   1575
   End
   Begin VB.CommandButton cbActualizar 
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3112
      TabIndex        =   0
      Top             =   2565
      Width           =   1575
   End
   Begin VB.Label LblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   900
      Left            =   15
      TabIndex        =   2
      Top             =   2190
      Width           =   9345
   End
End
Attribute VB_Name = "FrmActualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const fichero = "PCGestion20.exe"
Const ficheroFTP = "PCGestion20.z"

Const fconfig = "update.ini"

Const fversion = "version.ini"    'fichero local de versión
Const fsversion = "sversion.ini"  'fichero remoto de versión

Const direc = "\Archivos de Programa"
Const titulo = "Actualización de Software"

Dim por_ftp As Boolean

Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1

'estructura del fichero update.INI
'Linea 1:   Ip del Server
'Linea 2:   User
'Linea 3:   Password

'estructura del fichero version.INI
'Linea 1:   versión actual del programa
'              ejemplo:   2.0.1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Enum CZErrors
[Insufficient Buffer] = -5
End Enum

'Custom compressed file header
Private Type CompressionHeader
    OriginalExt As String * 3
    OriginalSize As Long
End Type
'Actual header variable
Dim FileHeader As CompressionHeader
'Used to compare compression ratios
Dim OriginalSize As Long, CompressedSize As Long

Public Sub CompressFile(ByVal SrcFilename As String, DstFilename As String)
'Used to strip the extension from the original filename
Dim fExtension As String * 3
   On Error GoTo CompressFile_Error

fExtension = Right(SrcFilename, 3)

'Allocate an array to receive the data from a file
Dim DataBytes() As Byte
ReDim DataBytes(FileLen(SrcFilename) - 1)

'Copy the data from the source into a numerical array
Open SrcFilename For Binary Access Read As #1
    Get #1, , DataBytes()
Close #1

'Track the original size
OriginalSize = UBound(DataBytes) + 1

'Allocate memory for a temporary compression array
Dim BUFFERSIZE As Long
Dim TempBuffer() As Byte
BUFFERSIZE = UBound(DataBytes) + 1
BUFFERSIZE = BUFFERSIZE + (BUFFERSIZE * 0.01) + 12
ReDim TempBuffer(BUFFERSIZE)

'Compress the data using zLib
Dim result As Long
result = compress(TempBuffer(0), BUFFERSIZE, DataBytes(0), UBound(DataBytes) + 1)

'Copy the compressed data back into our first array
ReDim DataBytes(BUFFERSIZE - 1)
CopyMemory DataBytes(0), TempBuffer(0), BUFFERSIZE

'Kill the now useless buffer
Erase TempBuffer

'Some very simple error handling
If result = 0 Then
    CompressedSize = UBound(DataBytes) + 1
Else
    Err.Raise 1, "CompressFile", "Se produjo un error al comprimir " & SrcFilename
    Exit Sub
End If

On Error Resume Next
Kill DstFilename

'Build our custom compressed file header
FileHeader.OriginalExt = fExtension
FileHeader.OriginalSize = OriginalSize
'Write the header and then the compressed data
Open DstFilename For Binary Access Write As #1
    Put #1, 1, FileHeader
    Put #1, , DataBytes()
Close #1

'Kill the now unnecessary compressed data
Erase DataBytes

   On Error GoTo 0
   Exit Sub

CompressFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CompressFile de Módulo modExport"

End Sub

Public Sub DecompressFile(ByVal SrcFilename As String, Optional ByRef DstFilename As String = 0)
'Allocate a temporary array for receiving the compressed data
Dim DataBytes() As Byte
Dim result As Long
   On Error GoTo DecompressFile_Error

ReDim DataBytes(FileLen(SrcFilename) - Len(FileHeader) - 1)
'Copy out the header and then the compressed data
Open SrcFilename For Binary Access Read As #1
    Get #1, 1, FileHeader
    Get #1, , DataBytes()
Close #1

'Get the compressed size
OriginalSize = UBound(DataBytes) + 1

'Allocate memory for buffers
Dim BUFFERSIZE As Long
Dim TempBuffer() As Byte
BUFFERSIZE = FileHeader.OriginalSize
BUFFERSIZE = BUFFERSIZE + (BUFFERSIZE * 0.01) + 12
ReDim TempBuffer(BUFFERSIZE)

'Decompress the data using zLib
result = uncompress(TempBuffer(0), BUFFERSIZE, DataBytes(0), UBound(DataBytes) + 1)

'Copy the uncompressed data back into our first array
ReDim DataBytes(BUFFERSIZE - 1)
CopyMemory DataBytes(0), TempBuffer(0), BUFFERSIZE

'Some very simple error handling
If result = 0 Then
    CompressedSize = UBound(DataBytes) + 1
Else
    Err.Raise 2, "DeCompressFile", "Se produjo un error al descomprimir " & SrcFilename
    Exit Sub
End If

'Kill the now unnecessary buffer
Erase TempBuffer

'Build the output path using the original filename
If Len(DstFilename) = 0 Then
    DstFilename = Left(SrcFilename, Len(SrcFilename) - 3)
    DstFilename = DstFilename & FileHeader.OriginalExt
End If
On Error Resume Next
Kill DstFilename

'Write the uncompressed data back into its original format
Open DstFilename For Binary Access Write As #1
    Put #1, , DataBytes()
Close #1

'Kill the now unnecessary data array
Erase DataBytes

   On Error GoTo 0
   Exit Sub

DecompressFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento DecompressFile de Módulo modExport"

End Sub


Public Function Registry(KeyName As String, EXEName As String)
 Dim HREG As Long
 Dim StringBuffer  As String
 RegOpenKeyEx &H80000002, "Software\Microsoft\Windows\CurrentVersion\Run", 0, &H20006, HREG
 StringBuffer = EXEName & vbNullChar
 RegSetValueEx HREG, KeyName, 0, 1, ByVal StringBuffer, Len(StringBuffer)
 RegCloseKey HREG
End Function


Private Sub cbActualizar_Click()
Dim var As Long

    On Error GoTo cbActualizar_Click_Error

    LblStatus.Caption = "Buscando fichero " & fichero & " ..."
    LblStatus.Refresh

    ListView1.Clear

    'buscar el fichero en C
    Call FilesSearch("c:" & direc, fichero)

    'buscar el fichero en D
    Call FilesSearch("d:" & direc, fichero)

    'buscar el fichero en E
    Call FilesSearch("e:" & direc, fichero)

    LblStatus.Caption = "Búsqueda finalizada."
    LblStatus.Refresh

    'si no se encuentra ninguno, salir
    If ListView1.ListCount = 0 Then
        MsgBox "No se ha encontrado el programa en este ordenador, imposible actualizar", vbCritical, titulo
        Exit Sub
    End If

    MsgBox "Se ha encontrado (" & ListView1.ListCount & ") fichero/s. ¿Proceder con la actualización?", vbQuestion + vbYesNo, titulo

    For var = 0 To ListView1.ListCount - 1

        'quitar atributos q pudiera tener
        SetAttr ListView1.List(var), vbNormal
   
        'borrar el fichero original
        Kill ListView1.List(var)
   
        Call FileCopy(App.Path & "\" & fichero, ListView1.List(var))
   
        'quitar atributos q pudiera tener
        SetAttr ListView1.List(var), vbNormal
   
    Next var

    LblStatus.Caption = "Se ha actualizado el programa correctamente"
    LblStatus.Refresh

    MsgBox "Se ha actualizado el programa correctamente", vbInformation, titulo
    
DoEvents

Unload Me

   On Error GoTo 0
   Exit Sub

cbActualizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbActualizar_Click de Formulario FrmActualiza"

End Sub


Function sAttr(Attr As VbFileAttribute) As String
  Dim sStr1 As String
  sStr1 = ""
  If Attr And vbReadOnly Then sStr1 = "r" Else sStr1 = "-"
  If Attr And vbArchive Then sStr1 = sStr1 + "a" Else sStr1 = sStr1 + "-"
  If Attr And vbHidden Then sStr1 = sStr1 + "h" Else sStr1 = sStr1 + "-"
  If Attr And vbSystem Then sStr1 = sStr1 + "s" Else sStr1 = sStr1 + "-"
  sAttr = sStr1
End Function

Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()

'si no se encuentra el update del programa en el mismo directorio, salir.
On Error GoTo Form_Load_Error

If FileExists(App.Path & "\" & fconfig) Then
    
   Call Registry("Actualizar PC Gestion", App.Path & "\Actualiza.exe")
   'si exist el fichero de configuración de FTP en el directorio actual, se asume que
   'se trata de la actualización del ejecutable por FTP.
   por_ftp = True
   cbActualizar.Visible = False
   cbCancelar.Visible = False
   
   ListView1.Visible = False
   LblStatus.Top = 5
   LblStatus.BorderStyle = 1
   
   LblStatus.Height = Me.Height - 10
   
   Show
   Me.Refresh
   
   Call Lee_Configuracion_Ftp
   
   Exit Sub
   
Else
   por_ftp = False
End If


If Not FileExists(App.Path & "\" & fichero) Then
    MsgBox "No se ha encontrado el fichero " & fichero & " (update del programa) en el directorio actual (" & App.Path & ")." & Chr(13) & "Imposible continuar.", vbExclamation, titulo
    End
End If

Me.Caption = titulo

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario FrmActualiza"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Lee_Configuracion_Ftp
' Fecha/Hora     : 09/06/2004 10:39
' Autor             : JCastillo
' Propósito       :  Lee la configuración del fichero y conecta al servidor.
'---------------------------------------------------------------------------------------
Private Sub Lee_Configuracion_Ftp()
Dim server As String
Dim user As String
Dim pwd As String
Dim Version As String
Dim sversion As String
Dim l As Variant
Const segundos = 10

Dim numseg As Byte

On Error GoTo Lee_Configuracion_Ftp_Error
   
    l = FileLen(App.Path & "\" & fconfig)
    numseg = segundos

   Open App.Path & "\" & fconfig For Input As #1
    Input #1, server
    Input #1, user
    Input #1, pwd
   Close #1
   
   Me.Caption = "Comprobando versión del programa en el servidor"
   
   Set mFTP = New cFTP
 
   With mFTP
   
   If .OpenConnection(server, user, pwd) Then
      .SetFTPDirectory "/"
      LblStatus.Caption = "Conectado al servidor central: " & server & " a las: " & Now
   Else
   
      Me.Caption = "Imposible conectar al servidor: " & server & " a las: " & Now
      
      LblStatus.ForeColor = vbRed
      LblStatus.FontBold = True
      
      For numseg = 1 To segundos
            Espera 1
            Beep
            LblStatus.Caption = "Imposible conectar al servidor: " & server & Chr(13) & .GetLastErrorMessage & Chr(13) & "Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
      Next numseg
      
      'ejecutar el programa ...
      If Dir(App.Path & "\" & fichero) <> "" Then Call Shell(App.Path & "\" & fichero)
      Unload Me
      Exit Sub
      
   End If

   .SetModeActive
   'ascii para bajar el fichero fsversion
   .SetTransferASCII
      
   'bajar fichero de versión
   If Not .FTPDownloadFile(App.Path & "\" & fsversion, "\ActPGM\" & fsversion) Then
           Me.Caption = "Error al actualizar datos (compruebe la conexion de red y/o internet)"
           .CloseConnection
           
           LblStatus.ForeColor = vbRed
           LblStatus.FontBold = True
                    
           For numseg = 1 To segundos
               Espera 1
               Beep
               LblStatus.Caption = "No se ha podido completar: " & .GetLastErrorMessage & Chr(13) & "Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
           Next numseg
         
         
         'ejecutar el programa ...
         If Dir(App.Path & "\" & fichero) <> "" Then Call Shell(App.Path & "\" & fichero)
         Unload Me
         Exit Sub
         
   Else
           
           LblStatus.ForeColor = vbBlue
           LblStatus.FontBold = True
           LblStatus.Caption = fsversion & " (comprobación de versión) descargado."
           
   End If
      
   'quitar atributos q pudiera tener
   SetAttr App.Path & "\" & fsversion, vbNormal
   
   'comprobar version ...
   'remota
   If FileLen(App.Path & "\" & fsversion) > 0 Then
        Open App.Path & "\" & fsversion For Input As #1
        Input #1, sversion
        Close #1
   End If
   
   If Dir(App.Path & "\" & fversion) = "" Then
    Version = ""
   Else
   
   'local
   If FileLen(App.Path & "\" & fversion) > 0 Then
        Open App.Path & "\" & fversion For Input As #1
        Input #1, Version
        Close #1
   End If
   
   End If
   
   'si son distintas las versiones actualizar el EXE...
   If (Trim(sversion) <> Trim(Version)) Or (fversion = "") Then
   
   LblStatus.Caption = LblStatus.Caption & Chr(13) & "Se ha encontrado una versión en el servidor distinta a la actual. Se va a proceder a la actualización, espere por favor ..."
   
   .SetTransferBinary
   
     If Not .FTPDownloadFile(App.Path & "\" & ficheroFTP, "\ActPGM\" & ficheroFTP) Then
          
          Me.Caption = "Error al actualizar datos (compruebe la conexion de red y/o internet)"
          MsgBox "No se ha podido completar: " & .GetLastErrorMessage
          .CloseConnection
          
          LblStatus.ForeColor = vbRed
          LblStatus.FontBold = True
          
         For numseg = 1 To segundos
                Espera 1
                Beep
                LblStatus.Caption = "No se ha podido completar: " & .GetLastErrorMessage & Chr(13) & "Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
         Next numseg
         
         
         'ejecutar el programa ...
         If Dir(App.Path & "\" & fichero) <> "" Then Call Shell(App.Path & "\" & fichero)
          
         Unload Me
         Exit Sub
          
      Else
      
          LblStatus.Caption = "Nueva versión del programa descargada (" & sversion & ")"
          LblStatus.Refresh
           
          If Dir(App.Path & "\" & fichero) <> "" Then
            SetAttr App.Path & "\" & fichero, vbNormal
            Kill App.Path & "\" & fichero
          End If
                 
          'descomprimiendo fichero ...
          Call DecompressFile(App.Path & "\" & ficheroFTP, App.Path & "\" & fichero)
                                    
          If Dir(App.Path & "\" & fsversion) <> "" Then
            SetAttr App.Path & "\" & fsversion, vbNormal
            Kill App.Path & "\" & fsversion
          End If
          
          If Dir(App.Path & "\" & ficheroFTP) <> "" Then
            SetAttr App.Path & "\" & ficheroFTP, vbNormal
            Kill App.Path & "\" & ficheroFTP
          End If
                    
          
          '--------------- copia OK, escribir nueva versión en fichero local ------------------
          If Dir(App.Path & "\" & fversion) <> "" Then
            SetAttr App.Path & "\" & fversion, vbNormal
            Kill App.Path & "\" & fversion
          End If
             'local
          Open App.Path & "\" & fversion For Output As #1
          Print #1, Trim(sversion)
          Close #1
          '-------------------------------------------------------------------------------------------
          
          LblStatus.ForeColor = vbBlue
          LblStatus.FontBold = True
         
          For numseg = 1 To segundos
                Espera 1
                LblStatus.Caption = "Nueva versión del programa descargada (" & sversion & ")." & Chr(13) & "Descarga OK. Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
          Next numseg
        
          'ejecutar el programa ...
          If Dir(App.Path & "\" & fichero) <> "" Then Call Shell(App.Path & "\" & fichero)
          
          Unload Me
          DoEvents
          
          'actualizar fichero del programa
          'Call cbActualizar_Click
           
     End If
     
     Else
     
           For numseg = 1 To segundos
               Espera 1
               Beep
               LblStatus.Caption = "Ya esta instalada la última versión" & Chr(13) & "Esta ventana se cerrara dentro de " & segundos - numseg & " segundos."
           Next numseg
           
          'ejecutar el programa ...
          If Dir(App.Path & "\" & fichero) <> "" Then Call Shell(App.Path & "\" & fichero)
          Unload Me
          Exit Sub
   
   End If

    .CloseConnection
    
   'local
   Open App.Path & "\" & fversion For Input As #1
   Input #1, Version
   Close #1
   
    
   End With
   

   On Error GoTo 0
   Exit Sub

Lee_Configuracion_Ftp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Lee_Configuracion_Ftp de Formulario FrmActualiza"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End

End Sub



Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
   ' pb.Max = lTotalBytes
   ' pb.Min = 0
    If lTotalBytes > lCurrentBytes Then
    
        Me.Caption = "Descargados " & Round(lCurrentBytes / 1024, 1) & " KB. de " & Round(lTotalBytes / 1024, 1) & " KB."
       ' pb.Value = lCurrentBytes
    ElseIf lTotalBytes = lCurrentBytes Then
    
        Me.Caption = "Descarga completa."
        
    End If
    DoEvents
End Sub

Private Function FileExists(Filename As String) As Boolean
    FileExists = (Dir(Filename, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function


Private Sub Espera(nSeg As Single)
   Dim nIni As Single
   Dim nFin As Single
   nIni = Timer
   nFin = nIni + nSeg
   Do While nFin > Timer
      DoEvents
   Loop
End Sub
