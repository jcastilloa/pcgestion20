Attribute VB_Name = "ImagenesADO"
'------------------------------------------------------------------------------
' C�digo para grabar y leer imagenes en campos de bases             ( 9/Abr/98)
' Adaptado para usarlo con ADO                                      (11/Jul/01)
'
' Adaptado de un par de ejemplos de la ayuda de VB5
'
' �Guillermo 'guille' Som, 1998-2001
' guille@costasol.net
'------------------------------------------------------------------------------
Option Explicit

Private nFile As Long
Private Chunk() As Byte
Private Const mBuffer As Long = 16384&

Public Sub LeerBinary(ADOField As ADODB.Field, unPicture As PictureBox)
    ' Leer la imagen del campo de la base y asignarlo al Picture
    '--------------------------------------------
    ' Este procedimiento no es necesario usarlo
    ' si el Picture est� ligado a un data control
    '--------------------------------------------
    Dim nChunks As Long
    Dim nSize As Long
    Dim Fragment As Long
    Dim i As Long
    '
    ' Se usa un fichero temporal para guardar la imagen
    nFile = FreeFile
    Open "pictemp" For Binary Access Write As nFile
    '
    ' Calcular los trozos completos y el resto
    nSize = ADOField.ActualSize
    nChunks = Int(nSize / mBuffer)
    Fragment = nSize Mod mBuffer
    Chunk() = ADOField.GetChunk(Fragment)
    Put nFile, , Chunk()
    For i = 1 To nChunks
        Chunk() = ADOField.GetChunk(mBuffer)
        Put nFile, , Chunk()
    Next
    Close nFile
    Erase Chunk
    ' Ahora se carga esa imagen en el control
    unPicture.Picture = LoadPicture("pictemp")
    
    ' Ya no necesitamos el fichero, as� que borrarlo
    On Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub

Public Sub GuardarBinary(ADOField As ADODB.Field, unPicture As PictureBox)
    ' Guardar el contenido del Picture en el campo de la base
    Dim i As Long
    Dim Fragment As Long
    Dim nSize As Long
    Dim nChunks As Long
    '
    ' Guardar el contenido del picture en un fichero temporal
    SavePicture unPicture.Picture, "pictemp"
    
    ' Leer el fichero y guardarlo en el campo
    nFile = FreeFile
    Open "pictemp" For Binary Access Read As nFile
    nSize = LOF(nFile)    ' Longitud de los datos en el archivo
    If nSize = 0 Then
        Close nFile
        Exit Sub
    End If
    '
    ' Calcular el n�mero de trozos y el resto
    nChunks = nSize \ mBuffer
    Fragment = nSize Mod mBuffer
    ReDim Chunk(Fragment)
    '
    Get nFile, , Chunk()
    ADOField.AppendChunk Chunk()
    ReDim Chunk(mBuffer)
    For i = 1 To nChunks
        Get nFile, , Chunk()
        ADOField.AppendChunk Chunk()
    Next i
    Close nFile
    '
    ' Ya no necesitamos el fichero, as� que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub


Public Sub GuardarArchivo(ADOField As ADODB.Field, fichero As String)
    ' Guardar el contenido del Picture en el campo de la base
    Dim i As Long
    Dim Fragment As Long
    Dim nSize As Long
    Dim nChunks As Long
    '
    ' Guardar el contenido del picture en un fichero temporal
    'SavePicture unPicture.Picture, "pictemp"
    
    ' Leer el fichero y guardarlo en el campo
    nFile = FreeFile
    Open fichero For Binary Access Read As nFile
    nSize = LOF(nFile)    ' Longitud de los datos en el archivo
    If nSize = 0 Then
        Close nFile
        Exit Sub
    End If
    '
    ' Calcular el n�mero de trozos y el resto
    nChunks = nSize \ mBuffer
    Fragment = nSize Mod mBuffer
    ReDim Chunk(Fragment)
    '
    Get nFile, , Chunk()
    ADOField.AppendChunk Chunk()
    ReDim Chunk(mBuffer)
    For i = 1 To nChunks
        Get nFile, , Chunk()
        ADOField.AppendChunk Chunk()
    Next i
    Close nFile
    '
    ' Ya no necesitamos el fichero, as� que borrarlo
  '  On Local Error Resume Next
  '  If Len(Dir$("pictemp")) Then
  '      Kill "pictemp"
  '  End If
  '  Err = 0
End Sub


