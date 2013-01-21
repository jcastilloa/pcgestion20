Attribute VB_Name = "Comentario"
'---------------------------------------------------------------------------------------
' Module    : Comentario
' DateTime  : 26/10/2003 21:20
' Author    : Administrador
' Purpose   : Rutinas para gestionar los comentarios por medio de
'             BWord (proyecto Editor).
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : guardar_rtf_campo
' DateTime  : 26/10/2003 21:22
' Author    : Administrador
' Purpose   : Guarda el contenido de un fichero RTF en un campo
'             (ADODB.Field). Devuelve FALSE si no ha ocurrido ningun error.
'---------------------------------------------------------------------------------------
Public Function guardar_rtf_campo(ficheroRTF As String, campo As ADODB.Field) As Boolean
Dim linea As String
Dim numfile As Integer


On Error GoTo guardar_rtf_campo_Error

numfile = FreeFile

'abrir el fichero
Open ficheroRTF For Input As #numfile

'leer linea a linea
Do While Not EOF(numfile)

    Input #numfile, linea
    'asignar el contenido al campo
    campo.Value = campo.Value & linea

Loop

Close #numfile

On Error GoTo 0
Exit Function

guardar_rtf_campo_Error:
    
    guardar_rtf_campo = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure guardar_rtf_campo of Módulo Comentario", vbExclamation, titulo

End Function


'---------------------------------------------------------------------------------------
' Procedure : lee_RTF_campo
' DateTime  : 26/10/2003 21:35
' Author    : Administrador
' Purpose   : leer contenido de un campo y guardar en un fichero de texto (RTF)
'---------------------------------------------------------------------------------------
Public Function lee_RTF_campo(campo As ADODB.Field, ficheroRTF As String)
Dim linea As String
Dim var As Long
   
   On Error GoTo lee_RTF_campo_Error

   
   For var = 1 To Len(campo.Value)
   
   
   Next
   
   linea = ""
    
  
   On Error GoTo 0
   Exit Function

lee_RTF_campo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lee_RTF_campo of Módulo Comentario", vbExclamation, titulo
End Function


