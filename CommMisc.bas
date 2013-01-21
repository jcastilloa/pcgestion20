Attribute VB_Name = "CommMisc"
'---------------------------------------------------------------------------------------
' Módulo     : CommMisc
' Fecha/Hora : 22/04/2004 09:54
' Autor      : JCastillo
' Propósito  :  Rutinas varias para las comunicaciones
'---------------------------------------------------------------------------------------
Option Explicit

' Timer Functions
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer


Public Function OpenURL(ByVal sUrl As String) As String
'****************************************************
'PURPOSE:       Returns Contents (including all HTML) from
'               a web page
'PARAMETER:     sURL (e.g., http://www.freevbcode.com)
'RETURN VALUE:  Contents of requested page, or
'               empty string if sURL is not available
'COMMENTS:  This is an alternative to using the Internet Transfer
'           Control 's OpenURL method.  That control has a bug
'           Whereby not all the contents of the page will be
'           returned in certain circumstances
'*****************************************************

    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, _
    vbNullString, vbNullString, 0)

hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
   INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
           Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
             lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
      
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer

End Function


Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
Dim tmpurl As String
    
    ' This Sub is identical to the Timer
    ' event in a standard timer.
    
   Static ElapsedTime As Long
        
   On Error GoTo TimerProc_Error

 If strLocCnn <> "" Then
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
 End If
       
    tmpurl = OpenURL("http://www.showmyip.com/simple/")
    
    tmpurl = Left(tmpurl, 50)
        
    ElapsedTime = ElapsedTime + 1
    If Trim(tmpurl) = "" Then Exit Sub
    
    'si cambio la url, actualizar
    If Trim(devuelve_campo("SELECT HOSTDIR FROM CENTROS WHERE CODIGO = " & CentroActual, locCnn)) <> Trim(tmpurl) Then
    
        locCnn.Execute "UPDATE CENTROS SET HOSTDIR = '" & Trim(tmpurl) & "' WHERE CODIGO = " & CentroActual
    
    End If

   On Error GoTo 0
   Exit Sub

TimerProc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento TimerProc de Módulo CommMisc"
    
End Sub


Public Sub ajusta_pedidos(ultimo As Long)
Dim rcCab As adodb.Recordset
Dim tmpvar As Long

   'On Error GoTo ajusta_pedidos_Error

 If strLocCnn <> "" Then
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
 End If

Set rcCab = New adodb.Recordset

tmpvar = ultimo + 1

rcCab.Open "SELECT NUMERO, ALMORIG FROM CABPEDPRO ORDER BY ALMORIG, NUMERO", locCnn, adOpenDynamic, adLockOptimistic

Do Until rcCab.EOF

    'actualizar registros del detalle
    locCnn.Execute "UPDATE DETPEDPRO SET NUMERO = " & tmpvar & " WHERE NUMERO = " & rcCab.fields("NUMERO") & " AND ALMORIG = " & rcCab.fields("ALMORIG")

    rcCab.fields("NUMERO") = tmpvar
    rcCab.Update
    rcCab.MoveNext
    tmpvar = tmpvar + 1

Loop

rcCab.Close
Set rcCab = Nothing

   On Error GoTo 0
   Exit Sub

ajusta_pedidos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ajusta_pedidos de Módulo CommMisc"

End Sub
