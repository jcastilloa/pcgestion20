Attribute VB_Name = "BackupSQLS"
Option Explicit


'* * * * * * * * * * * * * *
'Passing Values
'* * * * * * * * * * * * * *
'nServer_Name = SQL server name
'nDB_Name = Database name
'nDB_Login = Login name
'nDB_Password = Password
'nBack_Dev =Backup device name
'nBack_Set = Backup set name
'nBack_Desc = Backup discription

'Backup device name has to be specified by the SQL ADMIN.
'Which comes under SQL Backup. The name you specified must
'be same as Passing value of backup device name.

'SQL ADMIN can only specify the device type(Tape, HD,...).

'* * * * * * * * * * * * * *

'* * * * * * * * * * * * * *


Private Declare Function GetTempPathA Lib "kernel32" _
   (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
    
Private Declare Function GetTempFileNameA Lib "kernel32" _
    (ByVal lpszPath As String, ByVal lpPrefixString As String, _
     ByVal wUnique As Long, ByVal lpTempFileName As String) _
     As Long
     
Private Const UNIQUE_NAME = &H0

Private oSQLServer As SQLDMO.SQLServer

Public Function DB_Backup(ByVal nServer_Name As String, _
   ByVal nDB_Name As String, _
   ByVal nDB_Login As String, ByVal nDB_Password As String, _
   ByVal nBack_Dev As String, ByVal nBack_Set As String, _
   ByVal nBack_Desc As String) As Boolean

' nServer_Name = SQL server name
' nDB_Name = Database name
' nDB_Login = Login name
' nDB_Password = Password
' nBack_Dev =Backup device name
' nBack_Set = Backup set name
' nBack_Desc = Backup discription
 Dim oBackup As SQLDMO.Backup
  On Error GoTo ErrorHandler
  Set oBackup = CreateObject("SQLDMO.Backup")
  
  If Connect_SQLDB(nServer_Name, nDB_Login, nDB_Password) Then
    'oBackup.Devices = "[" & nBack_Dev & "]"
    oBackup.Files = nBack_Dev
    oBackup.Action = SQLDMOBackup_Database
    oBackup.Database = nDB_Name
    oBackup.BackupSetName = nBack_Set
    oBackup.BackupSetDescription = nBack_Desc
    oBackup.SQLBackup oSQLServer
        
    DoEvents
    
    oSQLServer.Disconnect
    DB_Backup = True
  End If
  
  Exit Function
ErrorHandler:
  DB_Backup = False
End Function


Private Function Connect_SQLDB(ByVal nServer_Name As String, _
    ByVal nDB_Login As String, _
    ByVal nDB_Password As String) As Boolean
  
  ' nServer_Name = SQL server name
  ' nDB_Login = Login name
  ' nDB_Password = Password

  Set oSQLServer = CreateObject("SQLDMO.SQLServer")
  On Error GoTo ErrorHandler
  Connect_SQLDB = False
  oSQLServer.LoginSecure = True
  oSQLServer.Connect nServer_Name, nDB_Login, nDB_Password
  Connect_SQLDB = True
  Exit Function
ErrorHandler:
  oSQLServer.Disconnect
  Connect_SQLDB = False
End Function




Public Function GetTempFileName() As String

   Dim sTmp    As String
   Dim sTmp2   As String

   sTmp2 = GetTempPath
   sTmp = Space(Len(sTmp2) + 256)
   Call GetTempFileNameA(sTmp2, App.EXEName, UNIQUE_NAME, sTmp)
   GetTempFileName = Left$(sTmp, InStr(sTmp, Chr$(0)) - 1)

End Function
Private Function GetTempPath() As String
  
   Dim sTmp       As String
   Dim i          As Integer

   i = GetTempPathA(0, "")
   sTmp = Space(i)

   Call GetTempPathA(i, sTmp)
   GetTempPath = AddBackslash(Left$(sTmp, i - 1))

End Function

Private Function AddBackslash(s As String) As String

   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If

End Function


