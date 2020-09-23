Attribute VB_Name = "ftpUpDowLoad"
'WICHTIG: Der Upload der Serverliste funkioniert in diesem Source Code nicht,
'da ich das Passwort für den FTP-account weggenommen habe (logisch oder? :) )
'Der Download der Server liste und der Rest des Chats ist natürlich
'weiterhin voll funktionstüchtig.


Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public hOpen As Long, hConnection As Long
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_SERVICE_FTP = 1
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const scUserAgent = "vb wininet"

Public Sub UploadNewServerlist()
  'internetverbindung öffnen
  hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
  'connecten:
  Dim nFlag As Long
  Dim szFileLocal As String
  Dim szFileRemote As String

  dwType = FTP_TRANSFER_TYPE_ASCII

  nFlag = INTERNET_FLAG_PASSIVE

  hConnection = InternetConnect(hOpen, "ftpserver", INTERNET_INVALID_PORT_NUMBER, _
  "Benutzername", "Passwort", INTERNET_SERVICE_FTP, nFlag, 0)
  
  If hConnection = 0 Then
    MsgBox "An error occured while connecting to the internet"
    End
  End If

  'Die datei wieder auf den server uploaden
  bRet = FtpPutFile(hConnection, "Quelldatei", "Zieldatei", _
  dwType, 0)
End Sub

Public Sub GetNgenerateNewServerlist()
  'die benötigten dateien öffen:
  Open App.Path + "\srvrlst.txt" For Output As #1
    
  'Die Serverliste downloaden
  Dim strBuffer As String
  frmHost!Inet.AccessType = icUseDefault
  strBuffer = frmHost!Inet.OpenURL("http://buerger.metropolis.de/nukegod/nukechat/srvrlst.txt", icString)
  
  'formatieren:
  strBuffer = Replace(strBuffer, vbCr, "")
  strBuffer = Replace(strBuffer, vbLf, "")
  
  strBuffer = strBuffer + UsedIP + "/" + hostname + "/" + NickName + "/" + _
   Trim(Str(MaxAnzClients)) + "/"
  
  'die serverliste in die datei schreiben
  Print #1, strBuffer

  Close #1
End Sub

Public Sub DeleteServerFromList()
  'die benötigten dateien öffen:
  Open App.Path + "\srvrlst.txt" For Output As #1
   
  'Die Serverliste downloaden
  Dim strBuffer As String
  frmHost!Inet.AccessType = icUseDefault
  strBuffer = frmHost!Inet.OpenURL("http://buerger.metropolis.de/nukegod/nukechat/srvrlst.txt", icString)
   
  'formatieren:
  strBuffer = Replace(strBuffer, vbCr, "")
  strBuffer = Replace(strBuffer, vbLf, "")

  'Den eigenen Server aus der Liste entfernen
  strBuffer = Replace(strBuffer, UsedIP + "/" + hostname + "/" + NickName + "/" + _
   Trim(Str(MaxAnzClients)) + "/", "")

  'die serverliste in die datei schreiben
  Print #1, strBuffer
    
  Close #1
  
  UploadNewServerlist
End Sub

Public Function GetServerlist()
  'Die Serverliste downloaden
  Dim strBuffer As String
  frmJoin!Inet.AccessType = icUseDefault
  strBuffer = frmJoin!Inet.OpenURL("http://buerger.metropolis.de/nukegod/nukechat/srvrlst.txt", icString)

  'formatieren:
  strBuffer = Replace(strBuffer, vbCr, "")
  strBuffer = Replace(strBuffer, vbLf, "")

  GetServerlist = strBuffer
End Function
