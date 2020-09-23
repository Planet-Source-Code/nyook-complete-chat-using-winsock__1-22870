Attribute VB_Name = "Deklarationen"
'Hier werden die verschiedenen Ports als Konstanten definiert
Public Const message2serverPort As Integer = 15831 'Port für Nachrichten vom Client zum Server
Public Const message2clientPort As Integer = 15832 'Port für Nachrichten vom Server zum Client
Public Const cmd2serverPort As Integer = 15833 'Port für Kommandos vom Client zum Server
Public Const cmd2clientPort As Integer = 15834 'Port für Kommandos vom Server zum Client
Public Const FTPort As Integer = 15835 'Port für den FileTransfer

Public UsedIP As String 'Die verwendete IP, es kann sein, dass der benutzer mehrer IPs hat
Public IP() As String 'Hier werden die IPs gespeichert
Public anzIPs As Integer 'Anzahl der zur verfügung stehenden IPs

Type Clients 'das Objekt clients wird definiert für die Benutzer des Chatraumes
  IP As String 'die IP des Benutzers
  NickName As String 'der Nickname des Benutzers
  FontColor As String 'die Schriftfarbe des Benutzers
  LastMessage As Single 'Zeitpunkt der letzten Nachricht
  anzFloodMsgs As Integer 'Die Anzahl der Nachrichten, die floodverdächtig sind
End Type

Type Servers 'das Objekt servers wird definiert für die Server
  IP As String 'die IP des Servers
  MaxAnzClients As Integer 'die Maximale anzahl an Clients im Chatraum (mit dem Hoster)
  Moderator As String 'und der Moderator (Hoster) des Server
  Name As String 'der Name des Servers
End Type

Type BannList 'Das Objekt BannList, wo die gebannten IPs aufgelistet sind
  IP As String 'Die Gesperrte IP
  Time As String 'Der Zeitpunkt, an dem die Ip gesperrt wurde
  BannMinutes As String 'Die Anzahl Minuten, während denen die Ip gesperrt wird
End Type

Public HostName As String 'der Name des Chatservers
Public ServerMod As String 'der Moderator des Servers
Public ServerIP As String 'die IP des aktuellen Servers

Public NickName As String 'Der Nickname des Benutzers

Public Server(63) As Servers 'das Programm wird die Informationen für maximal 64 Server aufnehmen
Public Client() As Clients 'hier werden die Informationen der Clients gespeichert
Public MaxAnzClients As Integer 'hier wird die maximale anzahl der Clients gespeichert

Public anzClients As Integer 'die Anzahl der Clients
Public anzServers As Integer 'die Anzahl der Servers
Public anzBannedIPs As Integer 'die Anzahl der gesperrten IPs

Public BannedIPs() As BannList 'Hier werden die gesperrten IPs gespeichert

'Die Folgenden zwei Zeilen werden für das Einfügen von Bildern in ein Textfeld benötigt
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302

'-----------------------------------------------------
'Der folgende Abschnitt beinhaltet den nötigen Code, um ein Icon im Systray-Menü
'zu erstellen. Der Code ist nicht von mir, also kann ich nicht genau darauf eingehen
Public Type SysTrayIcon 'der Type für ein Systrayicon
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
'Mit dieser Funktion werden die Icons im Systray gesteuert
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As SysTrayIcon) As Long
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
'noch das Objekt für das Systrayicon
Public LanChatIcon As SysTrayIcon
'-----------------------------------------------------

Public MessageDB() As String 'Hier werden alle eingegebenen Nachrichten gespeichert
Public anzMessages As Integer 'Anzahl aller Nachrichten
Public aktMessage As Integer 'die aktuell gewählte Nachricht aus der NachrichtenDatenbank

Public FileTransferClient As String 'Der Client, zu dem eine Datei versendet werden soll
