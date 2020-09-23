VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmJoin 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Join"
   ClientHeight    =   3705
   ClientLeft      =   4050
   ClientTop       =   3105
   ClientWidth     =   6450
   ControlBox      =   0   'False
   Icon            =   "frmJoin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6450
   Begin VB.Frame framAdvanced 
      Caption         =   "Advanced:"
      Height          =   2415
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cbUsedIP 
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblChat2 
         Caption         =   "Klick on ""Get I-net Chatrooms"" to receive a list of open nukechat- chatrooms all over the world."
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblChat 
         Caption         =   "If you want to chat over the Internet, you must choose your Internet-IP."
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUsedIP 
         Caption         =   "Used IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2760
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSSrequestSnd 
      Left            =   1560
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSSanswerRcv 
      Left            =   840
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtNN 
      Height          =   285
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame framServers 
      Caption         =   "Chatrooms:"
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdSpecify 
         Caption         =   "&Specify"
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdGetInetChats 
         Caption         =   "Get &I-net Chatrooms"
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "LAN - &Refresh"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ListBox lstServers 
         Height          =   2595
         ItemData        =   "frmJoin.frx":0442
         Left            =   120
         List            =   "frmJoin.frx":0444
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblNN 
      Caption         =   "Your Nickname:"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbUsedIP_Click()
  UsedIP = IP(cbUsedIP.ListIndex)
  WSSanswerRcv.Close
  WSSanswerRcv.Bind cmd2clientPort, UsedIP  'ist für den Empfang einer Antwort der Serveranfrage über Port -cmd2clientPort- bereit
  RefreshServerList
End Sub

Private Sub cmdCancel_Click()  'zum hauptformular zurückkehren...
  Load frmMain 'das Haupt-Formular wird geladen
  Unload Me
End Sub

Private Sub cmdGetInetChats_Click()
  'Die Serverliste Downloaden
  Dim Serverliste As String
  Serverliste = GetServerlist
  
  Dim Teil 'Zwischenspeicher
  Teil = Serverliste

  Do While Teil <> ""
    anzServers = anzServers + 1 'es ist ein weiterer Server vorhanden
    
    'die Infos rausfiltern
    Server(anzServers - 1).IP = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Server(anzServers - 1).Name = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Server(anzServers - 1).Moderator = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Server(anzServers - 1).MaxAnzClients = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))

    lstServers.AddItem Server(anzServers - 1).Name 'der Server wird der Server-Liste hinzugefügt
  Loop
End Sub

Public Sub cmdJoin_Click() 'einem Server beitreten...
  If Trim(txtNN.Text) = "" Then 'überprüfen, ob ein Nickname angegeben wurde
    Meldung = MsgBox("Enter your Nickname!", vbOKOnly, "Unable to join") 'wenn nicht, zuerst ein Hinweis an den Benutzer...
    txtNN.SetFocus 'Das Textfeld aktivieren
    'vorhandenen Text auswählen...
    txtNN.SelStart = 0
    txtNN.SelLength = Len(txtNN.Text)
    Exit Sub  '...dann wird die Prozdedur beendet, d. h. der Beitrittsvorgang abgebrochen
  End If
  If lstServers.ListIndex = -1 Then 'prüft, ob ein server ausgewählt ist
    Meldung = MsgBox("No chatroom selected", vbOKOnly, "Unable to join") 'gibt eine meldung, dass kein chatraum ausgewählt ist
    Exit Sub 'beitrittsvorgang abbrechen
  End If
  
  NickName = FormNickName(txtNN.Text) 'eventuelle Leerzeichen links und rechts entfernen
  
  WSSrequestSnd.RemoteHost = Trim(Server(anzServers - 1).IP) 'an die Ip des ausgewählten Servers senden
  WSSrequestSnd.SendData ("RequestChat/" + NickName) 'Chatanfrage stellen
End Sub

Private Sub cmdRefresh_Click() 'der Benutzer will die Liste der Server aktualisieren
  RefreshServerList 'die Serverliste aktualisieren
End Sub

Private Sub cmdSpecify_Click() 'der benutzer möchte die Ip des Servers selbst eingeben. (da er vielleicht über internet chatten will)
  frmJoin.Enabled = False 'das Join-Fenster deaktivieren
  Load frmSpecify 'das Specify-Fenster laden
End Sub

Private Sub Form_Load() 'das Join-Formular wird geladen....
  getIPs 'Die IPs heraussuchen:
  For i = 0 To anzIPs - 1
    cbUsedIP.AddItem IP(i)
  Next i
  UsedIP = IP(0)
  cbUsedIP.ListIndex = 0
  
  Show  'anzeigen der Form
  WSSanswerRcv.Close
  WSSanswerRcv.Bind cmd2clientPort, UsedIP  'ist für den Empfang einer Antwort der Serveranfrage über Port -cmd2clientPort- bereit
  WSSrequestSnd.RemotePort = cmd2serverPort 'der Port für das Senden einer Serveranfrage wird zugewiesen
  RefreshServerList 'die Serverliste wird aktualisiert
End Sub

Private Sub Form_Unload(Cancel As Integer)  'die Form wird geschlossen...
  'verbindungen werden getrennt...
  WSSanswerRcv.Close
  WSSrequestSnd.Close
  
  anzIPs = 0
End Sub

Private Sub WSSanswerRcv_DataArrival(ByVal BytesTotal As Long) 'eine Antwort vom Server...
  Dim Nachricht As String 'in diese Variable wird die Antwort gespeichert
  Dim Teil As String 'diese Variable wird als eine Art Zwischenspeicher benutzt
  
  WSSanswerRcv.GetData Nachricht 'die Antwort wird in der Variable Nachricht gespeichert
   
  If Left(Nachricht, 14) = "NewClientList/" Then 'die Anfrage wurde angenommen und die Clientlist wird empfangen
    On Error Resume Next
    ReDim Client(Server(lstServers.ListIndex).MaxAnzClients) 'das Variablenfeld für die Clients auf die maximale Anzahl von Clients im ausgewählten Chatraum begrenzen
    
    'die Daten werden rausgefiltert...
    Teil = Right(Nachricht, Len(Nachricht) - 14)
    FilterClientList (Teil)
    
    'serverdaten werden in die Variablen geschrieben...
    hostname = lstServers.List(lstServers.ListIndex)
    ServerMod = Server(lstServers.ListIndex).Moderator
    ServerIP = Server(lstServers.ListIndex).IP
            
    'und die Serverinformationen gelöscht
    For i = 0 To anzServers
      Server(i).IP = ""
      Server(i).Moderator = ""
      Server(i).Name = ""
    Next i
    anzServers = 0
    
    'die winsock-verbindungen werden geschlossen
    WSSrequestSnd.Close
    WSSanswerRcv.Close
    
    'für das Chatten vorbereiten...
    Unload Me
    Load frmCChat
    Exit Sub
  End If
  
  
  If Left(Nachricht, 11) = "ServerInfo/" Then 'Infos über den Server
    anzServers = anzServers + 1 'es ist ein weiterer Server vorhanden
    
    'die Infos rausfiltern
    Teil = Right(Nachricht, Len(Nachricht) - 11)
    Server(anzServers - 1).Name = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Server(anzServers - 1).Moderator = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Server(anzServers - 1).MaxAnzClients = Teil
    Server(anzServers - 1).IP = WSSanswerRcv.RemoteHostIP
    
    lstServers.AddItem Server(anzServers - 1).Name 'der Server wird der Server-Liste hinzugefügt
    Exit Sub
  End If
  
  
  If Left(Nachricht, 13) = "UserIsBanned/" Then 'Meldung vom Server, dass der Benutzer gebannt wurde
    Meldung = MsgBox("You were banned from the chatroom for " + Right(Nachricht, Len(Nachricht) - 13) + " minutes", vbCritical, "Unable to join") 'eine Meldung an den Benuzer, dass er für 5 Minuten gebannt wurde
    Exit Sub
  End If
  
  If Nachricht = "RoomIsFull" Then 'Meldung vom Server, dass der ausgewählte Chatraum voll ist (128 Clients)
    Meldung = MsgBox("The selected chatroom is full", vbCritical, "Unable to join") 'eine Meldung an den Benuzer, dass der Chatroom schon voll ist
    Exit Sub
  End If
  
  If Nachricht = "NickNameIsInUse" Then 'Meldung vom Server, dass schon ein Client mit dem gleichen NickNamen im Chat ist
    Meldung = MsgBox("This nickname is already in use." + Chr$(13) + Chr$(10) + "Change your nickname!", vbCritical, "Unable to join") 'Eine Meldung an den Benutzer, dass der Nickname schon von jemand anderem verwendet wird, und dass er seinen Nicknamen ändern soll
    txtNN.SetFocus 'Das Textfeld aktivieren
    'vorhandenen Text auswählen...
    txtNN.SelStart = 0
    txtNN.SelLength = Len(txtNN.Text)
    Exit Sub
  End If
End Sub

Private Sub RefreshServerList() 'diese Prozedur aktualisiert die Serverliste
  On Error Resume Next  'falls ein fehler auftritt, z. b. weil der benutzer mehrmals hintereinander auf den button klickt, einfach weitermachen
  
  frmJoin!lstServers.Clear 'das Listenelement mit den Servern löschen
  
  'es wird an alle Rechner im Netzwerk eine Anforderung für serverinformationen gesendet. der Server wird mit "ServerInfo/..." antworten
  frmJoin!WSSrequestSnd.RemoteHost = "255.255.255.255"
  frmJoin!WSSrequestSnd.SendData "RequestInfo"
  
  On Error GoTo 0  'errors ab jetzt wieder melden
End Sub

