Attribute VB_Name = "ServerFunktionen"
Public Sub SendBroadCast(BCNachricht As String) 'diese Funktion führt einen Rundruf aus
  Dim BC As String
  BC = (".-= " + BCNachricht + " =-.")
  
  'den Rundruf im eigenen Chatfenster anzeigen
  frmSChat!RtxtChat.SelStart = Len(frmSChat!RtxtChat.Text) 'die Texteinfügemarke ans ende des Textes setzen
  frmSChat!RtxtChat.SelColor = vbBlack 'schriftfarbe schwarz auswählen
  frmSChat!RtxtChat.SelText = Chr$(13) + Chr$(10) + BC + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) 'den Rundruf anzeigen
  frmSChat!RtxtChat.SelStart = Len(frmSChat!RtxtChat.Text)  'die Texteinfügemarke an das ende des Textes setzen. das hat zur folge, dass das text-element automatisch nach unten scrollt
  
  For i = 1 To (anzClients - 1) 'an alle clients...
    frmSChat!WSScmdSnd.RemoteHost = Client(i).IP 'an die ip des clients senden
    frmSChat!WSScmdSnd.SendData ("BrdCst/" + BC) 'der Rundruf wird jetzt gesendet
  Next i
End Sub

Public Sub SendClientList() 'diese Prozedur verschickt die aktuelle ClientList zu allen Clients
  Dim NewClientList As String

  NewClientList = "NewClientList" 'die Antwort:
  For i = 0 To (anzClients - 1) 'die Clientlist wird zusammentgestellt = alle daten aller clients
    NewClientList = NewClientList + "/" + Client(i).FontColor + "/" + Client(i).IP + "/" + Client(i).NickName
  Next i
  
  For i = 1 To (anzClients - 1) 'die clientlist wird nun an alle clients verschickt. der neue client erhält also gleich die clientlist, und alle anderen werden ihre clientlist updaten.
    frmSChat!WSScmdSnd.RemoteHost = Client(i).IP
    frmSChat!WSScmdSnd.SendData NewClientList
  Next i
End Sub

Public Sub RemoveClient(ClientNr As String)  'Diese Prozedur entfernt einen Client aus der Clientlist
  If ClientNr = (anzClients - 1) Then 'wenn der zu entfernende Client der letzte in der Clientliste ist...
    'die Infos über den client löschen:
    Client(ClientNr).FontColor = ""
    Client(ClientNr).IP = ""
    Client(ClientNr).NickName = ""
    Client(ClientNr).LastMessage = 0
    Client(ClientNr).anzFloodMsgs = 0
    anzClients = anzClients - 1 'es ist ein Client weniger vorhanden
    Exit Sub
  End If

  For i = ClientNr To (anzClients - 2) 'bps: client 3 wird gelöscht. nun gehen die informationen von client 4 zu client 3 rüber, die von Client 5 zu client 4 usw.
    Client(i).NickName = Client(i + 1).NickName
    Client(i).IP = Client(i + 1).IP
    Client(i).FontColor = Client(i + 1).FontColor
    Client(i).LastMessage = Client(i + 1).LastMessage
    Client(i).anzFloodMsgs = Client(i + 1).anzFloodMsgs
  Next i
  
  'und am schluss die informationen des letzten Client-Platzes löschen, da dieser ja nicht mehr genutzt wird
  Client(anzClients - 1).FontColor = ""
  Client(anzClients - 1).IP = ""
  Client(anzClients - 1).NickName = ""
  Client(ClientNr).LastMessage = 0
  Client(ClientNr).anzFloodMsgs = 0
  anzClients = anzClients - 1 'es ist ein Client weniger vorhanden
End Sub

Public Function IPbanned(NewIP As String) 'Diese Funktion prüft, ob die IP gesperrt ist
  If anzBannedIPs = 0 Then GoTo Weiter 'falls noch keine gesperrten IPs vorhanden sind
  
  For i = 1 To (anzBannedIPs) 'geht alle gesperrten ips durch
    If BannedIPs(i).IP = NewIP Then 'prüft, ob die IP gesperrt wurde
      
      If Not (Timer - BannedIPs(i).Time) > (BannedIPs(i).BannMinutes * 60) Then 'prüft, ob die Bannzeit seit der Sperrung noch nicht vergangen sind
        IPbanned = True 'Die IP ist gesperrt
        Exit Function
      End If
      
      'die gesperrte ip kann wieder freigegeben werden, da die 5 min vergangen sind
      DeleteBannedIP (i) 'die IP aus der Liste wieder entfernen
      GoTo Weiter
      
    End If
  Next i

Weiter:
  IPbanned = False 'Die IP ist nicht gesperrt
End Function

Public Sub DeleteBannedIP(deleteIP As String) 'Diese Prozedur entfernt eine gesperrte IP aus der Liste der gesperrten IPs
  If deleteIP = (anzBannedIPs) Then 'falls die ip die letzte in der liste ist
    anzBannedIPs = anzBannedIPs - 1 'es ist eine gesperrte IP weniger vorhanden
    ReDim BannedIPs(anzBannedIPs) 'die Liste anpassen (um 1 verkleinern)
    Exit Sub 'fertig
  End If

' |4|
' |3|<-löschen   --> |3(4)|
' |2|                |2|
' |1|                |1|
' |-|                |-|
  
  For i = deleteIP To (anzBannedIPs - 1) 'geht alle gesperrten IPs oberhalb der deleteIP durch

    BannedIPs(i).IP = BannedIPs(i + 1).IP
    BannedIPs(i).Time = BannedIPs(i + 1).Time

    anzBannedIPs = (anzBannedIPs - 1)
    ReDim BannedIPs(anzBannedIPs)

  Next i
End Sub

Public Sub KickClient(ClientNr As String) 'Diese Prozedur kickt einen Benutzer aus dem Chatraum
    'dem Benutzer wird vorgeschaukelt, dass der Chatraum geschlossen worden sei
    frmSChat!WSScmdSnd.RemoteHost = Client(ClientNr).IP
    frmSChat!WSScmdSnd.SendData ("ServerClosed")
    
    SendBroadCast (Client(ClientNr).NickName + " was kicked!")  'ein Rundruf, dass ein Client den Chatraum verlässt, wird gesendet
    RemoveClient (ClientNr) 'der Client wird aus der Liste entfernt
    
    'das eigene listenelement wird aktualiesiert...
    frmSChat!lstClients.Clear 'der inhalt des listenelements wird gelöscht...
    For i = 0 To (anzClients - 1) '...und wieder eingelesen
      frmSChat!lstClients.AddItem Client(i).NickName
    Next i
    
    SendClientList 'die neue Clientlist wird versendet
End Sub

Public Sub BannClient(ClientNr As String) 'Diese Prozedur bannt einen Client aus dem Chatraum
  anzBannedIPs = anzBannedIPs + 1 'es ist eine gesperrte IP mehr vorhanden
  ReDim BannedIPs(anzBannedIPs) 'das Datenfeld anpassen
  
  'Daten über die Sperrung aufnehmen...
  BannedIPs(anzBannedIPs).IP = Client(ClientNr).IP 'zu sperrende IP
  BannedIPs(anzBannedIPs).Time = Timer 'Zeitpunkt der Sperrung
  BannedIPs(anzBannedIPs).BannMinutes = Val(frmServerOptions!txtBannMinutes.Text) 'Anzahle der minuten, während denen die IP gesperrt wird
  
  'Dem Client vorgaukeln, dass der Chatraum geschlossen worden sei
  frmSChat!WSScmdSnd.RemoteHost = Client(ClientNr).IP
  frmSChat!WSScmdSnd.SendData ("ServerClosed")
  
  'ein Rundruf, dass ein Client den Chatraum verlässt, wird gesendet
  SendBroadCast (Client(ClientNr).NickName + " was BANNED for" + Str(Val(frmServerOptions!txtBannMinutes.Text)) + " Minutes!")
  
  RemoveClient (ClientNr) 'der Client wird aus der Liste entfernt

  'das eigene listenelement wird aktualiesiert...
  frmSChat!lstClients.Clear 'der inhalt des listenelements wird gelöscht...
  For i = 0 To (anzClients - 1) '...und wieder eingelesen
    frmSChat!lstClients.AddItem Client(i).NickName
  Next i

  SendClientList 'die neue Clientlist wird versendet
End Sub

Public Sub NotifyMessage(ClientNr As String)
  'Diese Prozedur wird aufgerufen, wenn ein Client eine Message versendet.
  'Sie notiert die Zeitpunkte der Messages und kickt wenn nötig Clients
  
  If ClientNr = 0 Then Exit Sub 'falls der Server eine Nachricht gesendet hat
   
  'prüft, ob die letzte Nachricht schon länger als das Limit zurückliegt
  If Timer - Client(ClientNr).LastMessage > (Val(frmServerOptions!txtFIms.Text) / 200) Then
    Client(ClientNr).LastMessage = Timer
    Client(ClientNr).anzFloodMsgs = 0
    Exit Sub
  End If

  'Der Client hat eine (weitere) Floodmessage geschickt
  Client(ClientNr).anzFloodMsgs = Client(ClientNr).anzFloodMsgs + 1
  
  'wenn das Maximum an Floodmessages überschritten wurde, den Client kicken
  If Client(ClientNr).anzFloodMsgs > Val(frmServerOptions!txtMaxMsg.Text) Then KickClient (ClientNr)
  
End Sub
