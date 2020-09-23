VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSChat 
   Caption         =   "nukechat (Server)"
   ClientHeight    =   5370
   ClientLeft      =   2340
   ClientTop       =   2700
   ClientWidth     =   7275
   Icon            =   "frmSChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7275
   Begin VB.PictureBox picSTIconCTRL 
      Height          =   615
      Left            =   3000
      Picture         =   "frmSChat.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer timerBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   3720
      Top             =   7920
   End
   Begin MSWinsockLib.Winsock WSScmdSnd 
      Left            =   720
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSScmdRcv 
      Left            =   240
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSmessageSnd 
      Left            =   720
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSmessageRcv 
      Left            =   240
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSSMessageSnd 
      Left            =   720
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WScmdRcv 
      Left            =   240
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WScmdSnd 
      Left            =   720
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSSMessageRcv 
      Left            =   240
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComDlg.CommonDialog ComDiag 
      Left            =   6720
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame framClients 
      Caption         =   "Clients"
      Height          =   5250
      Left            =   5280
      TabIndex        =   2
      Top             =   -10
      Width           =   1935
      Begin VB.ListBox lstClients 
         Height          =   4935
         ItemData        =   "frmSChat.frx":0884
         Left            =   120
         List            =   "frmSChat.frx":0886
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   5175
   End
   Begin RichTextLib.RichTextBox RtxtChat 
      Height          =   5025
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8864
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmSChat.frx":0888
   End
   Begin VB.Image imgSTIcon 
      Height          =   480
      Index           =   1
      Left            =   4680
      Picture         =   "frmSChat.frx":0951
      Top             =   7920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSTIcon 
      Height          =   480
      Index           =   0
      Left            =   4200
      Picture         =   "frmSChat.frx":0D93
      Top             =   7920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuCloseChat 
      Caption         =   "&Close Chat"
   End
   Begin VB.Menu mnuChatT 
      Caption         =   "Chattext"
      Begin VB.Menu mnuChatTClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuChatTSave 
         Caption         =   "&Save in *.rtf"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuFT 
      Caption         =   "&FileTransfer"
      Begin VB.Menu mnuFTReceive 
         Caption         =   "&Receive File..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewClientlist 
         Caption         =   "Show &Clientlist"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsColor 
         Caption         =   "&Color"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuOptionsSTI 
         Caption         =   "&SysTray Icon"
         Begin VB.Menu mnuOptionsSTIMinimize 
            Caption         =   "&Minimize to SysTray Icon"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuOptionsSTIBlink 
            Caption         =   "&Blink on Message"
            Checked         =   -1  'True
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuOptionsServeroptions 
         Caption         =   "Show &Server Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpCommands 
         Caption         =   "&Commands"
      End
      Begin VB.Menu mnuHelpEmoticons 
         Caption         =   "&Emoticons"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpBann 
         Caption         =   "&Bann for 5 min"
      End
      Begin VB.Menu mnuPopUpKick 
         Caption         =   "&Kick!"
      End
      Begin VB.Menu mnuPopUpSend 
         Caption         =   "&Send File"
      End
      Begin VB.Menu mnuPopUpWhisper 
         Caption         =   "&Whisper"
      End
   End
   Begin VB.Menu mnuSysTrayPopUp 
      Caption         =   "&SysTrayPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayPopUpOpen 
         Caption         =   "&Open"
      End
   End
End
Attribute VB_Name = "frmSChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load() 'das Formular wird geladen...
  Show '...angezeigt...
  '...und den Winsock-Controls werden die Ports zugewiesen
  WSSMessageRcv.Bind message2serverPort, UsedIP
  WSSMessageSnd.RemotePort = message2clientPort
  WScmdRcv.Bind cmd2clientPort, UsedIP
  WScmdSnd.RemotePort = cmd2serverPort 'wird während der Programmausführung manchmal an cmd2clientport gebunden
  WSScmdRcv.Bind cmd2serverPort, UsedIP
  WSScmdSnd.RemotePort = cmd2clientPort
  WSmessageRcv.Bind message2clientPort, UsedIP
  WSmessageSnd.RemotePort = message2serverPort
  
  WScmdSnd.RemoteHost = ServerIP
  WSmessageSnd.RemoteHost = ServerIP

  lstClients.AddItem Client(0).NickName 'der client list wird ein client (du selbst) hinzugefügt
  
  frmSChat.Caption = "nukechat (Server)   -   " + hostname 'der name des Chatraumes wird in der Titelleiste angezeigt

  aktMessage = 1
End Sub

Public Sub Form_Resize() 'Die Grösse des Chatforumlars wurde vom Benutzer geändert
  'Die Position der Objekte anpassen
  On Error Resume Next
  Select Case mnuViewClientlist.Checked
    Case True
      framClients.Visible = True
      lstClients.Visible = True
      framClients.Left = Me.Width - framClients.Width - 130
      framClients.Height = Me.Height - 700
      lstClients.Height = framClients.Height - 350
      txtMessage.Top = Me.Height - 1000
      txtMessage.Width = Me.Width - (Me.Width - framClients.Left) - 60
      RtxtChat.Width = Me.Width - (Me.Width - framClients.Left) - 57
      RtxtChat.Height = Me.Width - (Me.Width - txtMessage.Top) - 30
    Case False
      framClients.Visible = False
      lstClients.Visible = False
      txtMessage.Top = Me.Height - 1000
      txtMessage.Width = Me.Width - 130
      RtxtChat.Width = Me.Width - 125
      RtxtChat.Height = Me.Width - (Me.Width - txtMessage.Top) - 30
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer) 'der Chatraum wird geschlossen
  'Meldung an alle Clients, dass der Chatraum geschlossen wird.
  If anzlcients = 1 Then GoTo Weiter
  For i = 1 To (anzClients - 1)
    WSScmdSnd.RemoteHost = Client(i).IP
    WSScmdSnd.SendData "ServerClosed"
  Next i
  
  'Die gesperrten IPs werden gelöscht...
  ReDim BannedIPs(0)
  BannedIPs(0).IP = ""
  BannedIPs(0).Time = ""
  BannedIPs(0).BannMinutes = ""
  anzBannedIPs = 0
  
Weiter:
  'Die Verbindungen geschlossen...
  WSSMessageRcv.Close
  WSSMessageSnd.Close
  WScmdRcv.Close
  WScmdSnd.Close
  WSScmdRcv.Close
  WSScmdSnd.Close
  WSmessageRcv.Close
  WSmessageSnd.Close
 
  DeleteClientInformations
  
  Unload frmCommands 'das Chatcommands-Fenster beenden (falls es überhaupt geöffnet ist)
  Unload frmEmoticons 'das Emoticons-Fenster beenden (falls es überhaupt geöffnet ist)
  Unload frmAbout
  Unload frmServerOptions
  Unload frmSend
  Unload frmReceive

  If frmHost!chkInternetChat.Value = 1 Then DeleteServerFromList

  frmHost.Visible = True '...und das Host-Formular wieder Sichtbar gemacht
End Sub


Private Sub lstClients_Click() 'wenn der Benutzer auf das Client-Listenelement klickt
  If NickName = lstClients.List(lstClients.ListIndex) Then Exit Sub 'wenn der Benuzter auf seinen eigenen Nicknamen geklickt hat, das PopUp-Menü nicht anzeigen
  mnuPopUpBann.Caption = ("Bann for" + Str(Val(frmServerOptions!txtBannMinutes)) + " Minutes") 'Die Anzahl der Bannminuten richtig anzeigen
  PopupMenu mnuPopUp 'das Popup menü anzeigen
End Sub

Private Sub mnuChatTClear_Click() 'Den Inhalt des Textfeldes löschen
  RtxtChat.Text = ""
End Sub

Private Sub mnuChatTSave_Click() 'Den Chattext speichern
  'ComDiag.Filter [=RTF-Dokument (*.rtf)|*.rtf] 'Dateityp für das Speichern-Formular festlegen
  ComDiag.ShowSave 'Das Speichernformular anzeigen
  RtxtChat.SaveFile ComDiag.FileName, rtfRTF 'Den Chattext in der angegebenen Datei speichern
End Sub

Private Sub mnuCloseChat_Click() 'beenden des Servers und zurückkehren zum Host-Forumlar
  Dim Antwort As String
  
  Antwort = MsgBox("Do you really want to close the Chatroom?", vbExclamation Or vbYesNo, "Warning")
  If Antwort = vbNo Then Exit Sub
  
  Unload Me
End Sub

Private Sub mnuFTReceive_Click()
  Load frmReceive
End Sub

Private Sub mnuHelpAbout_Click() 'Der benutzer möchte Informationen über den nukechat
  Load frmAbout
End Sub

Private Sub mnuHelpCommands_Click() 'Der benutzer möchte eine Auflistung aller Kommandos haben (z.b. /me , /w)
  Load frmCommands
End Sub

Private Sub mnuHelpEmoticons_Click() 'Lädt das Emoticons-Formular, wo die verschiedenen Emoticons beschrieben sind
  Load frmEmoticons
  frmEmoticons.Show
End Sub

Private Sub mnuOptionsColor_Click() 'der Benutzer möchte seine Farbe ändern
  ComDiag.ShowColor 'das Standart-Formular für Farben anzeigen
  WScmdSnd.SendData ("NewColor/" + NickName + "/" + Str(ComDiag.Color))  'dem Server mitteilen, dass der Benutzer eine neue Farbe gewählt hat
End Sub

Private Sub mnuOptionsServeroptions_Click()
  'Der Benutzer möchte die Server-Optionen sichtbar machen
  frmServerOptions.Visible = True
End Sub

Private Sub mnuOptionsSTIBlink_Click()
  'setzt bzw. enfternt das Häckchen
  If mnuOptionsSTIBlink.Checked = True Then
    mnuOptionsSTIBlink.Checked = False
    Exit Sub
  End If
  mnuOptionsSTIBlink.Checked = True
End Sub

Private Sub mnuOptionsSTIMinimize_Click() 'Der Benutzer möchte den Chat zu einem Systrayicon minimieren
  With LanChatIcon
    .cbSize = Len(LanChatIcon)
    .hwnd = picSTIconCTRL.hwnd 'an welches Fenster sollen die Nachrichten?
    .uID = 2& 'unveränderlich
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = &H200 '= MouseMove
    .hIcon = imgSTIcon(0).Picture 'hier bekommen wir das Bild/Icon her
    .szTip = "nukechat v2.1 by nukegod@gmx.net" + vbNullChar
  End With

  Shell_NotifyIcon NIM_ADD, LanChatIcon 'Das SysTrayIcon erstellen
  frmSChat.Visible = False 'Das Chatfenster unsichtbar machen
End Sub

Private Sub mnuPopUpBann_Click() 'Der Benutzer möchte einen teilnehmer für 5 Minuten aus dem Chat ausschliessen
  BannClient (lstClients.ListIndex) 'Den Client bannen
End Sub

Private Sub mnuPopUpKick_Click() 'der Benutzer will einen Client aus dem Chatraum entfernen (kicken):
  KickClient (lstClients.ListIndex) 'Den Client kicken
End Sub

Private Sub mnuPopUpSend_Click() 'Der Benutzer will jemandem eine Datei schicken
  'vermerken, an wen die Datei gesendet werden soll
  FileTransferClient = lstClients.List(lstClients.ListIndex)
  Load frmSend
End Sub

Private Sub mnuPopUpWhisper_Click() 'der Benutzer möchte zu einem Client flüstern
  If Client(lstClients.ListIndex).NickName = NickName Then 'prüft, ob der Benutzer seinen eingen NickNamen angeklickt hat
    Exit Sub 'abbrechen
  End If
  txtMessage.Text = ("/w " + Client(lstClients.ListIndex).NickName + " ") 'den erforderlichen Text im Nachricht-Textfeld einfügen
  txtMessage.SetFocus 'Den Fokus auf das Texteingabefeld setzen, damit der Benutzer gleich schreiben kann
End Sub

Private Sub mnuSysTrayPopUpOpen_Click() 'Der Benutzer klickt auf Öffnen im SysTrayIcon-PopUpMenü
  Shell_NotifyIcon NIM_DELETE, LanChatIcon 'Das SysTrayIcon wegnehmen
  frmSChat.Visible = True 'das Chatfenster wieder Sichtbar machen
  timerBlink.Enabled = False 'Das Blinken wieder ausschalten
End Sub

Private Sub mnuViewClientlist_Click()
  'Setzt bzw. entfernt das häckchen
  If mnuViewClientlist.Checked = True Then
    mnuViewClientlist.Checked = False
    Form_Resize 'Das Formular aktualisieren
  Exit Sub
  End If
  mnuViewClientlist.Checked = True
  Form_Resize 'Das Formular aktualisieren
End Sub

Private Sub picSTIconCTRL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Diese Prozedur ist nötig, um ein PopUp-Menü anzuzeigen, wenn der Benutzer auf
  'das Systrayicon klickt. Fragt mich nicht, weshalb so :)
  Select Case Hex(X)
    Case "1E3C"
      PopupMenu mnuSysTrayPopUp 'Das PopUpMenü anzeigen
  End Select
End Sub

Private Sub timerBlink_Timer() 'das Icon wechseln...
  If LanChatIcon.hIcon = imgSTIcon(0).Picture Then
    LanChatIcon.hIcon = imgSTIcon(1).Picture
    GoTo Weiter
  End If
  LanChatIcon.hIcon = imgSTIcon(0).Picture
Weiter:
  Shell_NotifyIcon NIM_MODIFY, LanChatIcon
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then 'wenn Nach-Unten gedrückt wurde:
    'die nachherige nachricht aus der Nachrichtendatenbank anzeigen
    If (aktMessage - 1) < 1 Then Exit Sub
    aktMessage = aktMessage - 1
    txtMessage.Text = MessageDB(aktMessage)
    Exit Sub
  End If

  If KeyCode = vbKeyDown Then 'wenn Nach-Oben gedrückt wurde:
    'die vorherige Nachricht aus der Nachrichtendatenbank anzeigen
    If (aktMessage) = (anzMessages + 1) Then Exit Sub
    If (aktMessage + 1) = (anzMessages + 1) Then
      aktMessage = aktMessage + 1
      txtMessage.Text = ""
      Exit Sub
    End If
    aktMessage = aktMessage + 1
    txtMessage.Text = MessageDB(aktMessage)
    Exit Sub
  End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer) 'Tastendruck im Textfeld für die Nachrichten
  Dim Teil As String 'diese Variable dient als eine Art Zwischenspeicher
  Dim WhisperClient As String 'der Client, zu dem eventuell geflüstert wird.
  Dim WhisperMessage As String 'die Nachricht, die geflüstert werden soll
  Dim WhisperClientIP As String 'die Ip des Clienten, zu dem geflüstert werden soll

  If KeyAscii = "13" Then 'wenn Enter gedrückt wurde:
    If Trim(txtMessage.Text) = "" Then Exit Sub 'wenn im Textfeld kein Text oder nur Leerzeichen vorhanden sind

    'die Nachricht der Datenbank hinzufügen
    anzMessages = anzMessages + 1
    ReDim Preserve MessageDB(anzMessages)
    MessageDB(anzMessages) = txtMessage.Text
    aktMessage = anzMessages + 1

    If Left(Trim(txtMessage.Text), 4) = "/me " Then 'prüft, ob der Benutzer in der dritten Person sprechen will
      Teil = Trim(txtMessage.Text) 'Leerzeichen links und rechts entfernen
      Teil = Trim(Right(Teil, Len(Teil) - 3)) 'die Mitteilung in der dritten Person
      If Teil = "" Then GoTo NormalNachricht 'falls keine 3.Person-Nachricht eingegeben wurde, den Text als normale Nachricht senden
      WScmdSnd.SendData ("ThrdP/" + NickName + "/" + Teil) 'das 3.Person-Kommando mit der 3.Person-Nachricht zusammen an den Server senden
      txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
      Exit Sub
    End If

    If Left(LTrim(txtMessage.Text), 3) = "/w " Then 'prüft, ob der Benutzer flüstern möchte
      Teil = Trim(Right(Trim(txtMessage.Text), Len(Trim(txtMessage.Text)) - 3)) 'die Nachricht ohne das "/w " speichern
      If Teil = "" Then GoTo NormalNachricht 'prüft, ob nur "/w" eingegeben wurde. wenn ja, dann die nachricht normal senden
      
      WhisperClient = Trim(Left(Teil, InStr(1, Teil, " "))) 'Liest den Client raus, an den die Nachricht gesendet wird.
      If WhisperClient = "" Then GoTo NormalNachricht 'wenn der benutzer nur "/W CLIENTNAME" eingegeben hat, dann die nachricht normal senden
      WhisperMessage = Trim(Right(Teil, Len(Teil) - Len(WhisperClient))) 'liest die Nachricht, die gesendet werden soll raus
      
      If WhisperClient = NickName Then GoTo NormalNachricht 'prüft, ob der Benutzer zu sich selbst flüstern wollte
      If WhisperMessage = "" Then GoTo NormalNachricht 'prüft, ob eine Flüster-Nachricht eingetippt wurde. wenn nicht, die nachricht normal senden
      
      WhisperClientIP = "" 'Die Ip auf *nichts* setzen....
      For i = 0 To (anzClients - 1) 'den Client raussuchen, zu dem geflüstert werden soll
        If Client(i).NickName = WhisperClient Then
          WhisperClientIP = Client(i).IP 'wenn der Client gefunden wurde, seine IP in die Variable speichern
          Exit For
        End If
      Next i
      If WhisperClientIP = "" Then GoTo NormalNachricht 'wenn der WhisperClient gar nicht vorhanden ist, dann die nachricht normal senden

      WScmdSnd.RemotePort = cmd2clientPort 'an den cmd2clientport senden, da das FlüsterKommando direkt an einen Client gesendet wird
      WScmdSnd.RemoteHost = WhisperClientIP 'die Ip des Clienten setzen, zu dem geflüstert werden soll
      WScmdSnd.SendData ("Whisper/" + NickName + "/" + WhisperMessage)  'das Flüster-Kommando senden.
      WScmdSnd.RemoteHost = ServerIP 'das Kommando-Senden-Winsock-Control wieder auf die IP des Servers setzen
      WScmdSnd.RemotePort = cmd2serverPort 'wieder auf den cmd2serverport setzen. (falls später ein Kommand an den Server gesendet werden muss
      
      'im eigenen Chatfenster anzeigen, dass man geflüstert hat:
      RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das Ende des Textes setzen
      RtxtChat.SelColor = vbBlack 'Schriftfarbe schwarz
      RtxtChat.SelText = ("You whisper to " + WhisperClient + ": ")
      For i = 0 To (anzClients - 1) 'sucht die Clientnummer des Benutzers raus
        If Client(i).NickName = NickName Then
          ClientNr = i
          Exit For
        End If
      Next i
      RtxtChat.SelColor = Client(ClientNr).FontColor 'Setzt die eigene Schriftfarbe
      RtxtChat.SelItalic = True 'Kursivschrift aktivieren
      asd = FilterText(WhisperMessage, Client(i).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
      RtxtChat.SelItalic = False 'Kursivschrift deaktivieren
      RtxtChat.SelStart = Len(RtxtChat.Text)
      
      txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
      Exit Sub
    End If

    If Left(Trim(txtMessage.Text), 6) = "/kick " Then 'prüft, ob der Benutzer einen Client kicken will
      Teil = Trim(txtMessage.Text) 'Leerzeichen links und rechts entfernen
      Teil = Trim(Right(Teil, Len(Teil) - 6)) 'der Benutzer, der gekickt werden soll
      If Teil = "" Then GoTo NormalNachricht 'falls kein Client eingegeben wurde, den Text als normale Nachricht senden
      If Teil = NickName Then GoTo NormalNachricht 'Wenn der Benutzer sich selbst kicken wollte
      
      For i = 0 To (anzClients - 1) 'prüft, ob der der zu kickende Benutzer überhaupt existiert
        If Client(i).NickName = Teil Then
          KickClient (i) 'Den Benutzer kicken
          GoTo weiterKick
        End If
      Next i
      GoTo NormalNachricht

weiterKick:
      txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
      Exit Sub
    End If

    If Left(Trim(txtMessage.Text), 6) = "/bann " Then 'prüft, ob der Benutzer einen Client kicken will
      Teil = Trim(txtMessage.Text) 'Leerzeichen links und rechts entfernen
      Teil = Trim(Right(Teil, Len(Teil) - 6)) 'der Benutzer, der gekickt werden soll
      If Teil = "" Then GoTo NormalNachricht 'falls kein Client eingegeben wurde, den Text als normale Nachricht senden
      If Teil = NickName Then GoTo NormalNachricht 'Wenn der Benutzer sich selbst kicken wollte

      For i = 0 To (anzClients - 1) 'prüft, ob der der zu kickende Benutzer überhaupt existiert
        If Client(i).NickName = Teil Then
          BannClient (i) 'Den Benutzer kicken
          GoTo weiterBann
        End If
      Next i
      GoTo NormalNachricht

weiterBann:
      txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
      Exit Sub
    End If

    If Trim(txtMessage.Text) = "/listclients" Then 'der benutzer möchte eine auflistung aller Clients
      RtxtChat.SelColor = vbBlack
      RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das Ende des Textes setzen
      RtxtChat.SelText = Chr$(13) + Chr$(10) + "----------------------------------------"
      RtxtChat.SelText = (Chr$(13) + Chr$(10) + "Clients in '" + hostname + "':" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10))
      For i = 0 To (anzClients - 1)
        RtxtChat.SelColor = vbBlack
        RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das Ende des Textes setzen
        RtxtChat.SelText = Str(i + 1) + ". "
        RtxtChat.SelColor = Client(i).FontColor
        RtxtChat.SelText = Client(i).NickName + Chr$(13) + Chr$(10)
      Next i
      RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das Ende des Textes setzen
      RtxtChat.SelColor = vbBlack
      RtxtChat.SelText = "----------------------------------------" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
      
      txtMessage.Text = ""
      KeyAscii = 0 'Den Windows-Ding unterdrücken
      Exit Sub
    End If
  
    If Trim(txtMessage.Text) = "/clear" Then 'Der benutzer möchte den Inhalt des Chattextfeldes löschen
      RtxtChat.Text = "" 'Inhalt leeren
      txtMessage.Text = ""
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
     Exit Sub
   End If
  
NormalNachricht:
    WSmessageSnd.SendData NickName + "/" + txtMessage.Text 'Nachricht an den Server senden. Die Nachricht setzt sich aus dem Nicknamen, einem / und der Nachricht zusammen
    txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert

    KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
  End If
End Sub

Private Sub WScmdRcv_DataArrival(ByVal BytesTotal As Long) 'es wird ein Kommando erhalten (vom Server oder von einem Client)
  Dim Nachricht As String 'das Kommando in der Variablen Nachricht speichern
  Dim Teil As String 'Zwischenspeicher
  Dim cmdClient As String '(bei Whisper: der Client, der flüstert, bei ThrdP: der Client, der eine Flüster-Nachricht gesendet hat
  Dim WhisperMessage As String 'die Nachicht, die geflüstert wird
  Dim ClientNr As String 'Die Clientnummer (wird bei whisper oder ThrdP verwendet
  Dim ThrdPClient As String 'der Client, der eine ThrdP-Nachricht gesendet hat
  
  WScmdRcv.GetData Nachricht 'Die Nachricht speichern
  
  If Left(Nachricht, 8) = "Whisper/" Then  'prüft, ob jemand zum Benutzer flüstert
    Teil = Right(Nachricht, Len(Nachricht) - 8)
    cmdClient = Left(Teil, InStr(1, Teil, "/") - 1) 'filtert den Namen des Clients raus, der flüstert
    WhisperMessage = Right(Teil, Len(Teil) - InStr(1, Teil, "/")) 'filtert die geflüsterte Nachricht heraus
    
    For i = 0 To (anzClients - 1) 'sucht die Clientnummer des Whisperclients heraus
      If Client(i).NickName = cmdClient Then
        ClientNr = i
        Exit For
      End If
    Next i
    
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke ans Ende des Textes setzen
    RtxtChat.SelColor = vbBlack 'Schriftfarbe auf schwarz (standart)
    RtxtChat.SelItalic = True 'Auf Kursivschrift stellen
    RtxtChat.SelText = (cmdClient + " whispers: ") 'Flüsterankündigung ausgeben
    RtxtChat.SelColor = Client(ClientNr).FontColor 'setzt die Schriftfarbe auf die Farbe des flüsterneden Clients
    weissnicht = FilterText(WhisperMessage, Client(i).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
    RtxtChat.SelItalic = False 'Kursivschrift deaktivieren
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das ende setzen
    'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
    'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
    If frmSChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
      timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
    End If
    Exit Sub
  End If
  
  If Left(Nachricht, 6) = "ThrdP/" Then 'ein Benutzer sendet eine 3.Person-Nachricht
    Teil = Right(Nachricht, Len(Nachricht) - 6) 'das "ThrdP/" -rausschneiden-
    cmdClient = Left(Teil, InStr(1, Teil, "/") - 1) 'der Nickname des Clients, der die 3.Person-Nachricht sendet.
    Teil = Right(Teil, Len(Teil) - Len(cmdClient) - 1) 'die 3.Person-Nachricht
    
    For i = 0 To (anzClients - 1) 'sucht die IP des cmdClients
      If Client(i).NickName = cmdClient Then
        ClientNr = i
      End If
    Next i
    
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das ende setzen
    RtxtChat.SelColor = Client(ClientNr).FontColor 'auf die Schriftfarbe des Clients setzen
    RtxtChat.SelItalic = True 'auf kursivschrift setzen
    RtxtChat.SelText = (cmdClient + " ") 'die 3.Person-Nachricht anzeigen
    weissnicht = FilterText(Teil, Client(ClientNr).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
    RtxtChat.SelItalic = False 'wieder auf normalschrift setzen
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das ende setzen
    'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
    'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
    If frmSChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
      timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
    End If
    Exit Sub
  End If
End Sub

Private Sub WSmessageRcv_DataArrival(ByVal BytesTotal As Long) 'eine Nachricht VOM Server ist angekommen
  Dim Nachricht As String
  WSmessageRcv.GetData Nachricht 'Nachricht speichern
  For i = 0 To (anzClients - 1) 'Es wird geprüft, von wem die Nachricht ist.
    If Left(Nachricht, Len(Client(i).NickName) + 1) = (Client(i).NickName + "/") Then 'wenn man weiss, von welchem Client die Nachricht ist
      RtxtChat.SelStart = Len(RtxtChat.Text) 'die Einfügemarke wird an das Ende des Textes gesetzt
      RtxtChat.SelColor = Client(i).FontColor 'die Schriftfarbe wird gesetzt
      RtxtChat.SelText = (Client(i).NickName + ": ")
      weissnicht = FilterText(Right(Nachricht, Len(Nachricht) - Len(Client(i).NickName) - 1), Client(i).FontColor) 'die Nachricht im Chatfenster anzeigen
      RtxtChat.SelStart = Len(RtxtChat.Text) 'die Texteinfügemarke an das ende des Textes setzen. das hat zur folge, dass das text-element automatisch nach unten scrollt
      'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
      'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
      If frmSChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
        timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
      End If
    End If
  Next i
End Sub

Private Sub WSScmdRcv_DataArrival(ByVal BytesTotal As Long) 'Empfang von Kommandos (CMDs) von einem Client
  Dim Nachricht As String
  WSScmdRcv.GetData Nachricht  'Nachricht speichern
    
  If Left(Nachricht, 9) = "NewColor/" Then 'prüft, ob ein Client seine Schriftfarbe geändert hat.
    Dim AktClient As String
    Dim Teil As String
    
    Teil = Right(Nachricht, Len(Nachricht) - 9)
    AktClient = Left(Teil, InStr(1, Teil, "/") - 1) 'den namen des Clienten rausfiltern, der seine Farbe geändert hat
    
    For i = 0 To anzClients - 1 'die index-nummer des clients rausfiltern
      If Client(i).NickName = AktClient Then
        Client(i).FontColor = Right(Teil, Len(Teil) - InStr(1, Teil, "/")) 'die neue Farbe dem Clienten zuweisen
      End If
    Next i
    
    SendClientList 'die Clientlist versenden
    Exit Sub
  End If
  
  
  If Nachricht = "RequestInfo" Then 'der Hostname wird angefordert und gesendet
    WSScmdSnd.RemoteHost = WSScmdRcv.RemoteHostIP 'cmdsnd soll an die ip senden, von der cmdRcv empfangen hat
    WSScmdSnd.SendData ("ServerInfo/" + hostname + "/" + NickName + "/" + Trim(Str(MaxAnzClients))) 'Informationen über den Server senden: NameDesChatRaums / Hoster / MaximaleAnzahlVonTeilnehmern
    Exit Sub
  End If
  
  
  If Left(Nachricht, 10) = "LeaveChat/" Then 'wenn ein benutzer denn chatraum verlassen möchte...
    For i = 1 To (anzClients - 1) 'den Namen des Clients rausfinden, der den chatraum verlassen will
      If Client(i).NickName = Right(Nachricht, Len(Nachricht) - 10) Then 'prüft, welcher Client den Chatraum verlassen will
        SendBroadCast (Client(i).NickName + " has left the chatroom") 'ein Rundruf, dass ein Client den Chatraum verlässt, wird gesendet
        
        RemoveClient (i) 'der Client wird aus der Liste entfernt
        
        'das eigene listenelement wird aktualiesiert...
        lstClients.Clear 'der inhalt des listenelements wird gelöscht...
        For X = 0 To (anzClients - 1) '...und wieder eingelesen
          lstClients.AddItem Client(X).NickName
        Next X
        
        SendClientList 'die neue Clientlist wird versendet
        Exit Sub
      End If
    Next i
  End If
  
  
  If Left(Nachricht, 12) = "RequestChat/" Then 'der Client will dem Chatraum beitreten
    If IPbanned(WSScmdRcv.RemoteHostIP) = True Then  'prüft, ob die ip gesperrt wurde
      For i = 1 To anzBannedIPs
        If BannedIPs(i).IP = WSScmdRcv.RemoteHostIP Then
          WSScmdSnd.SendData ("UserIsBanned/" + BannedIPs(i).BannMinutes) 'eine Meldung zurückschicken, dass der Benutzer gebannt wurde + die Anzahl Minuten die er gebannt ist
          Exit For
        End If
      Next i
      Exit Sub 'und den neuen Client nicht in den Chatraum lassen
    End If
    
    If anzClients = MaxAnzClients Then 'prüft, ob noch ein client-platz frei ist
      WSScmdSnd.SendData ("RoomIsFull") 'wenn nicht, eine Meldung zurückschicken, dass der Chatraum voll ist (128 Clients)
      Exit Sub 'und den neuen Client nicht in den Chatraum lassen
    End If
    
    For i = 0 To (anzClients - 1) 'prüft, ob schon ein Client mit dem Nicknamen, den der neue Client gewählt hat, vorhanden ist
      If LCase(Right(Nachricht, Len(Nachricht) - 12)) = LCase(Client(i).NickName) Then 'wenn der Nickname schon von einem Clienten benutzt wird...
        WSScmdSnd.SendData ("NickNameIsInUse") '...eine Meldung an den Benutzer geben, dass der Nickname schon benutzt wird
        Exit Sub 'und den Benutzer nicht am Chat teilnehmen lassen
      End If
    Next i
    
    anzClients = anzClients + 1 'es ist ein client mehr vorhanden
    
    With Client(anzClients - 1) 'den neuen Client "registrieren"
      .FontColor = vbBlack 'die Schriftfarbe ist Schwarz (standart)
      .IP = WSScmdRcv.RemoteHostIP 'seine ip
      .NickName = Right(Nachricht, Len(Nachricht) - 12) 'sein Nickname
    End With
    
    SendClientList 'die neue Clientlist versenden
    
    'das eigene listenelement wird aktualiesiert...
    lstClients.Clear 'der inhalt des listenelements wird gelöscht...
    For i = 0 To (anzClients - 1) '...und wieder eingelesen
      lstClients.AddItem Client(i).NickName
    Next i
  
    SendBroadCast (Client(anzClients - 1).NickName + " enters the chatroom") 'einen Rundruf senden, dass ein neuer Benutzer dem Chatraum beigetreten ist
    Exit Sub
  End If
  
  If Left(Nachricht, 6) = "ThrdP/" Then 'ein Client möchte eine 3.Person-Nachrichten senden
    
    For i = 0 To (anzClients - 1) 'für alle Clients...
      WSScmdSnd.RemoteHost = Client(i).IP 'Das Winsock-Control des Servers für das Senden von Commands nimmt die Ip des Clients i an
      WSScmdSnd.SendData Nachricht 'die Nachricht wird gesendet
    Next i 'u.s.w...
    
    'prüfen, von welchem Client die Message ist, und auf Flood überprüfen
    For i = 1 To (anzClients - 1)
      If Left(Nachricht, InStr(1, Nachricht, "/") - 1) = Client(i).NickName Then
        NotifyMessage (i)
        Exit Sub
      End If
    Next i
    
    Exit Sub
  End If
End Sub

Private Sub WSSMessageRcv_DataArrival(ByVal BytesTotal As Long) 'eine Nachricht FÜR den Server ist angekommen
  Dim Nachricht As String
  WSSMessageRcv.GetData Nachricht 'die Empfangene Nachricht in der Variable Nachricht speichern
  
  If frmServerOptions!chkMaxMsgL.Value = 1 Then 'wenn die Maximale Nachrichtenlänge festgelegt ist
    If Len(Right(Nachricht, Len(Nachricht) - InStr(1, "nachricht", "/"))) > Val(frmServerOptions!txtMaxMsgL.Text) Then 'wenn die Nachrichtenlänge grösser ist als das maximum
      Nachricht = Left(Nachricht, InStr(1, Nachricht, "/") + Val(frmServerOptions!txtMaxMsgL.Text)) 'Die Nachricht wird begrenzt
    End If
  End If
     
  For i = 0 To (anzClients - 1) 'für alle Clients...
    WSSMessageSnd.RemoteHost = Client(i).IP 'Das Winsock-Control des Servers für das Senden von Nachrichten nimmt die Ip des Clients i an
    WSSMessageSnd.SendData Nachricht 'die Nachricht wird gesendet
  Next i 'u.s.w...
  
  'prüfen, von welchem Client die Message ist, und auf Flood überprüfen
  For i = 1 To (anzClients - 1)
    If Left(Nachricht, InStr(1, Nachricht, "/") - 1) = Client(i).NickName Then
      NotifyMessage (i)
      Exit Sub
    End If
  Next i
End Sub

Public Function FilterText(FilterMsg As String, ClientColor As String) 'Die Nachricht wird gefiltert, d.h. Smileys und anderes wird durch Grafiken ersetzt
  Dim Semikolon As Integer
  Dim Doppelpunkt As Integer
  Dim Teil As String 'Zwischenspeicher
  Teil = FilterMsg
  
Anfang:
  If InStr(1, Teil, ":") = 0 And InStr(1, Teil, ";") = 0 Then GoTo Weiter 'wenn keine : oder ; im Text vorkommen, normal weitermachen
  
  'folgende Zeilen entscheiden, ob zuerst ein ; oder ein : da ist...
  Semikolon = InStr(1, Teil, ";") 'Position des Semikolons
  Doppelpunkt = InStr(1, Teil, ":") 'Doppelpunkt
  If Semikolon = 0 Then
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
    RtxtChat.SelColor = ClientColor 'Die Schriftfarbe wieder setzten, da sie beim einfügen des bildes verlorengeht
    RtxtChat.SelText = (Left(Teil, InStr(1, Teil, ":") - 1)) 'Den ganzen Text bis zum ersten : einfügen
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, ":") + 1) 'den ganzen Text links vom : wegnehmen
  ElseIf Doppelpunkt = 0 Then
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
    RtxtChat.SelColor = ClientColor 'Die Schriftfarbe wieder setzten, da sie beim einfügen des bildes verlorengeht
    RtxtChat.SelText = (Left(Teil, InStr(1, Teil, ";") - 1)) 'Den ganzen Text bis zum ersten : einfügen
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, ";") + 1) 'den ganzen Text links vom : wegnehmen
  ElseIf (Semikolon - Doppelpunkt) > 0 Then
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
    RtxtChat.SelColor = ClientColor 'Die Schriftfarbe wieder setzten, da sie beim einfügen des bildes verlorengeht
    RtxtChat.SelText = (Left(Teil, InStr(1, Teil, ":") - 1)) 'Den ganzen Text bis zum ersten : einfügen
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, ":") + 1) 'den ganzen Text links vom : wegnehmen
  ElseIf (Doppelpunkt - Semikolon) > 0 Then
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
    RtxtChat.SelColor = ClientColor 'Die Schriftfarbe wieder setzten, da sie beim einfügen des bildes verlorengeht
    RtxtChat.SelText = (Left(Teil, InStr(1, Teil, ";") - 1)) 'Den ganzen Text bis zum ersten : einfügen
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, ";") + 1) 'den ganzen Text links vom : wegnehmen
  End If
 
  For i = 0 To 47 'geht alle möglichen Icons durch
    'die Erforderliche Zeichenfolge für ein Icon (z.b. :-) oder :nono) steht in der .Tag eigenschaft des imgIcon-Objekts
    If Left(Teil, Len(frmEmoticons!imgIcon(i).Tag)) = frmEmoticons!imgIcon(i).Tag Then 'falls eine richtige zeichenfolge eingegeben wurde
      Clipboard.Clear 'Inhalt der Zwischenablage leeren
      Clipboard.SetData frmEmoticons!imgIcon(i).Picture 'Icon in die Zwischenablage kopieren
      
      RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das Ende des Textes setzen
      RtxtChat.Locked = False 'Änderungen im Chattextfeld zulassen (wird für das Einfügen des Bildes benötigt)
      SendMessage RtxtChat.hwnd, WM_PASTE, 0, 0  'Das Bild im Chattextfeld einfügen
      RtxtChat.Locked = True 'Änderungen im Chattextfeld nicht zulassen
      
      Teil = Right(Teil, Len(Teil) - Len(frmEmoticons!imgIcon(i).Tag)) 'die :etc. Zeichenfolge wegnehmen
      GoTo Anfang 'Wieder von vorne anfangen
    End If
  Next i
  
  'wenn keine Icons dafür vorgesehen waren, dass : normal schreiben
  RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
  RtxtChat.SelText = ":"  'das : normal in den Text einfügen
  Teil = Right(Teil, Len(Teil) - 1) 'das : wegnehmen
  GoTo Anfang 'wieder von vorne anfangen

Weiter: 'es kommen keine :etc. -Zeichenfolgen (mehr) im Text vor.
  RtxtChat.SelStart = Len(RtxtChat.Text) 'Texteinfügemarke an das Ende des Textes setzen
  RtxtChat.SelColor = ClientColor 'Die Schriftfarbe wieder setzten, da sie beim einfügen des bildes verlorengeht
  RtxtChat.SelText = Teil + Chr$(13) + Chr$(10) 'den Rest des Textes einfügen
End Function
