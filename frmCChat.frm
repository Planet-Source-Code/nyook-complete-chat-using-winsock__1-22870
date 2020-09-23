VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCChat 
   Caption         =   "nukechat"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7260
   Icon            =   "frmCChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picSTIconCTRL 
      Height          =   615
      Left            =   6840
      Picture         =   "frmCChat.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer timerBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   7560
      Top             =   6480
   End
   Begin VB.Frame framClients 
      Caption         =   "Clients"
      Height          =   5295
      Left            =   5280
      TabIndex        =   1
      Top             =   -10
      Width           =   1935
      Begin VB.ListBox lstClients 
         Height          =   4935
         ItemData        =   "frmCChat.frx":0884
         Left            =   120
         List            =   "frmCChat.frx":0886
         TabIndex        =   2
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
   Begin MSComDlg.CommonDialog ComDiag 
      Left            =   1560
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WScmdRcv 
      Left            =   840
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WScmdSnd 
      Left            =   240
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSmessageRcv 
      Left            =   840
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock WSmessageSnd 
      Left            =   240
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin RichTextLib.RichTextBox RtxtChat 
      Height          =   5025
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8864
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCChat.frx":0888
   End
   Begin VB.Image imgSTIcon 
      Height          =   480
      Index           =   0
      Left            =   8040
      Picture         =   "frmCChat.frx":0951
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSTIcon 
      Height          =   480
      Index           =   1
      Left            =   8520
      Picture         =   "frmCChat.frx":0D93
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuExitRoom 
      Caption         =   "&Exit Room"
   End
   Begin VB.Menu mnuChatT 
      Caption         =   "&Chattext"
      Begin VB.Menu mnuChatTClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuChatTCSave 
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
      Caption         =   "&PopUpMenu"
      Visible         =   0   'False
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
Attribute VB_Name = "frmCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub mnuChatTClear_Click() 'Den Inhalt des Textfeldes löschen
  RtxtChat.Text = ""
End Sub

Private Sub mnuChatTCSave_Click() 'Den Chattext speichern
  ComDiag.ShowSave 'Das Speichernformular anzeigen
  RtxtChat.SaveFile ComDiag.FileName, rtfRTF 'Den Chattext in der angegebenen Datei speichern
End Sub

Private Sub mnuExitRoom_Click() 'den Chat beenden
  Unload Me
End Sub

Private Sub mnuFTReceive_Click() 'Für das empfangen einer Datei vorbereiten
  Load frmReceive
End Sub

Private Sub mnuHelpAbout_Click() 'Der benutzer möchte Informationen über den nukechat
  Load frmAbout
End Sub

Private Sub mnuHelpCommands_Click() 'Lädt das Chatcommands-Formular, wo die Chatcommands verzeichnet sind
  Load frmCommands
End Sub

Private Sub mnuHelpEmoticons_Click() 'Lädt das Emoticons-Formular, wo die verschiedenen Emoticons beschrieben sind
  Load frmEmoticons
  frmEmoticons.Show
End Sub

Private Sub mnuOptionsColor_Click() 'der Benutzer will seine Schriftfarbe wählen
  ComDiag.ShowColor 'das Standart-Forumlar für die Farbe wird angezeigt
  'dem Server die neue Schriftfarbe mitteilen
  WScmdSnd.SendData ("NewColor/" + NickName + "/" + Str(ComDiag.Color))
End Sub

Private Sub mnuOptionsSTIBlink_Click()
  'setzt bzw. enfternt das Häckchen
  If mnuOptionsSTIBlink.Checked = True Then
    mnuOptionsSTIBlink.Checked = False
    Exit Sub
  End If
  mnuOptionsSTIBlink.Checked = True
End Sub

Private Sub mnuOptionsSTIMinimize_Click() 'Der Benutzer möchte den Chat zu einem
                                          'Systrayicon minimieren
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
  frmCChat.Visible = False 'Das Chatfenster unsichtbar machen
End Sub

Private Sub mnuPopUpSend_Click() 'Der Benutzer will jemandem eine Datei schicken
  'vermerken, an wen die Datei gesendet werden soll
  FileTransferClient = lstClients.List(lstClients.ListIndex)
  Load frmSend
End Sub

Private Sub mnuViewClientlist_Click() 'Die Clientliste anzeigen bzw. nicht anzeigen
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

Private Sub Form_Load()
  'die Winsock-Controls werden gebunden, das heisst die benötigten eigenschaften der
  'Controls werden zugewiesen, wie z.B. an welchem Port wsscmdrcv auf ein Kommando vom
  'Server warten muss
  WScmdSnd.RemoteHost = ServerIP
  WScmdSnd.RemotePort = cmd2serverPort
  WScmdRcv.Bind cmd2clientPort, UsedIP
  WSmessageSnd.RemoteHost = ServerIP
  WSmessageSnd.RemotePort = message2serverPort
  WSmessageRcv.Bind message2clientPort, UsedIP
  
  WScmdSnd.RemoteHost = ServerIP
  
  'den Namen des Chatraumes in der Titelleiste anzeigen
  frmCChat.Caption = "nukechat   -   " + hostname
  
  'die Clients im Listenelement anzeigen
  For i = 0 To (anzClients - 1)
    lstClients.AddItem Client(i).NickName
  Next i

  Show 'das Formular anzeigen
  
  aktMessage = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next 'damit es nicht zu einem absturz kommt
  
  'Meldung an den Chatserver, dass man den Raum verlässt
  WScmdSnd.SendData "LeaveChat/" + NickName
  
  'die Verbindungen werden geschlossen
  WScmdSnd.Close
  WScmdRcv.Close
  WSmessageSnd.Close
  WSmessageRcv.Close
  
  DeleteClientInformations 'die Clientinformationen werden gelöscht
  
  'eventuell geöffnete Fenster schliessen
  Unload frmEmoticons
  Unload frmAbout
  Unload frmCommands
  Unload frmSend
  Unload frmReceive

  'frmJoin.Visible = True 'das Join-Forumlar wieder sichtbar machen
  Load frmJoin
End Sub

Private Sub lstClients_Click()
  'wenn der Benuzter auf seinen eigenen Nicknamen geklickt hat, das PopUp-Menü nicht
  'anzeigen
  If NickName = lstClients.List(lstClients.ListIndex) Then Exit Sub
  PopupMenu mnuPopUp
End Sub

Private Sub mnuPopUpWhisper_Click()
  'prüft, ob der Benutzer seinen eingen NickNamen angeklickt hat
  If Client(lstClients.ListIndex).NickName = NickName Then
    Exit Sub 'abbrechen
  End If
  
  'den erforderlichen Text im Nachricht-Textfeld einfügen
  txtMessage.Text = ("/w " + Client(lstClients.ListIndex).NickName + " ")
  
  'Den Fokus auf das Texteingabefeld setzen, damit der Benutzer gleich schreiben kann
  txtMessage.SetFocus
End Sub

Private Sub mnuSysTrayPopUpOpen_Click()
  'Der Benutzer klickt auf Öffnen im SysTrayIcon-PopUpMenü
  
  'Den Timer deaktivieren, da nun kein Blinken mehr benötigt wird
  timerBlink.Enabled = False
  Shell_NotifyIcon NIM_DELETE, LanChatIcon 'Das SysTrayIcon wegnehmen
  frmCChat.Visible = True 'das Chatfenster wieder Sichtbar machen
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
  If KeyCode = vbKeyUp Then 'wenn Nach-Oben gedrückt wurde:
    'die vorherige nachricht aus der Nachrichtendatenbank anzeigen
    If (aktMessage - 1) < 1 Then Exit Sub
    aktMessage = aktMessage - 1
    txtMessage.Text = MessageDB(aktMessage)
    Exit Sub
  End If

  If KeyCode = vbKeyDown Then 'wenn Nach-Unten gedrückt wurde:
    'die nächste Nachricht aus der Nachrichtendatenbank anzeigen
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

Private Sub txtMessage_KeyPress(KeyAscii As Integer) 'der benutzer drückt eine taste
  Dim Teil As String 'diese Variable dient als eine Art Zwischenspeicher
  Dim WhisperClient As String 'der Client, zu dem eventuell geflüstert wird.
  Dim WhisperMessage As String 'die Nachricht, die geflüstert werden soll
  Dim WhisperClientIP As String 'die Ip des Clienten, zu dem geflüstert werden soll
  Dim ClientNr As Integer 'Die Nummer eines Clients
  
  If KeyAscii = "13" Then 'wenn Enter gedrückt wurde:
    'wenn im Textfeld kein Text oder nur Leerzeichen vorhanden sind
    If Trim(txtMessage.Text) = "" Then Exit Sub
    
    'die Nachricht der Datenbank hinzufügen
    anzMessages = anzMessages + 1
    ReDim Preserve MessageDB(anzMessages)
    MessageDB(anzMessages) = txtMessage.Text
    aktMessage = anzMessages + 1

    'prüft, ob der Benutzer flüstern möchte
    If Left(LTrim(txtMessage.Text), 3) = "/w " Then
      'die Nachricht ohne das "/w " speichern
      Teil = Trim(Right(Trim(txtMessage.Text), Len(Trim(txtMessage.Text)) - 3))
      'prüft, ob nur "/w" eingegeben wurde. wenn ja, dann die nachricht normal senden
      If Teil = "" Then GoTo NormalNachricht
      
      WhisperClient = Trim(Left(Teil, InStr(1, Teil, " "))) 'Liest den Client raus, an den die Nachricht gesendet wird.
      If WhisperClient = "" Then GoTo NormalNachricht 'wenn der benutzer nur "/W CLIENTNAME" eingegeben hat
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
      If WhisperClientIP = "" Then GoTo NormalNachricht 'wenn der WhisperClient gar nicht vorhanden ist
      
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
      weissnicht = FilterText(WhisperMessage, Client(ClientNr).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
      RtxtChat.SelItalic = False 'Kursivschrift deaktivieren
      RtxtChat.SelStart = Len(RtxtChat.Text)

      txtMessage.Text = "" 'der Inhalt des Textfeldes wird geleert
      KeyAscii = "0" 'Die gedrückte Taste löschen, um den Windows "Ding" zu unterdrücken
      Exit Sub
    End If
    
    If Left(Trim(txtMessage.Text), 4) = "/me " Then 'prüft, ob der Benutzer in der dritten Person sprechen will
      Teil = Trim(txtMessage.Text) 'Leerzeichen links und rechts entfernen
      Teil = Trim(Right(Teil, Len(Teil) - 3)) 'die Mitteilung in der dritten Person
      If Teil = "" Then GoTo NormalNachricht 'falls keine 3.Person-Nachricht eingegeben wurde, den Text als normale Nachricht senden
      WScmdSnd.SendData ("ThrdP/" + NickName + "/" + Teil) 'das 3.Person-Kommando mit der 3.Person-Nachricht zusammen an den Server senden
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

Private Sub WScmdRcv_DataArrival(ByVal BytesTotal As Long) 'es wird ein Cmd erhalten
  Dim Nachricht As String 'das Kommando in der Variablen Nachricht speichern
  Dim Teil As String 'Zwischenspeicher
  Dim cmdClient As String '(bei Whisper: der Client, der flüstert, bei ThrdP: der Client, der eine Flüster-Nachricht gesendet hat
  Dim WhisperMessage As String 'die Nachicht, die geflüstert wird
  Dim ClientNr As String 'Die Clientnummer (wird bei whisper oder ThrdP verwendet
  Dim ThrdPClient As String 'der Client, der eine ThrdP-Nachricht gesendet hat
  WScmdRcv.GetData Nachricht
  
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
    weissnicht = FilterText(WhisperMessage, Client(ClientNr).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
    RtxtChat.SelItalic = False 'Kursivschrift deaktivieren
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das ende setzen
    'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
    'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
    If frmCChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
      timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
    End If
    Exit Sub
  End If
  
  If Left(Nachricht, 7) = "BrdCst/" Then 'der Server hat einen Rundruf gesendet.
    RtxtChat.SelStart = Len(RtxtChat.Text) 'texteinfügemarke ans ende des Textes setzen
    RtxtChat.SelColor = vbBlack 'schriftfarbe schwart wählen
    RtxtChat.SelText = Chr$(13) + Chr$(10) + Right(Nachricht, Len(Nachricht) - 7) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) 'den Rundruf anzeigen
    RtxtChat.SelStart = Len(RtxtChat.Text)  'die Texteinfügemarke an das ende des Textes setzen. das hat zur folge, dass das text-element automatisch nach unten scrollt
    Exit Sub
  End If
  
  If Left(Nachricht, 12) = "ServerClosed" Then 'der Server wurde geschlossen
    Unload Me  'den Chat beenden
  End If
  
  If Left(Nachricht, 14) = "NewClientList/" Then 'es gibt eine neue clientlist
    Teil = Right(Nachricht, Len(Nachricht) - 14)
    FilterClientList (Teil) 'die daten aus dem "teil" für die Clientlist rausfiltern
    
    'nun wird das listenelement mit den Clients aktualisiert
    lstClients.Clear
    For i = 0 To (anzClients - 1) '...und wieder eingelesen
      lstClients.AddItem Client(i).NickName
    Next i
  End If
  
  If Left(Nachricht, 6) = "ThrdP/" Then 'ein Benutzer sendet eine 3.Person-Nachricht
    Teil = Right(Nachricht, Len(Nachricht) - 6) 'das "ThrdP/" -rausschneiden-
    cmdClient = Left(Teil, InStr(1, Teil, "/") - 1) 'der Nickname des Clients, der die 3.Person-Nachricht sendet.
    Teil = Right(Teil, Len(Teil) - Len(cmdClient) - 1) 'die 3.Person-Nachricht
    
    For i = 0 To (anzClients - 1) 'sucht die IP des cmdClients
      If Client(i).NickName = cmdClient Then
        ClientNr = i
        Exit For
      End If
    Next i
    
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke einfügen
    RtxtChat.SelColor = Client(ClientNr).FontColor 'auf die Schriftfarbe des Clients setzen
    RtxtChat.SelItalic = True 'auf kursivschrift setzen
    RtxtChat.SelText = (cmdClient + " ") 'die 3.Person-Nachricht anzeigen
    weissnicht = FilterText(Teil, Client(ClientNr).FontColor) 'die Nachricht im Chatfenster anzeigen, mit eventuellen Emoticons
    RtxtChat.SelItalic = False 'wieder auf normalschrift setzen
    RtxtChat.SelStart = Len(RtxtChat.Text) 'Die Texteinfügemarke an das ende setzen
    'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
    'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
    If frmCChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
      timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
    End If
    Exit Sub
  End If
  
  If Nachricht = "Ping" Then 'der Server prüft die Empfangsbereitschaft
    'eine Antwort an den Server schicken
    WScmdSnd.RemoteHost = ServerIP
    WScmdSnd.SendData "Pong"
    Exit Sub
  End If
End Sub

Private Sub WSmessageRcv_DataArrival(ByVal BytesTotal As Long) 'eine Nachricht VOM Server ist angekommen
  Dim Nachricht As String
  WSmessageRcv.GetData Nachricht
  For i = 0 To (anzClients - 1) 'Es wird geprüft, von wem die Nachricht ist.
    If Left(Nachricht, Len(Client(i).NickName) + 1) = (Client(i).NickName + "/") Then 'wenn man weiss, von welchem Client die Nachricht ist
      RtxtChat.SelStart = Len(RtxtChat.Text) 'die Einfügemarke wird an das Ende des Textes gesetzt
      RtxtChat.SelColor = Client(i).FontColor 'die Schriftfarbe wird gesetzt
      RtxtChat.SelText = (Client(i).NickName + ": ") 'Den Namen und einen : davor
      weissnicht = FilterText(Right(Nachricht, Len(Nachricht) - Len(Client(i).NickName) - 1), Client(i).FontColor) 'die Nachricht im Chatfenster anzeigen , mit eventuellen Emoticons
      'wenn das Fenster unsichtbar ist (weil der chat zu einem systrayicon minimiert
      'wurde) und das SysTrayIcon blinken soll, den Timer fürs Blinken aktivieren
      If frmCChat.Visible = False And mnuOptionsSTIBlink.Checked = True Then
        timerBlink.Enabled = True 'den Timer aktivieren, um das Symbol zum blinken zu bringen
      End If
    End If
  Next i
End Sub

Public Function FilterText(FilterMsg As String, ClientColor As String)  'Die Nachricht wird gefiltert, d.h. Smileys und anderes wird durch Grafiken ersetzt
  Dim Semikolon As Integer 'Position des ersten Semikolons
  Dim Doppelpunkt As Integer 'Position des ersten Doppelpunkts
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
