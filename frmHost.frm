VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmHost 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Host"
   ClientHeight    =   3120
   ClientLeft      =   3885
   ClientTop       =   3720
   ClientWidth     =   5985
   Icon            =   "frmHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5985
   Begin VB.CheckBox chkInternetChat 
      Caption         =   "Internetchat"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame framAdvanced 
      Caption         =   "Advanced:"
      Height          =   2295
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtMaxClients 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "100"
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox cbUsedIP 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblArray 
         Caption         =   "(2 - 100)"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblMaxClients 
         Caption         =   "Max. Chatters:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblChat 
         Caption         =   "If you want to chat over the Internet, you must choose your Internet-IP."
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUsedIP 
         Caption         =   "Used IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   1320
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSMyIP 
      Left            =   960
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtNN 
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtSName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdHost 
      Caption         =   "Host!"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblInetchat 
      Caption         =   "This allows other users to see your chatroom in the chatroomlist (a list of all chatrooms in the internet)"
      Height          =   615
      Left            =   480
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblYNN 
      Alignment       =   1  'Rechts
      Caption         =   "Your Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblSName 
      Alignment       =   1  'Rechts
      Caption         =   "Chatroom Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click() 'das Host-Formular schliessen und zum Hauptfenster zurückkehren
  Unload Me 'Formular beenden
End Sub

Public Sub cmdHost_Click() 'Host-Vorgang starten
  If Trim(txtNN.Text) = "" Then  'Überprüfen, ob ein Nickname angegeben wurde
    Meldung = MsgBox("Enter your Nickname!", vbOKOnly, "Unable to host") 'wenn nicht, dann Meldung an den Benutzer...
    txtNN.SetFocus
    Exit Sub '...und die Prozedur beenden, d. h. den Host-Vorgang abbrechen
  ElseIf txtSName.Text = "" Then 'ist ein Nickname vorhanden, dann wird überprüft, ob ein Name für den Server vorhanden ist.
    Meldung = MsgBox("Enter a Servername!", vbOKOnly, "Unable to host") 'wenn nicht, Meldung an den benutzer...
    txtSName.SetFocus
    Exit Sub '...und die Prozedur beenden, d. h. den Host-Vorgang abbrechen
  End If
  
  If Val(txtMaxClients.Text) < 2 Or Val(txtMaxClients.Text) > 100 Then 'prüft, ob die Maximale Teilnehmerzahl gültig ist (der Hoster selbst ist auch mitgezählt)
    Meldung = MsgBox(txtMaxClients.Text + " is not a valid number.", vbCritical, "Unable to host") 'ist sie nicht gültig, dann wird eine Meldung an den Benutzer gegeben, um ihn darauf hinzuweisen
    txtMaxClients.SetFocus
    Exit Sub 'den Host-Vorgang abbrechen
  End If
  
  ReDim Client(txtMaxClients.Text - 1) 'maximal 100 Benutzer in einem Server
  MaxAnzClients = txtMaxClients.Text 'die Maximale anzahl an Clients festlegen
  
  NickName = FormNickName(txtNN.Text) 'Leerzeichen links und rechts entfernen und leerzeichen im namen durch _ ersetzen

  With Client(0) 'die Eigenschaften für den ersten Benutzer (in diesem falle der Host selbst) des Chatraumes festlegen.
    .FontColor = vbBlack 'Textfarbe auf schwarz (Standart)
    .IP = WSMyIP.LocalIP  'seine IP
    .NickName = NickName 'und sein Nickname
  End With
  
  ServerIP = Client(0).IP 'die ServerIP. ist in diesem fall die gleiche wie die des benuzers, da dieser ja den server hostet
  NickName = Client(0).NickName  'der Nickname des Users.
  hostname = FormHostName(txtSName.Text) 'der Name des Chatservers
  anzClients = anzClients + 1 'es ist ein Client mehr vorhanden. der hoster selbtst wird als client gehandelt.
  'auf der frmSChat-form sind die client-winsocks, sowie die server winsocks vorhanden.
  'das lässt das Ganze etwas komplizierter erscheinen
  
  anzBannedIPs = 0 'es sind noch keine gesperrten IPs vorhanden
  
  'wenn es ein internetchat werden soll, so wird nun die liste der server downgeloadet,
  'der eigene server hinzugefügt, und das ganze zeug wieder uploadet.
  If chkInternetChat.Value = 1 Then
    'Die serverliste herunterladen und den eigenen server hinzufügen
    GetNgenerateNewServerlist
    'Die neue serverliste uploaden
    UploadNewServerlist
  End If

  Load frmSChat 'dann kann das Chat-Formular für Server geladen werden
  frmHost.Visible = False
  
  frmSChat!WSScmdSnd.RemoteHost = "255.255.255.255" 'an alle im netzwerk senden
  frmSChat!WSScmdSnd.SendData ("ServerInfo/" + hostname + "/" + NickName + "/" + Trim(Str(MaxAnzClients))) 'Informationen über den Server senden: NameDesChatRaums / Hoster / MaximaleAnzahlVonTeilnehmern
  
End Sub

Private Sub cbUsedIP_Click()
  UsedIP = IP(cbUsedIP.ListIndex)
End Sub

Private Sub Form_Load() 'das Formular wird geladen...
  getIPs 'Die IPs heraussuchen:
  For i = 0 To anzIPs - 1
    cbUsedIP.AddItem IP(i)
  Next i
  UsedIP = IP(0)
  cbUsedIP.ListIndex = 0
  
  frmHost.Caption = "Host - " + WSMyIP.LocalIP
  Show '...und angezeigt
End Sub

Private Sub Form_Unload(Cancel As Integer)
  anzIPs = 0
  frmMain.Visible = True 'das Haupt-Formular wird wieder sichtbar gemacht
End Sub

Private Function FormHostName(hostname As String)  'Diese Prozedur ersetzt eventuelle / im Hostnamen durch \
Anfang: 'Hier werden / durch \ ersetzt
  If InStr(1, hostname, "/") = 0 Then 'wenn keine / mehr im Hostnamen vorkommen
    FormHostName = hostname
    Exit Function
  End If
  hostname = Left(hostname, InStr(1, hostname, "/") - 1) + "\" + Right(hostname, Len(hostname) - InStr(1, hostname, "/")) 'den ersten / von Links durch \ ersetzen
  GoTo Anfang 'wieder von vorne anfangen
End Function

Private Sub txtMaxClients_KeyPress(KeyAscii As Integer) 'Der Benutzer drückt eine Taste im Textfeld
  If KeyAscii = 13 Then 'prüft, ob Enter gedruckt wurde
    KeyAscii = 0 'den Windows-Ding unterdrücken
    cmdHost_Click 'Den Host-vorgang starten
  End If
End Sub

Private Sub txtNN_KeyPress(KeyAscii As Integer) 'Der Benutzer drückt eine Taste im Textfeld
  If KeyAscii = 13 Then 'prüft, ob Enter gedruckt wurde
    KeyAscii = 0 'den Windows-Ding unterdrücken
    txtMaxClients.SetFocus 'Das MaxClients-Textfeld aktivieren
  End If
End Sub

Private Sub txtSName_KeyPress(KeyAscii As Integer) 'Der Benutzer drückt eine Taste im Textfeld
  If KeyAscii = 13 Then 'prüft, ob Enter gedruckt wurde
    KeyAscii = 0 'den Windows-Ding unterdrücken
    txtNN.SetFocus 'Das Nickname-Textfeld aktivieren
  End If
End Sub
