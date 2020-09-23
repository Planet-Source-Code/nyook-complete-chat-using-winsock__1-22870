VERSION 5.00
Begin VB.Form frmSpecify 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Specify"
   ClientHeight    =   810
   ClientLeft      =   4680
   ClientTop       =   3750
   ClientWidth     =   2625
   Icon            =   "frmSpecify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   2625
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblIP 
      Caption         =   "Chatroom - IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSpecify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click() 'der Benutzer möchte zum Join Fenster zurückkehren
  Unload Me
End Sub

Private Sub cmdSearch_Click() 'der Benutzer möchte einen Chatraum unter der eingegebenen IP suchen
  On Error Resume Next 'bei einem Fehler weitermachen
  frmJoin!WSSrequestSnd.RemoteHost = Trim(txtIP.Text) 'auf die eingegebene IP ausrichtsen
  frmJoin!WSSrequestSnd.SendData "RequestInfo"  'Anfrage auf Serverinfos senden
  Unload Me
End Sub

Private Sub Form_Load() 'beim Laden des Fensters
  Show 'Fenster anzeigen
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Routine beim Beenden des Specify-Fensters
  frmJoin.Enabled = True 'das Join-Fenster wieder aktivieren
End Sub
