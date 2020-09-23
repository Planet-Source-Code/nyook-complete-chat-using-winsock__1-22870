VERSION 5.00
Begin VB.Form frmServerOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Server Options"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frmServerOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtBannMinutes 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   "5"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkMaxMsgL 
      Caption         =   "Max. text lenght:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Value           =   1  'Aktiviert
      Width           =   1500
   End
   Begin VB.TextBox txtMaxMsgL 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "700"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Frame framAutoKickOpts 
      Caption         =   "Autokick Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtFIms 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   "400"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtMaxMsg 
         Height          =   280
         Left            =   1680
         TabIndex        =   2
         Text            =   "4"
         Top             =   360
         Width           =   200
      End
      Begin VB.CheckBox chkFlood 
         Caption         =   "Anti Flood:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Aktiviert
         Width           =   1095
      End
      Begin VB.Label lblMS 
         Caption         =   "ms"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblInterval 
         Caption         =   "Flood-Interval of:"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblMax 
         Caption         =   "Max."
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblMsg 
         Caption         =   "Messages at a"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label lblBannMinutes 
      Caption         =   "Bann-minutes:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmServerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFlood_Click() 'Der Benutzer möchte die AutokickFunkion einschalten oder ausschalten
  If chkFlood.Value = 0 Then 'Das Häckchen wird weggenommen
    lblMax.Enabled = False
    txtMaxMsg.Enabled = False
    lblMsg.Enabled = False
    lblInterval.Enabled = False
    txtFIms.Enabled = False
    lblMS.Enabled = False
  Else 'Das Häckchen wird gesetzt
    lblMax.Enabled = True
    txtMaxMsg.Enabled = True
    lblMsg.Enabled = True
    lblInterval.Enabled = True
    txtFIms.Enabled = True
    lblMS.Enabled = True
  End If
End Sub

Private Sub chkMaxMsgL_Click() 'Der Benutzer möchte die Maximale Nachrichtlänge aktivieren bzw. deaktivieren
  Select Case chkMaxMsgL.Value
    Case 0 'Wenn das Häckchen weggenommen wurde
      'die Objekte deaktivieren...
      txtMaxMsgL.Enabled = False
    Case 1 'Wenn das Häckchen gesetzt wurde
      'die Objekte aktivieren...
      txtMaxMsgL.Enabled = True
  End Select
End Sub

Private Sub cmdClose_Click()
  Me.Visible = False
End Sub
