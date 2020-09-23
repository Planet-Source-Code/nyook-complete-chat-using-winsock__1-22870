VERSION 5.00
Begin VB.Form frmCommands 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Chatcommands"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmCommands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame framSpecial 
      Caption         =   "Special Commands (for Server)"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4335
      Begin VB.Label lblBann2 
         Caption         =   "Banns a Client for a specified time"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblBann 
         Caption         =   "/bann Client"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblKick2 
         Caption         =   "Kicks a Client "
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblKick 
         Caption         =   "/kick Client"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label lblExit2 
      Caption         =   "Leaves the Chat"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblExit 
      Caption         =   "/exit"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblKeyUpDown2 
      Caption         =   "Switches trough all sent messages"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblKeyUpDown 
      Caption         =   "Key-Up , Key-Down"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblClear2 
      Caption         =   "Deletes the content of the Chatbox"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblClear 
      Caption         =   "/clear"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblListClients2 
      Caption         =   "Generates a list of all Clients"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblListClients 
      Caption         =   "/listclients"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblW2 
      Caption         =   "You whisper to another Client"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblW 
      Caption         =   "/w Client Message"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblMe2 
      Caption         =   "Thirdperson-message"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblMe 
      Caption         =   "/me Message"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Show
End Sub
