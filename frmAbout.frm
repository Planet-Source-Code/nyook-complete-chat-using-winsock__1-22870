VERSION 5.00
Object = "{F9F5B250-E80B-11D4-95D4-E305F180C055}#1.0#0"; "PKLINK.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "About"
   ClientHeight    =   3450
   ClientLeft      =   4650
   ClientTop       =   3915
   ClientWidth     =   3135
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3135
   Begin pkLinkControl.pkLink pklIDev 
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   16
      Href            =   "http://www.inter-dev.de"
      Caption         =   "www.inter-dev.de"
      BeginProperty FontNormal {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   16
   End
   Begin pkLinkControl.pkLink pklNResidence 
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   16
      Href            =   "http://www.nukegod.ixy.de"
      Caption         =   "www.nukegod.ixy.de"
      BeginProperty FontNormal {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   16
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblIMS2 
      Caption         =   "for all-a-round testing"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblIMS 
      Alignment       =   1  'Rechts
      Caption         =   "IMS2000"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblSC2 
      Caption         =   "for good help"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Rechts
      Caption         =   "Sentcool"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblIDev 
      Caption         =   "for the FileTransfer- Example"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks to:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblBy 
      Caption         =   "by nukegod@gmx.net"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblVer 
      Caption         =   "v2.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblTitel 
      Alignment       =   2  'Zentriert
      Caption         =   "nukechat"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Show
End Sub
