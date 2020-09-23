VERSION 5.00
Object = "{F9F5B250-E80B-11D4-95D4-E305F180C055}#1.0#0"; "PKLINK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " "
   ClientHeight    =   3525
   ClientLeft      =   3690
   ClientTop       =   3720
   ClientWidth     =   2325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   2325
   Begin pkLinkControl.pkLink pklNResidence 
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   397
      BackColor       =   14737632
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
      ColorNormal     =   8421504
      BeginProperty FontNormal {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   16
   End
   Begin VB.Line lnJoin 
      X1              =   720
      X2              =   480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line lnDownExit 
      X1              =   480
      X2              =   480
      Y1              =   2520
      Y2              =   1920
   End
   Begin VB.Line lnDownJoin 
      X1              =   480
      X2              =   480
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line lnExit 
      X1              =   720
      X2              =   480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Shape shpExit 
      Height          =   495
      Left            =   720
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line lnHost 
      X1              =   480
      X2              =   720
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line lnDownHost 
      X1              =   480
      X2              =   480
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Label lblTitel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "nukechat v2.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Shape shpTitel 
      BorderColor     =   &H000000C0&
      Height          =   615
      Left            =   360
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape shpHost 
      Height          =   495
      Left            =   720
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblHost 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Host"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblJoin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Join"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape shpJoin 
      Height          =   495
      Left            =   720
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblBy 
      BackColor       =   &H00E0E0E0&
      Caption         =   "by nukegod@gmx.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'nukechat 1.5 by nukegod@gmx.net
'www.nukegod.ixy.de
'
'Dieses Programm dient dazu, einfach und bequem im LAN oder Internet zu Chatten.
'Dieses Programm (und der Source Code) kann in jeder Form weitergegeben, modifiziert
'und als Basis für ähnliche Programme verwendet werden (auch für Remote-Access
'Programme, jedoch NICHT für Trojaner-> Ich hafte für keine Schäden die durch
'modifizierten Source Code entstehen. Wenn Sie sich nicht sicher sind, ob ihre Kopie
'Trojanerfrei ist, laden sie sich am besten die aktuellste Version auf
'www.nukegod.ixy.de. Die dort zum Download bereitgestellten Programme sind garantiert
'Trojanerfrei). Im Falle einer Modifikation oder Ähnlichem bitte in den Credits eine
'Information über diesen Source Code hinterlassen (z.B. "Basiert auf dem nukechat von
'nukegod@gmx.net. www.nukegod.ixy.de) Ich würde mich auch über ein Feedback freuen.
'Für Fragen bezüglich dieses Codes oder über Visual Basic stehe ich gerne zur Verfügung.

Private Sub Form_Load()
  If App.PrevInstance = True Then Unload Me  'wenn der LAN-Chat schon einmal gestartet wurde, soll beendet werden
  Show 'Fenster anzeigen
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHost.ForeColor = vbBlack
  shpHost.BorderColor = vbBlack
  lnHost.BorderColor = vbBlack
  lnDownHost.BorderColor = vbBlack
  lblJoin.ForeColor = vbBlack
  shpJoin.BorderColor = vbBlack
  lnJoin.BorderColor = vbBlack
  lnDownJoin.BorderColor = vbBlack
  lblExit.ForeColor = vbBlack
  shpExit.BorderColor = vbBlack
  lnExit.BorderColor = vbBlack
  lnDownExit.BorderColor = vbBlack
End Sub

Private Sub lblExit_Click() 'das Programm beenden...
  Unload Me
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHost.ForeColor = vbBlack
  shpHost.BorderColor = vbBlack
  lnHost.BorderColor = vbBlack
  lnDownHost.BorderColor = &HC0&
  lblJoin.ForeColor = vbBlack
  shpJoin.BorderColor = vbBlack
  lnJoin.BorderColor = vbBlack
  lnDownJoin.BorderColor = &HC0&
  lblExit.ForeColor = &HC0&
  shpExit.BorderColor = &HC0&
  lnExit.BorderColor = &HC0&
  lnDownExit.BorderColor = &HC0&
End Sub

Private Sub lblHost_Click() 'einen Chatserver hosten....
  frmMain.Visible = False 'das Haupt-Formular wird unsichtbar gemacht
  Load frmHost 'das Host-Formular wird geladen
End Sub

Private Sub lblHost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHost.ForeColor = &HC0&
  shpHost.BorderColor = &HC0&
  lnHost.BorderColor = &HC0&
  lnDownHost.BorderColor = &HC0&
  lblJoin.ForeColor = vbBlack
  shpJoin.BorderColor = vbBlack
  lnJoin.BorderColor = vbBlack
  lnDownJoin.BorderColor = vbBlack
  lblExit.ForeColor = vbBlack
  shpExit.BorderColor = vbBlack
  lnExit.BorderColor = vbBlack
  lnDownExit.BorderColor = vbBlack
End Sub

Private Sub lblJoin_Click() 'einen bestehenden Chatserver beitreten...
  Unload Me 'das Haupt-Formular wird beendet
  Load frmJoin 'das Host-Formular wird geladen
End Sub

Private Sub lblJoin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHost.ForeColor = vbBlack
  shpHost.BorderColor = vbBlack
  lnHost.BorderColor = vbBlack
  lnDownHost.BorderColor = &HC0&
  lblJoin.ForeColor = &HC0&
  shpJoin.BorderColor = &HC0&
  lnJoin.BorderColor = &HC0&
  lnDownJoin.BorderColor = &HC0&
  lblExit.ForeColor = vbBlack
  shpExit.BorderColor = vbBlack
  lnExit.BorderColor = vbBlack
  lnDownExit.BorderColor = vbBlack
End Sub
