VERSION 5.00
Begin VB.Form frmEmoticons 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Emoticons"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmEmoticons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   47
      Left            =   1200
      Picture         =   "frmEmoticons.frx":0442
      Tag             =   ":uriel"
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label Label48 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":uriel"
      Height          =   255
      Left            =   1560
      TabIndex        =   47
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label47 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":ugly"
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label46 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":wow"
      Height          =   255
      Left            =   3720
      TabIndex        =   45
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":drink"
      Height          =   255
      Left            =   1800
      TabIndex        =   44
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":satan"
      Height          =   255
      Left            =   5040
      TabIndex        =   43
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":elk"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":clown"
      Height          =   255
      Left            =   3840
      TabIndex        =   41
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":cat"
      Height          =   255
      Left            =   2760
      TabIndex        =   40
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label40 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":bear"
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":evil"
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   46
      Left            =   2400
      Picture         =   "frmEmoticons.frx":081F
      Tag             =   ":bear"
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   45
      Left            =   4680
      Picture         =   "frmEmoticons.frx":0BAF
      Tag             =   ":satan"
      Top             =   960
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   44
      Left            =   3480
      Picture         =   "frmEmoticons.frx":0F90
      Tag             =   ":wow"
      Top             =   2880
      Width           =   225
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   43
      Left            =   1200
      Picture         =   "frmEmoticons.frx":1318
      Tag             =   ":drink"
      Top             =   3120
      Width           =   570
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   42
      Left            =   3480
      Picture         =   "frmEmoticons.frx":1723
      Tag             =   ":evil"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   270
      Index           =   41
      Left            =   2400
      Picture         =   "frmEmoticons.frx":1AB0
      Tag             =   ":cat"
      Top             =   2760
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   420
      Index           =   40
      Left            =   4680
      Picture         =   "frmEmoticons.frx":1E58
      Tag             =   ":elk"
      Top             =   480
      Width           =   450
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   39
      Left            =   3480
      Picture         =   "frmEmoticons.frx":2351
      Tag             =   ":clown"
      Top             =   2160
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   38
      Left            =   4680
      Picture         =   "frmEmoticons.frx":2752
      Tag             =   ":ugly"
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":blah"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":cry"
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":arnie"
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":beat"
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":fuckyou"
      Height          =   255
      Left            =   3960
      TabIndex        =   33
      Top             =   1800
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   12
      Left            =   120
      Picture         =   "frmEmoticons.frx":2ADA
      Tag             =   ":'("
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   9
      Left            =   120
      Picture         =   "frmEmoticons.frx":2E73
      Tag             =   ":cool"
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   255
      Index           =   11
      Left            =   120
      Picture         =   "frmEmoticons.frx":31F6
      Tag             =   ":baby"
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   8
      Left            =   120
      Picture         =   "frmEmoticons.frx":3596
      Tag             =   ":smoke"
      Top             =   2160
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   3
      Left            =   120
      Picture         =   "frmEmoticons.frx":3931
      Tag             =   ":p"
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   7
      Left            =   120
      Picture         =   "frmEmoticons.frx":3CB8
      Tag             =   ":?"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   6
      Left            =   120
      Picture         =   "frmEmoticons.frx":404D
      Tag             =   ":sleep"
      Top             =   1560
      Width           =   420
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   5
      Left            =   120
      Picture         =   "frmEmoticons.frx":43E1
      Tag             =   ":grr"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   4
      Left            =   120
      Picture         =   "frmEmoticons.frx":4771
      Tag             =   ":("
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   2
      Left            =   120
      Picture         =   "frmEmoticons.frx":4AEF
      Tag             =   ":D"
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   1
      Left            =   120
      Picture         =   "frmEmoticons.frx":4E77
      Tag             =   ";)"
      Top             =   360
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   0
      Left            =   120
      Picture         =   "frmEmoticons.frx":51F9
      Tag             =   ":)"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   10
      Left            =   120
      Picture         =   "frmEmoticons.frx":5576
      Tag             =   ":nono"
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":)"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   ";)"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":D"
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":p"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":("
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":grr"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":sleep"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":?"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":smoke"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":cool"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":nono"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":baby"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":'("
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":shoot"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   375
      Index           =   35
      Left            =   3480
      Picture         =   "frmEmoticons.frx":5958
      Tag             =   ":toiletclaw"
      Top             =   120
      Width           =   705
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   29
      Left            =   2400
      Picture         =   "frmEmoticons.frx":5D83
      Tag             =   ":wave"
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   18
      Left            =   1200
      Picture         =   "frmEmoticons.frx":611B
      Tag             =   ":heart"
      Top             =   840
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   14
      Left            =   3480
      Picture         =   "frmEmoticons.frx":649E
      Tag             =   ":arnie"
      Top             =   1320
      Width           =   795
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Index           =   32
      Left            =   2400
      Picture         =   "frmEmoticons.frx":6905
      Tag             =   ":ass"
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   28
      Left            =   2400
      Picture         =   "frmEmoticons.frx":6CCA
      Tag             =   ":bala"
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   26
      Left            =   3480
      Picture         =   "frmEmoticons.frx":7049
      Tag             =   ":cry"
      Top             =   3120
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   17
      Left            =   1200
      Picture         =   "frmEmoticons.frx":73D9
      Tag             =   ":devil"
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   19
      Left            =   1200
      Picture         =   "frmEmoticons.frx":77A1
      Tag             =   ":erm"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   390
      Index           =   33
      Left            =   2400
      Picture         =   "frmEmoticons.frx":7B23
      Tag             =   ":flush"
      Top             =   1920
      Width           =   390
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   15
      Left            =   3480
      Picture         =   "frmEmoticons.frx":7F01
      Tag             =   ":fuckyou"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   31
      Left            =   2400
      Picture         =   "frmEmoticons.frx":82B9
      Tag             =   ":guitar"
      Top             =   1200
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   20
      Left            =   1200
      Picture         =   "frmEmoticons.frx":867F
      Tag             =   ":tilt"
      Top             =   1320
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   23
      Left            =   1200
      Picture         =   "frmEmoticons.frx":8A0D
      Tag             =   ":gg"
      Top             =   2160
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   34
      Left            =   2400
      Picture         =   "frmEmoticons.frx":8D99
      Tag             =   ":guns"
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   30
      Left            =   2400
      Picture         =   "frmEmoticons.frx":91A4
      Tag             =   ":light"
      Top             =   960
      Width           =   225
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   25
      Left            =   3480
      Picture         =   "frmEmoticons.frx":953D
      Tag             =   ":beat"
      Top             =   960
      Width           =   525
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   36
      Left            =   3480
      Picture         =   "frmEmoticons.frx":993F
      Tag             =   ":blah"
      Top             =   600
      Width           =   660
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   21
      Left            =   1200
      Picture         =   "frmEmoticons.frx":9D0D
      Tag             =   ":lol"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   22
      Left            =   1200
      Picture         =   "frmEmoticons.frx":A091
      Tag             =   ":mad"
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   24
      Left            =   1200
      Picture         =   "frmEmoticons.frx":A417
      Tag             =   ":shit"
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   27
      Left            =   2400
      Picture         =   "frmEmoticons.frx":A78D
      Tag             =   ":king"
      Top             =   120
      Width           =   345
   End
   Begin VB.Image imgIcon 
      Height          =   315
      Index           =   37
      Left            =   4320
      Picture         =   "frmEmoticons.frx":AB41
      Tag             =   ":angel"
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image imgIcon 
      Height          =   375
      Index           =   16
      Left            =   1200
      Picture         =   "frmEmoticons.frx":AF93
      Tag             =   ":alien"
      Top             =   120
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Index           =   13
      Left            =   4320
      Picture         =   "frmEmoticons.frx":B33E
      Tag             =   ":shoot"
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":shit"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":gg"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":mad"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":lol"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":flush"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":erm"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":angel"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":devil"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":alien"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":tilt"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":heart"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":king"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":bala"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":wave"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":ass"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":light"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":guitar"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":guns"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":toiletclaw"
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmEmoticons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
