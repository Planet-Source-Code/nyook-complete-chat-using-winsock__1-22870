VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSend 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Ready"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   960
      Width           =   3015
   End
   Begin VB.Timer timerSpeed 
      Interval        =   1000
      Left            =   2640
      Top             =   1200
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Open"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   2880
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar FileBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog ComDiag 
      Left            =   2400
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSelectedFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected File:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblFileSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Filesize: 0 kb"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblComplete 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete: 0%"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Speed: 0.0 / KBps"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Basiert auf dem OriginalCode von:
' Ronny R. Germany Berlin
' Contact me: manager@directbox.com

Dim UserClose As Boolean
Dim DoneBytes As Long 'Die Anzahl versendeter Bytes in der Sekunde
Dim TotalBytes As Long
Dim NextPart As Boolean
Dim FileSize As Long

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdOpen_Click()
  On Error GoTo Quit
  
  ComDiag.ShowOpen

  txtFile.Text = ComDiag.FileName
  lblFileSize.Caption = "Filesize: " & FileLen(ComDiag.FileName)
  lblFileName.Caption = "Filename: " & ComDiag.FileTitle

Quit:
End Sub

Private Sub cmdSend_Click()
  On Error GoTo ErrorHandler:
  Dim StartTime As Long
  Dim ClientIP As String
  TotalBytes = 0

  'the following routines are nessessary to beware of errors
  If wsSend.State <> sckClosed Then wsSend.Close '# Reset if winsock was in use
        
  '# Init the Winsock
  wsSend.RemotePort = FTPort '# set the winsock send remoteport; on the same port the client should listen already
  
  'sucht die IP des Clients raus...
  For i = 0 To (anzClients - 1)
    If Client(i).NickName = FileTransferClient Then ClientIP = Client(i).IP
  Next i
  wsSend.RemoteHost = ClientIP '# that should be the same ip the client uses (Local 127.0.0.1)
        
  wsSend.Connect '# connecting to port
  DoEvents

  StartTime = Timer
  Do While wsSend.State <> 7 And Timer - StartTime < 30
    DoEvents '# Wait until the connections ethablishes
    'prüfen, ob der benutzer den FileTransfer abbrechen möchte
    If UserClose = True Then Exit Sub
  Loop       '  there must be a timeout check else it will never end
  If Timer - StartTime > 30 Then GoTo Timeout '# When Timeout
       
  
  '-----------------------------------------------------
  '# Now we come to the send routine
  '# You have to open a file in binary mode, read out 2k packages and send them to the connected port
  '# Letz start
        
  Dim OpenedFileNbr, Back
  Dim Temp As String
  Dim PackageSize As Long
  Dim LastData As Boolean
            
  FileSize = FileLen(txtFile.Text)
  FileBar.Value = 0
            
  wsSend.SendData ("FILEINFO|" & FileSize & "|" & ComDiag.FileTitle & "|") '# You can add more like filename , description ...
            
  StartTime = Timer
  Do While NextPart = False And Timer - StartTime < 30 '# When the next Package where not send the procedure will quit after 30 secs timeout
    DoEvents
    'prüfen, ob der benutzer den FileTransfer abbrechen möchte
    If UserClose = True Then Exit Sub
  Loop
  If Timer - StartTime > 30 Then GoTo Timeout '# When Timeout
                        
  PackageSize = 2048  '#  Declare the size of the packages to send
  LastData = False '#  You'll see that we need that to make the received
                   '   file excactly the same size like the original one
  NextPart = True  '#  NextPart is a form-global variable which
                   '   contains wheter the package was send or not
                   '   take a look at the winsock_sendcomplete event
  OpenedFileNbr = FreeFile '# Find a free Filenumber to open your file
  Open txtFile.Text For Binary Access Read As OpenedFileNbr
                        
  Temp = ""
  Do Until EOF(OpenedFileNbr)
    ' Adjust PackageSize at end so we don't read too much data
    If FileSize - Loc(OpenedFileNbr) <= PackageSize Then
      PackageSize = FileSize - Loc(OpenedFileNbr) + 1
      LastData = True
    End If
                            
    Temp = Space$(PackageSize) '# Make string empty for data
    Get OpenedFileNbr, , Temp '# Load data into string
                            
    If wsSend.State <> 7 Then Exit Sub '# Checks again wether the connections exist or not
    On Error Resume Next
                            
    StartTime = Timer
    Do While NextPart = False And Timer - StartTime < 30 '# When the next Package where not send the procedure will quit after 30 secs timeout
      DoEvents
      'prüfen, ob der benutzer den FileTransfer abbrechen möchte
      If UserClose = True Then Exit Sub
    Loop
    If Timer - StartTime > 30 Then GoTo Timeout '# When Timeout
                            
    If wsSend.State = 7 Then '# Check state again
                            
      If LastData = True Then Temp = Mid(Temp, 1, Len(Temp) - 1) '# We added one byte above, which we don't wanna send
                                                                 '   therefore we need lastdata
      DoneBytes = DoneBytes + Len(Temp)
      TotalBytes = TotalBytes + Len(Temp)
      FileBar.Value = TotalBytes * 100 / FileSize
      lblComplete.Caption = "Complete: " & Int(TotalBytes * 100 / FileSize) & " %"
      wsSend.SendData Temp '# Send datapackage
      NextPart = False '# Set the senddata check
     Else
      Exit Sub
    End If
  Loop

  Close #OpenedFileNbr '# Last package was send, now you can close the file
                            
  Do While NextPart = False '# You have to wait until the sendprogress is done because
    'prüfen, ob der benutzer den FileTransfer abbrechen möchte
    If UserClose = True Then Exit Sub
    DoEvents                '   when we close the winsock before the file was send completly
  Loop                      '   data will be lost --> We use the close event in the client to
                            '   close the received file too
  
  NextPart = False
  
  wsSend.SendData "EOF"
  DoEvents
  
  Do While NextPart = False '# You have to wait until the sendprogress is done
    DoEvents
    'prüfen, ob der benutzer den FileTransfer abbrechen möchte
    If UserClose = True Then Exit Sub
  Loop
    
  DoneBytes = 0
  lblSpeed.Caption = "Speed: 0.0 / KBps"
  
  wsSend.Close
  frmSend.Caption = "Ready"
  Exit Sub

Timeout:
  MsgBox "Timeout" '# write what you want to say to the user
  Exit Sub
        
ErrorHandler:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
  frmSend.Caption = "Sending File to: " + FileTransferClient
  frmSend.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UserClose = True
End Sub

Private Sub timerSpeed_Timer()
  lblSpeed.Caption = "Speed: " & Format(DoneBytes / 1000, "###0.0") & " / KBps"
  DoneBytes = 0
End Sub

Private Sub wsSend_SendComplete()
  NextPart = True
End Sub
