VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReceive 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Receive"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "frmReceive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer timerSpeed 
      Interval        =   1000
      Left            =   1920
      Top             =   240
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin MSWinsockLib.Winsock wsReceive 
      Left            =   2520
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar FileBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Speed: 0.0 / KBps"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblFileSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Filesize: 0 kb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblComplete 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete: 0%"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Basiert auf dem Originalcode von:
' Ronny R. Germany Berlin
' Contact me: manager@directbox.com

Dim DoneBytes As Long  '# for calculating kbps
Dim TotalBytes As Long
Dim DownloadingFile As Integer '# freefile for open files
Dim BytesTotal As Long
Dim FileName As String
Dim FileSize As Long

Public Function GetField(Field As String, FieldPos As Long) As String
'# That 's an routine to get elements from a string
Dim FieldCounter As Long
Dim IPPositionStart As Long
Dim IPPositionEnde As Long
Dim TempPosition As Long
Dim OpenedID As String
    
  TempPosition = 1
    
  For FieldCounter = 1 To FieldPos - 1 Step 1
    IPPositionStart = InStr(TempPosition, Field, "|", vbTextCompare)
    TempPosition = IPPositionStart + 1
  Next FieldCounter
  
  IPPositionStart = IPPositionStart + 1
  IPPositionEnde = InStr(IPPositionStart, Field, "|", vbTextCompare)
  
  On Error Resume Next
  If IPPositionEnde >= IPPositionStart Then
    GetField = Mid(Field, IPPositionStart, IPPositionEnde - IPPositionStart)
  End If
End Function

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Show

  On Error GoTo ErrorHandler:
          
  'the following routines are nessessary to beware of errors
  If wsReceive.State <> sckClosed Then '# Reset if winsock was in use
    wsReceive.Close
  End If
  
  '# Init the Winsock
  wsReceive.LocalPort = FTPort '# set the winsock receive port to the selected one
  wsReceive.Listen
  Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
  wsReceive.Close
  Close
End Sub

Private Sub timerSpeed_Timer()
  '# Hehe Tricky...set a global variable to count up the bytes and calculate them to KBps every second
  lblSpeed.Caption = "Speed: " & Format(DoneBytes / 1000, "###0.0") & " / KBps"
  DoneBytes = 0
End Sub

Private Sub wsReceive_ConnectionRequest(ByVal requestID As Long)
  '# accept the connections
  
  If wsReceive.State <> sckClosed Then
    wsReceive.Close
  End If
  
  wsReceive.Accept requestID
  
  frmReceive.Visible = True
  
  frmReceive.Caption = "Receiving File..."
  BytesTotal = 0
  DoneBytes = 0
  
  frmReceive.Show
  
  '# We use the close event to close the file afterwards
End Sub

Private Sub wsReceive_DataArrival(ByVal BytesTotal As Long)
  Dim StrData As String
  Dim lNewValue As Long
  Dim Info As String
    
  StrData = "" '# You only get filedata trought that winsock
               ' so you only have to write it in the file opened before
  wsReceive.GetData StrData, vbString
   
  '# Thats some file info send before we receive the first package
  Info = Left(StrData, 8)
  If Info = "FILEINFO" Then
    FileSize = GetField(StrData, 2)
    FileName = GetField(StrData, 3)
    lblFileSize.Caption = "Filesize: " + Str(FileSize)
    lblFileName.Caption = "Filename: " + FileName
    DownloadingFile = FreeFile
    Open "C:\Windows\Desktop\" & FileName For Binary Access Write As #DownloadingFile
    Exit Sub
  End If

  Info = Left(StrData, 3)
  If Info = "EOF" Then
    Close #DownloadingFile '# File Ready
    frmReceive.Caption = "File completed. Connection closed"
    FileBar.Value = 100
    lblComplete.Caption = "Complete: 100%"
    DoneBytes = 0
    lblSpeed.Caption = "Speed: 0.0 / KBps"
    wsReceive.Close '# Close the winsock
    DoEvents
    TotalBytes = 0
    Exit Sub
  End If

  FileBar.Value = TotalBytes * 100 / FileSize
  DoneBytes = DoneBytes + BytesTotal
  TotalBytes = TotalBytes + BytesTotal
  lblComplete.Caption = "Complete: " & Int(TotalBytes * 100 / FileSize) & " %"
    
  Put #DownloadingFile, , StrData
  DoEvents
End Sub
