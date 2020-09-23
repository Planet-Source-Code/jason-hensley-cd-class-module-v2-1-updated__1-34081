VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " CD Player Class Module v2.1"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   Icon            =   "CDclass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Eject"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox track 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Track: 01"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   720
   End
   Begin VB.CommandButton cmdRewind 
      Caption         =   "Rewind"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label L1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Error Status:"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CD As New CDClass

Private Sub cmdForward_Click()
    CD.fastForward 10
    L1.Caption = "Error Status: " & CD.checkError
End Sub

Private Sub cmdOpen_Click()
    If cmdOpen.Caption = "Eject" Then
        CD.setDoorOpen
        L1.Caption = "Error Status: " & CD.checkError
        cmdOpen.Caption = "Close"
    Else
        CD.setDoorClosed
        L1.Caption = "Error Status: " & CD.checkError
        cmdOpen.Caption = "Eject"
    End If
End Sub

Private Sub cmdPause_Click()
    If cmdPause.Caption = "Pause" Then
        CD.pauseCD
        L1.Caption = "Error Status: " & CD.checkError
        cmdPause.Caption = "Resume"
    Else
        CD.resumeCD
        L1.Caption = "Error Status: " & CD.checkError
        cmdPause.Caption = "Pause"
    End If
End Sub

Private Sub cmdRewind_Click()
    CD.fastRewind 10
    L1.Caption = "Error Status: " & CD.checkError
End Sub

Private Sub Command1_Click()
    CD.playCD
    L1.Caption = "Error Status: " & CD.checkError
End Sub

Private Sub Command2_Click()
    CD.stopCD
    L1.Caption = "Error Status: " & CD.checkError
End Sub

Private Sub Form_Load()
    'start the cd
    'be sure to change "e:\" to your cd drive
    CD.startCD ("e:\")
    
    L1.Caption = "Error Status: " & CD.checkError
    
    For i = 1 To CD.getNumberTracks
        If i <= 9 Then
            track.AddItem "Track: 0" & i
        Else
            track.AddItem "Track: " & i
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CD.stopCD
    CD.closeAll
End Sub

Private Sub Timer1_Timer()
    L2.Caption = CD.getPositionTMSF
    'L3.Caption = CD.getPositionMSF
    'L4.Caption = CD.getPositionTMSF
    'L5.Caption = CD.getPositioninTracks
End Sub

Private Sub track_Click()
    CD.setTrack track.ListIndex + 1
End Sub
