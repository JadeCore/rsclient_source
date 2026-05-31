VERSION 5.00
Begin VB.Form Autoclicker 
   Caption         =   "Auto Clicker"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3480
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop F2 or F7"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start F1 or F6"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3480
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   840
   End
   Begin VB.Label Label7 
      Caption         =   "Usage"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "2. Hit start"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "1. Set a speed"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "Autoclicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text5 = Text5 - -1
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True
If Text3 = "" Then Text3 = "0"
If Text4 = "" Then Text4 = "0"
Timer2.Interval = (Text3 * 1000) - -(Text4)
End Sub

Private Sub Command3_Click()
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim pos
    Dim pt As PointAPI
    pos = GetCursorPos(pt)
    Text1.Text = pt.X
    Text2.Text = pt.Y
End Sub

Private Sub Timer2_Timer()
    Dim xP As Long
    Dim yP As Long
    xP = Text1.Text
    yP = Text2.Text
    MouseMove (xP), (yP)
    LeftClick (xP), (yP)
End Sub

Private Sub Timer3_Timer()
If GetAsyncKeyState(vbKeyF1) < 0 Or GetAsyncKeyState(vbKeyF6) < 0 Then
Call Command2_Click
End If
If GetAsyncKeyState(vbKeyF2) Or GetAsyncKeyState(vbKeyF7) < 0 Then
Call Command3_Click
End If
End Sub
