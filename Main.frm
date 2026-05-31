VERSION 5.00
Begin VB.Form Autotyper 
   Caption         =   "RuneScape Auto Typer"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   3
      Left            =   2400
      Top             =   1200
   End
   Begin VB.Frame MFr 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Timer WayTime 
         Interval        =   1
         Left            =   7320
         Top             =   360
      End
      Begin VB.ListBox MsgList 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2970
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Panel"
         Height          =   1455
         Left            =   180
         TabIndex        =   6
         Top             =   3600
         Width           =   7455
         Begin VB.PictureBox Pic1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1035
            Left            =   60
            ScaleHeight     =   1035
            ScaleWidth      =   7095
            TabIndex        =   7
            Top             =   180
            Width           =   7095
            Begin VB.TextBox Way1 
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Text            =   "2"
               Top             =   240
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.CommandButton CmdEdit 
               Caption         =   "Edit"
               Height          =   315
               Left            =   2820
               TabIndex        =   3
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton CmdAddMsg 
               Caption         =   "Add"
               Height          =   315
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton CmdStart 
               Caption         =   "Start (F12)"
               Height          =   315
               Left            =   5580
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton CmdDelete 
               Caption         =   "Delete"
               Height          =   315
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtTime 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   1380
               MaxLength       =   4
               TabIndex        =   5
               Text            =   "5"
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   2340
               TabIndex        =   10
               Top             =   720
               Width           =   4575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Message Rate:"
               Height          =   195
               Left            =   60
               TabIndex        =   9
               Top             =   720
               Width           =   1260
            End
            Begin VB.Label Label2 
               Caption         =   "sec"
               Height          =   255
               Left            =   1920
               TabIndex        =   8
               Top             =   720
               Width           =   375
            End
         End
      End
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Index           =   0
         Left            =   480
         Top             =   1080
      End
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Index           =   1
         Left            =   1080
         Top             =   1080
      End
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Index           =   2
         Left            =   1680
         Top             =   1080
      End
      Begin VB.Timer KeyCatch 
         Interval        =   100
         Left            =   3240
         Top             =   960
      End
   End
   Begin VB.Menu MFile 
      Caption         =   "File (&F)"
      Begin VB.Menu MSave 
         Caption         =   "Quick Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MLoad 
         Caption         =   "Load Quick Save"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MEdit 
      Caption         =   "Edit (&E)"
      Begin VB.Menu MAddMsg 
         Caption         =   "Add Message"
         Shortcut        =   ^A
      End
      Begin VB.Menu MEditMsg 
         Caption         =   "Edit Message"
         Shortcut        =   ^E
      End
      Begin VB.Menu MDelete 
         Caption         =   "Delete Message"
         Shortcut        =   ^D
      End
      Begin VB.Menu MClear 
         Caption         =   "Clear all messages"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu autoweb 
      Caption         =   ""
   End
End
Attribute VB_Name = "autotyper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MsgI
Dim Started As Boolean
Dim Prev As String

Private Sub CmdAddMsg_Click()
    FrmAddMsg.Show vbModal
End Sub

Private Sub CmdDelete_Click()
    On Error GoTo Err
    MsgList.RemoveItem (MsgList.ListIndex)
    Exit Sub
Err:
   MsgBox "Please select an item you want to delete!", vbCritical
End Sub

Private Sub CmdEdit_Click()
    On Error GoTo Err
    If Len(MsgList.List(MsgList.ListIndex)) = 0 Then GoTo Err
    
    Dim a As String
    a = InputBox("Edit Message", , MsgList.List(MsgList.ListIndex))
    If a <> vbNullString Then MsgList.List(MsgList.ListIndex) = a
    Exit Sub
Err:
   MsgBox "Please select an item you want to edit!", vbCritical
End Sub

Private Sub CmdStart_Click()
    Call SStart
End Sub
Function InputTest() As String
    If Not IsNumeric(TxtTime.Text) Then
        InputTest = "Message rate is not a valid number!"
        Exit Function
    End If
    If Val(TxtTime.Text) > 50 Or Val(TxtTime.Text) < 1 Then
        InputTest = "Message rate should be between 1 and 50!"
        Exit Function
    End If
    If MsgList.ListCount < 1 Then
        InputTest = "There is no message to talk!"
        Exit Function
    End If
    InputTest = "1"
End Function
Private Sub Form_Load()
    On Error Resume Next
    RD = Int(GetSetting("Delay", "Delay", "Delay"))
    If GetSetting("Delay", "Delay", "Delay") = "" Then RD = 1
    Started = False
    Way1 = 1
    MsgI = 0
    MsgList.BackColor = RGB(77, 127, 230)
    TxtTime.BackColor = RGB(77, 127, 230)
    TxtTime = GetSetting("TxtTime", "TxtTime", "TxtTime")
    If TxtTime = "" Then TxtTime = "5"
    ShowWay
End Sub

Private Sub Form_LostFocus()
    Timer(Way).Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MFr.Left = 100
    MFr.Top = 10
    MFr.Height = Me.ScaleHeight - 100
    MFr.Width = Me.ScaleWidth - 200
    Frame1.Width = Me.ScaleWidth - 400
    MsgList.Height = MFr.Height - Frame1.Height - 400
    MsgList.Width = MFr.Width - 300
    Frame1.Top = MsgList.Top + MsgList.Height + 100
    Frame1.Left = 100
    Pic1.Left = Frame1.Width / 2 - Pic1.Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "TxtTime", "TxtTime", "TxtTime", TxtTime.Text
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.BorderStyle = 0
End Sub

Private Sub KeyCatch_Timer()
    
    If GetAsyncKeyState(vbKeyF12) < 0 Then
        Prev = "12"
    Else
        If Prev = "12" Then
            Call SStart
            Prev = ""
        End If
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Way1 = 4 Then
    Way1 = 0
End If

If Way1 < 4 Then
    Way1 = Way1 - -1
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.BorderStyle = 1
End Sub


Private Sub MAddMsg_Click()
    FrmAddMsg.Show vbModal
End Sub

Private Sub MClear_Click()
    MsgList.Clear
End Sub

Private Sub MDelete_Click()
    On Error GoTo Err
    MsgList.RemoveItem (MsgList.ListIndex)
    Exit Sub
Err:
   MsgBox "Please select an item you want to delete!", vbCritical
End Sub

Private Sub MFr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.BorderStyle = 0
End Sub

Private Sub MLoad_Click()
On Error GoTo EeRR
    Dim DataR
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.opentextfile(App.Path & "\quicksave.txt", 1)
    DataR = f.readall
    DataR = Split(DataR, vbCrLf)
    For i = 0 To UBound(DataR)
        If Trim(DataR(i)) <> "" Then MsgList.AddItem DataR(i)
    Next
    f.Close
    Set f = Nothing
    Set fs = Nothing
    Exit Sub
EeRR:
    MsgBox "Error Opening File!", vbCritical
End Sub

Private Sub MSave_Click()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path & "\quicksave.txt")
    For i = 0 To MsgList.ListCount - 1
        f.write MsgList.List(i) & vbCrLf
    Next
    f.Close
    Set f = Nothing
    Set fs = Nothing
End Sub


Private Sub MSettings_Click()

End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.BorderStyle = 0
End Sub

Private Sub Timer_Timer(Index As Integer)
If Way1.Text = 1 Then
            If MsgI > MsgList.ListCount - 1 Then MsgI = 0
            SendKeys (MsgList.List(MsgI)) & "{ENTER}"
            MsgList.Selected(MsgI) = True
            MsgI = MsgI + 1
End If
If Way1.Text = 2 Then
            If MsgI < 0 Then MsgI = MsgList.ListCount - 1
            SendKeys (MsgList.List(MsgI)) & "{ENTER}"
            MsgList.Selected(MsgI) = True
            MsgI = MsgI - 1
End If
If Way1.Text = 3 Then
            Randomize
            Rfactor = Int(Rnd * MsgList.ListCount)
            SendKeys (MsgList.List(Rfactor)) & "{ENTER}"
            MsgList.Selected(Rfactor) = True
            End If
If Way1.Text = 4 Then
            SendKeys (MsgList.List(MsgList.ListIndex)) & "{ENTER}"
    Randomize
    Timer(Index).Interval = Val(TxtTime.Text) * 1000 + Rnd * 1000 * RD
    End If

End Sub
Sub DisableAll()
    Me.WindowState = 1
    TxtTime.Locked = True
    CmdDelete.Visible = False
    CmdAddMsg.Visible = False
CmdEdit.Visible = False
    CmdStart.Caption = "&Stop (F12)"
    Started = True
    Me.Caption = "ATX - Started"
End Sub
Sub EnableAll()
    For i = 0 To Timer.UBound
        Timer(i).Enabled = False
    Next

    TxtTime.Locked = False
    CmdDelete.Visible = True
    CmdAddMsg.Visible = True

    CmdEdit.Visible = True
    CmdStart.Caption = "&Start (F12)"
    Started = False
    MsgI = 0
    Me.Caption = "ATX - Stopped"
End Sub
Sub SStart()
    If Started = False Then
        If InputTest <> "1" Then
            MsgBox InputTest, vbCritical
            Exit Sub
        End If
        Timer(Way).Interval = Val(TxtTime.Text) * 1000
        Timer(Way).Enabled = True
        Call DisableAll
        
    Else
        Call EnableAll
    End If
End Sub
Sub ShowWay()
If Way1.Text = 1 Then
            Label3.Caption = "Messages will be sent from top to bottom"
End If
If Way1.Text = 2 Then
            Label3.Caption = "Messages will be sent from bottom to top"
End If
If Way1.Text = 3 Then
            Label3.Caption = "Messages will be sent in random"
End If
If Way1.Text = 4 Then
            Label3.Caption = "Send only seleted message"
End If
End Sub

Private Sub WayTime_Timer()
If Way1.Text = 1 Then
            Label3.Caption = "Messages will be sent from top to bottom"
End If
If Way1.Text = 2 Then
            Label3.Caption = "Messages will be sent from bottom to top"
End If
If Way1.Text = 3 Then
            Label3.Caption = "Messages will be sent in random"
End If
If Way1.Text = 4 Then
            Label3.Caption = "Send only seleted message"
End If
End Sub
