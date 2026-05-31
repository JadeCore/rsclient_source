VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Rsclient - By: Gary's Hood"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Usage Instructions"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Auto Clicker"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Update / Help"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Typer"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculator"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo W98
Shell "C:\Windows\System32\calc.exe"
Exit Sub
W98:
Shell "C:\Windows\calc.exe"
End Sub

Private Sub Command10_Click()
ShellExecute 0, vbNullString, "http://www.garyshood.com/rsclient/", "", "", 1
End
End Sub

Private Sub Command11_Click()
Autoclicker.Show
End Sub

Private Sub Command2_Click()
autotyper.Show
End Sub
