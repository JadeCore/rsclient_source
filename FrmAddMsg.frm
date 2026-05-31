VERSION 5.00
Begin VB.Form FrmAddMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Message"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   18
      Top             =   1860
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame FrTxtEffect 
      Caption         =   "Text Effect"
      Height          =   1575
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton OptEffect 
         Caption         =   "Wave2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   21
         Tag             =   "wave2:"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptEffect 
         Caption         =   "Slide"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   20
         Tag             =   "slide:"
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptEffect 
         Caption         =   "Shake"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   19
         Tag             =   "shake:"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptEffect 
         Caption         =   "Wave"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Tag             =   "wave:"
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptEffect 
         Caption         =   "Scroll"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Tag             =   "scroll:"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptEffect 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame FrColor 
      Caption         =   "Text Colour"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
      Begin VB.OptionButton OptColor 
         Caption         =   "Flash3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   12
         Tag             =   "flash3:"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Flash2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   11
         Tag             =   "flash2:"
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Flash1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   10
         Tag             =   "flash1:"
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Glow3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   9
         Tag             =   "glow3:"
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Glow2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Tag             =   "glow2:"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Glow1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Tag             =   "glow1:"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Tag             =   "green:"
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Tag             =   "white:"
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "red:"
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptColor 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   180
      Width           =   5595
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAddMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TxtColor
Dim TxtEffect
Private Sub CmdAdd_Click()
    Call Add
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Add
End Sub

Private Sub Form_Load()
    TxtMsg.BackColor = RGB(77, 127, 230)
End Sub

Sub Add()
    If Trim(TxtMsg.Text) <> "" Then
        For i = 0 To OptColor.UBound
            If OptColor(i).Value = True Then
                TxtColor = OptColor(i).Tag
            End If
        Next
        For i = 0 To OptEffect.UBound
            If OptEffect(i).Value = True Then
                TxtEffect = OptEffect(i).Tag
            End If
        Next
        autotyper.MsgList.AddItem TxtColor & TxtEffect & TxtMsg.Text
        Unload Me
    Else
        MsgBox "The message box can not be blank!", vbCritical
    End If
End Sub


Private Sub TxtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Add
End Sub
