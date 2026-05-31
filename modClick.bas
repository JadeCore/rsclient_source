Attribute VB_Name = "Module1"
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const Author = "GXSLKGKSD"

Declare Function GetCursorPos& Lib "user32" (lpPoint As PointAPI)
Type PointAPI
X As Long
Y As Long
End Type

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4


Sub MouseMove(xP As Long, yP As Long)
Dim move
move = SetCursorPos(xP, yP)
End Sub


Sub LeftClick(xP As Long, yP As Long)
mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
End Sub


