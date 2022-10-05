VERSION 5.00
Begin VB.Form MinWindow 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   Begin VB.Timer FunTimer 
      Interval        =   200
      Left            =   960
      Top             =   840
   End
   Begin VB.Image BackGround 
      Height          =   765
      Left            =   0
      Picture         =   "MinWindow.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   765
   End
End
Attribute VB_Name = "MinWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mouseIsDown As Boolean
Dim cx As Single
Dim cy As Single
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置

Private Sub BackGround_DblClick()
  Select Case Meta.WindowLastState
    Case "Max"
      Main.Window_Display_Max.Value = True
      Meta.WindowState = "Max"
      MinWindow.Top = Main.Top - 2300
      MinWindow.Left = Main.Left - 4600
      Meta.WindowLastState = Meta.WindowState
      Meta.WindowState = "Max"
    Case "Normal"
      Main.Top = MinWindow.Top - 400
      Main.Left = MinWindow.Left - 1000
      Main.Window_Display_Max.Value = True
      Meta.WindowLastState = Meta.WindowState
      Meta.WindowState = "Normal"
  End Select
  Main.Show
  Unload MinWindow
End Sub

Private Sub Form_Load()
  If Meta.WindowLastState = "Max" Then
    MinWindow.Top = Main.Top + 2300
    MinWindow.Left = Main.Left + 4600
  ElseIf Meta.WindowLastState = "Normal" Then
    MinWindow.Top = Main.Top + 400
    MinWindow.Left = Main.Left + 1000
  End If
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub BackGround_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  mouseIsDown = True
  cx = x
  cy = y
End Sub
Private Sub BackGround_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If mouseIsDown Then
    Move Left + (x - cx), Top + (y - cy)
  End If
End Sub
Private Sub BackGround_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  mouseIsDown = False
End Sub

Private Sub FunTimer_Timer()
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
