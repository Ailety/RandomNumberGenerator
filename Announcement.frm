VERSION 5.00
Begin VB.Form Announcement 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "更新公告: 3.1.2 版本 20221005"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer AnnouncementLoad 
      Interval        =   25
      Left            =   4080
      Top             =   6960
   End
   Begin VB.CheckBox NoLongerRemind 
      BackColor       =   &H8000000E&
      Caption         =   "不再提醒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   7080
      Width           =   1125
   End
   Begin VB.TextBox AnnouncementText 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton Confirm 
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   7440
      Width           =   2535
   End
End
Attribute VB_Name = "Announcement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AnnouncementMessage(1 To 40) As String
Dim AnnouncementMessageAmount As Integer
Dim AnnouncementMessageLoadAmount As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub AnnouncementLoad_Timer()
  If AnnouncementMessageLoadAmount <= AnnouncementMessageAmount Then
    If AnnouncementMessageLoadAmount = 1 Then
      AnnouncementText.Text = AnnouncementMessage(AnnouncementMessageLoadAmount)
      AnnouncementMessageLoadAmount = AnnouncementMessageLoadAmount + 1
    Else
      AnnouncementText.Text = AnnouncementText.Text + vbCrLf + AnnouncementMessage(AnnouncementMessageLoadAmount)
      AnnouncementMessageLoadAmount = AnnouncementMessageLoadAmount + 1
    End If
  End If
End Sub

Private Sub Confirm_Click()
  Unload Announcement
End Sub

Private Sub Form_Load()
  Announcement.Icon = Main.Icon
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
  AnnouncementMessageLoadAmount = 1
  AnnouncementMessageAmount = 24
  AnnouncementMessage(1) = "1.1.2 - 2.7.2 更新内容"
  AnnouncementMessage(2) = ""
  AnnouncementMessage(3) = "1.重写随机数生成逻辑"
  AnnouncementMessage(4) = "2.修复了14个影响体验的BUG"
  AnnouncementMessage(5) = "3.优化部分代码逻辑"
  AnnouncementMessage(6) = "4.重写班级数据功能，现在支持2008、"
  AnnouncementMessage(7) = "2009 、2024班级"
  AnnouncementMessage(8) = "5.新增了两个方便老师使用的功能: 强制覆盖数据、最小化显示程序"
  AnnouncementMessage(9) = "6.对界面进行了调整性优化"
  AnnouncementMessage(10) = "7.修复了在初次启动或切换班级时偶现"
  AnnouncementMessage(11) = "文件无法找到的非程序本身的错误信息"
  AnnouncementMessage(12) = "8.优化了随机数结果的显示文本大小"
  AnnouncementMessage(13) = ""
  AnnouncementMessage(14) = "2.7.2 - 3.1.2 更新内容"
  AnnouncementMessage(15) = ""
  AnnouncementMessage(16) = "1.重写窗口缩放功能，现支持三种模式: 窗口最大火(默认)、窗口缩放化及窗口最小化"
  AnnouncementMessage(17) = "2.改进部分文本"
  AnnouncementMessage(18) = "3.新增欢迎界面"
  AnnouncementMessage(19) = "4.程序本体UI迭代优化"
  AnnouncementMessage(20) = "5.修复了“缩放功能切换界面卡顿、坐标错误”、“随机数生成功能偶现数据错误”、“部分功能实际情况与描述不符”等问题"
  AnnouncementMessage(21) = "6.对部分功能进行了逻辑优化"
  AnnouncementMessage(22) = "7.加入了重复程序进程判断，现在只允许启动一个进程以保护数据安全"
  AnnouncementMessage(23) = "8.针对单次随机数生成进行了逻辑性优化，现在生成10次后，每一次都将不会抽中前10次被抽中的人。"
  AnnouncementMessage(24) = "9.更新了数据统计功能"
End Sub

Private Sub NoLongerRemind_Click()
  Dim WriteString As Long
  If NoLongerRemind.Value Then
    If Dir(App.Path & "\config.ini") <> "" Then
      WriteString = WritePrivateProfileString("Application_Data", "Announcement", "Nolonger", App.Path & "\config.ini")
    Else
      Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
      WriteString = WritePrivateProfileString("Application_Data", "Announcement", "Nolonger", App.Path & "\config.ini")
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
    End If
  Else
    If Dir(App.Path & "\config.ini") <> "" Then
      WriteString = WritePrivateProfileString("Application_Data", "Announcement", "Always", App.Path & "\config.ini")
    Else
      Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
      WriteString = WritePrivateProfileString("Application_Data", "Announcement", "Always", App.Path & "\config.ini")
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
    End If
  End If
End Sub
