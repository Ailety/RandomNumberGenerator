VERSION 5.00
Begin VB.Form Announcement 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "更新公告: Ver. 版本 NULLNULL"
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
      Interval        =   10
      Left            =   3960
      Top             =   8040
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
Dim AnnouncementMessage(1 To 100) As String
Dim AnnouncementMessageAmount As Integer
Dim AnnouncementMessageLoadAmount As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub AnnouncementLoad_Timer()
  If AnnouncementMessageLoadAmount <= AnnouncementMessageAmount Then
    If AnnouncementMessageLoadAmount = 1 Then
      AnnouncementText.Text = AnnouncementMessage(AnnouncementMessageLoadAmount)
      AnnouncementMessageLoadAmount = AnnouncementMessageLoadAmount + 1
    Else
      AnnouncementText.Text = AnnouncementText.Text + vbCrLf + AnnouncementMessage(AnnouncementMessageLoadAmount)
      AnnouncementMessageLoadAmount = AnnouncementMessageLoadAmount + 1
    End If
  Else
    AnnouncementLoad.Enabled = False
  End If
End Sub

Private Sub Confirm_Click()
  Unload Announcement
End Sub

Private Sub Form_Load()

  Dim ReadString As Long
  Dim ReadValue As String
  ReadValue = String(255, 0)
  If Dir(App.Path & "\Meta.vbd") <> "" Then
    Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
  End If
  ReadString = GetPrivateProfileString("Application_Data", "Announcement", "NULL", ReadValue, 256, App.Path & "\config.ini")
  Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
  If Left(ReadValue, 8) = "Nolonger" Then
    NoLongerRemind.Value = 1
  End If

  Announcement.Icon = Main.Icon
  Announcement.Caption = "更新公告: " + Meta.Version + " 20230827"
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
  AnnouncementMessageLoadAmount = 1
  AnnouncementMessageAmount = 91
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
  AnnouncementMessage(16) = "1.重写窗口缩放功能，现支持三种模式: 窗口最大化(默认)、窗口缩放化及窗口最小化"
  AnnouncementMessage(17) = "2.改进部分文本"
  AnnouncementMessage(18) = "3.新增欢迎界面"
  AnnouncementMessage(19) = "4.程序本体UI迭代优化"
  AnnouncementMessage(20) = "5.修复了“缩放功能切换界面卡顿、坐标错误”、“随机数生成功能偶现数据错误”、“部分功能实际情况与描述不符”等问题"
  AnnouncementMessage(21) = "6.对部分功能进行了逻辑优化"
  AnnouncementMessage(22) = "7.加入了重复程序进程判断，现在只允许启动一个进程以保护数据安全"
  AnnouncementMessage(23) = "8.针对单次随机数生成进行了逻辑性优化，现在生成10次后，每一次都将不会抽中前10次被抽中的人。"
  AnnouncementMessage(24) = "9.更新了数据统计功能"
  AnnouncementMessage(25) = ""
  AnnouncementMessage(26) = "3.1.2 - 3.2.3 更新内容"
  AnnouncementMessage(27) = ""
  AnnouncementMessage(28) = "1.修复了班级选择窗口偶现加载异常及程序主窗体错位的问题"
  AnnouncementMessage(29) = "2.优化了保底机制的启用机制，并增加了启用开关"
  AnnouncementMessage(30) = ""
  AnnouncementMessage(31) = "3.2.3 - 3.2.9 更新内容"
  AnnouncementMessage(32) = ""
  AnnouncementMessage(33) = "1.修复了窗体缩放功能因部分代码逻辑异常导致切换窗体时显示位置错位的问题 - 20221014"
  AnnouncementMessage(34) = "2.修复了查看上次数据功能因为更改生成次数导致的下标越界错误 - 20221022"
  AnnouncementMessage(35) = "3.修复了因最小值和最大值相同导致的抽取不符逻辑（如抽到不应该抽到的同学）- 20221108"
  AnnouncementMessage(36) = "4.修复了性别筛选部分代码逻辑错误导致的部分功能异常的问题 - 20221119"
  AnnouncementMessage(37) = "5.对导入班级数据的部分代码进行了调优 - 20221119"
  AnnouncementMessage(38) = "6.修改了更新公告的文本显示方式，整合了最近几次的更新 - 20221119"
  AnnouncementMessage(39) = ""
  AnnouncementMessage(40) = "3.2.9 - 3.3.0 更新内容 20221203"
  AnnouncementMessage(41) = ""
  AnnouncementMessage(42) = "1.对导入班级数据的部分代码和配置文件进行了调整，以免出现因性别参数错误导致的性别筛选功能异常的问题"
  AnnouncementMessage(43) = "2.在数据统计界面增加了部分功能"
  AnnouncementMessage(44) = "3.优化了生成次数、最大值、最小值文本框的参数判断，现在直接禁用了除数字以外的文本输入"
  AnnouncementMessage(45) = "4.修复了因先前版本改动初始化界面导致的无班级数据或切换班级时出现的异常错误（例如主界面载入失败、切换班级后同学数据异常等）"
  AnnouncementMessage(46) = ""
  AnnouncementMessage(47) = "3.3.0 - 3.3.1 更新内容 20221207"
  AnnouncementMessage(48) = ""
  AnnouncementMessage(49) = "1.对2009、2024班级的错误班级序号进行了修正"
  AnnouncementMessage(50) = ""
  AnnouncementMessage(51) = "3.3.1 - 3.3.2 更新内容 20221210"
  AnnouncementMessage(52) = ""
  AnnouncementMessage(53) = "1.修复了因部分随机结算代码逻辑导致的最大值被抽取概率偏高的问题"
  AnnouncementMessage(54) = "2.修复了因数据统计代码的部分逻辑错误导致的在关闭生成多次后会出现统计学生数据导致的下标溢出问题"
  AnnouncementMessage(55) = ""
  AnnouncementMessage(56) = "3.3.2 - 3.3.3 更新内容 20221217"
  AnnouncementMessage(57) = ""
  AnnouncementMessage(58) = "1.对随机数生成代码进行了部分逻辑改动和优化，略微提升了执行效率"
  AnnouncementMessage(59) = "2.对部分读取配置代码进行了调整以适应即将推出的[随机数生成器配置定义程序]所生成的配置文件"
  AnnouncementMessage(60) = ""
  AnnouncementMessage(61) = "3.3.3 - 3.3.4 更新内容 20221220"
  AnnouncementMessage(62) = ""
  AnnouncementMessage(63) = "1.对随机数生成代码进行了部分逻辑改动和优化，再次略微提升了执行效率"
  AnnouncementMessage(64) = "2.对生成功能的代码追加了性能监控器(Part1)，现在会在生成结果界面显示生成所需时间"
  AnnouncementMessage(65) = "3.对生成次数框的次数限制功能进行了调整，使其更符合逻辑"
  AnnouncementMessage(66) = "4.修复了公告界面的不再提醒复选框无法显示对应设置的问题"
  AnnouncementMessage(67) = "5.修复了在数据强制覆盖功能未启用时生成器的[生成]和[查看上次数据]功能逻辑表现错误的问题"
  AnnouncementMessage(68) = ""
  AnnouncementMessage(69) = "3.3.4 - 3.3.5 更新内容 20221221"
  AnnouncementMessage(70) = ""
  AnnouncementMessage(71) = "1.修复了在初始化中的特殊情况下弹出“不能在模式化窗体中打开非模式化窗体”报错的问题 解决方案: 在公告显示与否代码执行完之前，禁用查看上次数据功能"
  AnnouncementMessage(72) = "2.修复了切换班级功能在显示当前班级时无法正常显示的错误"
  AnnouncementMessage(73) = "3.对程序初始化数据的部分代码进行了修改及调优"
  AnnouncementMessage(74) = ""
  AnnouncementMessage(75) = "3.3.5 - 3.3.6 更新内容 20230207"
  AnnouncementMessage(76) = ""
  AnnouncementMessage(77) = "1.修复了性能监视器在3.3.5版本调整代码时所遗留的错误，现在耗时能够正确地显示数值，且不再会出现负数"
  AnnouncementMessage(78) = "2.对随机数生成机制进行了调优 : )"
  AnnouncementMessage(79) = ""
  AnnouncementMessage(80) = "3.3.6 - 3.3.7 更新内容 20230510"
  AnnouncementMessage(81) = ""
  AnnouncementMessage(82) = "1.追加了程序日志功能，在程序启动后会将大多数操作都写入\Log\操作ID.log文件下"
  AnnouncementMessage(83) = "2.修复了部分影响不大的错误"
  AnnouncementMessage(84) = ""
  AnnouncementMessage(85) = "3.3.7 - 3.3.8 更新内容 20230827"
  AnnouncementMessage(86) = ""
  AnnouncementMessage(87) = "1.移除了 [2008] [2009] [2024] 班级的数据"
  AnnouncementMessage(88) = "2.导入了 [2109] [2111] 班级的数据"
  AnnouncementMessage(89) = "3.优化了班级配置文件的数据结构，以方便修改和读取"
  AnnouncementMessage(90) = "4.修复了切换班级时，部分数据无法正确同步的问题"
  AnnouncementMessage(91) = "5.调整了更新公告的加载逻辑"
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
