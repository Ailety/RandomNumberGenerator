VERSION 5.00
Begin VB.Form Welcome 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "欢迎界面"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10575
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10575
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Confirm 
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton OpenBrowser 
      Caption         =   "点击打开浏览器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "2942060024"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "AiletyAccount"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Ailety"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "@Ailety"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Ailety"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Ailety@outlook.com"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "https://github.com/Ailety/RandomNumberGenerator"
      Top             =   3480
      Width           =   5895
   End
   Begin VB.TextBox WelcomeText 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
   Begin VB.Label Way1 
      BackStyle       =   0  'Transparent
      Caption         =   "QQ："
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
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Way2 
      BackStyle       =   0  'Transparent
      Caption         =   "微信："
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
      Left            =   600
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Way6 
      BackStyle       =   0  'Transparent
      Caption         =   "Telegram："
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
      Left            =   600
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Way4 
      BackStyle       =   0  'Transparent
      Caption         =   "Twitter："
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
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Way5 
      BackStyle       =   0  'Transparent
      Caption         =   "Discord："
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
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Way3 
      BackStyle       =   0  'Transparent
      Caption         =   "邮箱："
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
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Way7 
      BackStyle       =   0  'Transparent
      Caption         =   "该项目已在github上开源，并遵循 GPL3.0 协议。如果你有能力改进或想学习本软件，可以通过github仓库下载源代码。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   2880
      Width           =   5895
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub OpenWebPage1()
  ShellExecute 0&, vbNullString, "https://github.com/Ailety/RandomNumberGenerator", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Confirm_Click()
  Me.Hide
  SelectClass.Show
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
    Unload Me
    Exit Sub
  End If
  WelcomeText.Text = "" + vbCrLf + "                                                             欢迎使用随机数生成器(RNG)" + vbCrLf + " " + vbCrLf + "　 这个软件因老师上课的需求而诞生，如今已迭代至 SNAPSHOT 3.1.2 (第三快照版本第一次更新+二次修订)，功能也相对趋于完善。初次开发花费1节课，后续的更新和维护共计35.7小时(实际开发时长)。" + vbCrLf + "　 当前版本解决了很多初代版本所存在的痛点，同时也修复了99%的BUG。但是受限于精力和技术，可能存在着极为隐性的漏洞，欢迎反馈。当然，如果你有好的建议，也可以与我联系，让软件更加完善。"
End Sub

Private Sub OpenBrowser_Click()
  Call OpenWebPage1
End Sub
