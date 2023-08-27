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
   Begin VB.TextBox PTS 
      Height          =   270
      Left            =   4080
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
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
      Caption         =   "点击转至浏览器"
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
Dim TimeResult As String
Dim varPass As String
Dim DefaultClass As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub OpenWebPage1()
  ShellExecute 0&, vbNullString, "https://github.com/Ailety/RandomNumberGenerator", vbNullString, vbNullString, vbNormalFocus
End Sub

Function TimeFormat(TimePapi As Integer)
  If TimePapi < 10 Then
    TimeResult = "0" & CStr(TimePapi)
  Else
    TimeResult = CStr(TimePapi)
  End If
End Function

Private Function JiaMi(ByVal varPass As String) As String
  Dim varJiaMi As String * 20
  Dim varTmp As Double
  Dim strJiaMi As String
  Dim i
  For i = 1 To Len(varPass)
    varTmp = AscW(Mid$(varPass, i, 1))
    varJiaMi = Str$(((((varTmp * 1.5) / 5.6) * 2.7) * i))
    strJiaMi = strJiaMi & varJiaMi
  Next i
  JiaMi = strJiaMi
End Function

Private Function JieMi(ByVal varPass As String) As String
  Dim varReturn As String * 20
  Dim varConvert As Double
  Dim varFinalPass As String
  Dim varKey As Integer
  Dim varPasslenth As Long
  varPasslenth = Len(varPass)
  For i = 1 To varPasslenth / 20
    varReturn = Mid(varPass, (i - 1) * 20 + 1, 20)
    varConvert = Val(Trim(varReturn))
    varConvert = ((((varConvert / 1.5) * 5.6) / 2.7) / i)
    varFinalPass = varFinalPass & ChrW(Val(varConvert))
  Next i
  JieMi = varFinalPass
End Function

Private Sub Confirm_Click()
  Dim WriteString As Long
  Dim ReadString As Long
  Dim ReadValue As String
  If Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") <> "" Then
    Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
  End If
  If Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") = "" Then
    Welcome.Hide
    MsgBox "加载配置文件时出现错误，" + vbCrLf + "如持续出现该错误，请报告给开发者。", vbOKOnly + vbCritical, "配置文件错误"
    End
  End If
  If Dir(App.Path & "\config.ini") <> "" Then
    ReadValue = String(255, 0)
    ReadString = GetPrivateProfileString("Application_Data", "DefaultClass", "NULL", ReadValue, 256, App.Path & "\config.ini")
    PTS.Text = ReadValue
    If PTS.Text = "" Then
      Welcome.Hide
      SelectClass.Show
      Exit Sub
    End If
    DefaultClass = JieMi(ReadValue)
    '判断班级数量？
    'XXXXX
    '判断班级名称符合？
    If Mid(DefaultClass, 1, 4) = "2109" Or Mid(DefaultClass, 1, 4) = "2111" Then
      Meta.Class = CStr(Val(DefaultClass))
      ReadString = GetPrivateProfileString(DefaultClass, "MateAmount", "NULL", ReadValue, 256, App.Path & "\config.ini")
      Meta.MateAmount = Val(ReadValue)
      Meta.MaleAmount = 0
      Meta.FemaleAmount = 0
      'ReadString = GetPrivateProfileString(DefaultClass, "MateMale", "NULL", ReadValue, 256, App.Path & "\config.ini")
      'Meta.MaleAmount = Val(ReadValue)
      'ReadString = GetPrivateProfileString(DefaultClass, "MateFemale", "NULL", ReadValue, 256, App.Path & "\config.ini")
      'Meta.FemaleAmount = Val(ReadValue)
      For i = 1 To Meta.MateAmount Step 1
        ReadString = GetPrivateProfileString(DefaultClass, "MateName(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
        PTS.Text = ReadValue
        If PTS.Text = "NULL" Or PTS.Text = "" Then
          Meta.Name(i) = "姓名异常"
        Else
          Meta.Name(i) = PTS.Text
        End If
        ReadString = GetPrivateProfileString(DefaultClass, "MateGender(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
        PTS.Text = ReadValue
        If PTS.Text = "NULL" Or PTS.Text = "" Then
          Meta.Gender(i) = "性别异常"
        Else
          Meta.Gender(i) = PTS.Text
        End If
        If Meta.Gender(i) = "男" Then
          Meta.MaleAmount = Meta.MaleAmount + 1
        ElseIf Meta.Gender(i) = "女" Then
          Meta.FemaleAmount = Meta.FemaleAmount + 1
        End If
      Next i
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
      Savetime = timeGetTime
      While timeGetTime < Savetime + 50
      DoEvents
      Wend
      Main.Show
      Main.SetFocus
      Welcome.Hide
      Exit Sub
    Else
      WriteString = WritePrivateProfileString("Application_Data", "DefaultClass", "", App.Path & "\config.ini")
      MsgBox "班级参数错误！" + vbCrLf + "数据载入失败，已清除配置文件中错误的班级数据。", vbOKOnly + vbCritical, "参数错误"
      SelectClass.Show
      Exit Sub
    End If
  End If
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
    End
    Exit Sub
  End If
  WelcomeText.Text = "" + vbCrLf + "                                                             欢迎使用随机数生成器(RNG)" + vbCrLf + " " + vbCrLf + "　 本软件因老师上课的需求而诞生，如今已迭代至 SNAPSHOT 3.3.8 (第三快照版本第三次更新+八次修正)，功能也相对趋于完善。初次开发花费1节课，后续的更新和维护共计45.7小时(实际开发时长)。" + vbCrLf + "　 当前版本解决了很多初代版本所存在的痛点，同时也修复了99%的BUG。但是受限于精力和技术，可能存在着极为隐性的漏洞，欢迎反馈。当然，如果你有好的建议，也可以与我联系，让软件更加完善。"
  
  Dim OperationTime As String
  
  T_Year = Year(Now)
  OperationTime = OperationTime & T_Year
  
  T_Month = Month(Now)
  TimeFormat (T_Month)
  OperationTime = OperationTime & TimeResult
  
  T_Day = Day(Now)
  TimeFormat (T_Day)
  OperationTime = OperationTime & TimeResult
  
  T_Hour = Hour(Now)
  TimeFormat (T_Hour)
  OperationTime = OperationTime & TimeResult
  
  T_Minute = Minute(Now)
  TimeFormat (T_Minute)
  OperationTime = OperationTime & TimeResult
  
  T_Second = Second(Now)
  TimeFormat (T_Second)
  OperationTime = OperationTime & TimeResult
  
  Meta.RNG_OperationID = "Operation - " & OperationTime
  Open App.Path & "\Log\" & Meta.RNG_OperationID & ".log" For Append As #1
  Print #1, Now & " " & "程序启动，载入欢迎界面"
  Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  Call Confirm_Click
  Cancel = -1
End Sub

Private Sub OpenBrowser_Click()
  Call OpenWebPage1
End Sub
