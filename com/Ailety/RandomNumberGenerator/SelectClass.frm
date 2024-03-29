VERSION 5.00
Begin VB.Form SelectClass 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "程序初始化 - 请选择班级"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5205
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox PTS 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox SelectClassCombo 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "SelectClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varPass As String
Dim DefaultClass As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub Confirm_Click()
  Dim WriteString As Long
  Dim ReadString As Long
  Dim ReadValue As String
  Dim Cache As String
  ReadValue = String(255, 0)
  If SelectClassCombo.Text <> "" Then
    ReadString = GetPrivateProfileString("Application_Data", "KeyConfirm", "NULL", ReadValue, 256, App.Path & "\config.ini")
    PTS.Text = ReadValue
    If SelectClassCombo.Text <> Meta.Class Then
      If Mid(PTS.Text, 1, 3) = "Key" And Right(PTS.Text, 3) = "key" Then
        Cache = JiaMi(SelectClassCombo.Text)
        WriteString = WritePrivateProfileString("Application_Data", "DefaultClass", Cache, App.Path & "\config.ini")
        WriteString = WritePrivateProfileString("Application_Data", "KeyConfirm", "", App.Path & "\config.ini")
        SelectClass.Hide
        Meta.Class = SelectClassCombo.Text
        ReadString = GetPrivateProfileString(CStr(Meta.Class), "MateAmount", "NULL", ReadValue, 256, App.Path & "\config.ini")
        Meta.MateAmount = Val(ReadValue)
        Meta.MaleAmount = 0
        Meta.FemaleAmount = 0
        'ReadString = GetPrivateProfileString(CStr(Meta.Class), "MateMale", "NULL", ReadValue, 256, App.Path & "\config.ini")
        'Meta.MaleAmount = Val(ReadValue)
        'ReadString = GetPrivateProfileString(CStr(Meta.Class), "MateFemale", "NULL", ReadValue, 256, App.Path & "\config.ini")
        'Meta.FemaleAmount = Val(ReadValue)
        For i = 1 To Meta.MateAmount Step 1
          ReadString = GetPrivateProfileString(Meta.Class, "MateName(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
          PTS.Text = ReadValue
          If PTS.Text = "NULL" Or PTS.Text = "" Then
            Meta.Name(i) = "姓名异常"
          Else
            Meta.Name(i) = PTS.Text
          End If
          ReadString = GetPrivateProfileString(Meta.Class, "MateGender(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
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
        If Dir(App.Path & "\config.ini") <> "" Then
          Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
        End If
        PTS.Text = "Load"
        SelectClass.Hide
        Main.Show
        Main.ClassDisplay.Caption = "当前已载入 " + Meta.Class + "班 学生数据"
        Main.MaxBox.Text = CStr(Meta.MateAmount)
        MsgBox "数据载入已完成！", vbOKOnly, "载入完成"
        Main.SetFocus
        Unload SelectClass
      Else
        MsgBox "更新数据所需的密钥错误！" + vbCrLf + "请联系开发者以获取更新班级数据密钥！", vbOKOnly + vbCritical, "更新密钥错误"
      End If
    Else
      MsgBox "当前加载的数据已经是" + Meta.Class + "班！", vbOKOnly + vbExclamation, "已加载" + Meta.Class + "班级数据"
    End If
  Else
    MsgBox "请选择一个班级用于初次载入班级配置，" + vbCrLf + "确认之后每次启动都会自行加载该班级配置。", vbOKOnly + vbExclamation, "请选择一个班级"
  End If
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
  If Meta.Class = "" Then
    If Dir(App.Path & "\config.ini") <> "" Then
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
    End If
    If PTS.Text <> "Load" Then
      End
    End If
  Else
    If Dir(App.Path & "\config.ini") <> "" Then
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
    End If
    Main.Show
    Main.SetFocus
  End If
End Sub

Private Sub Form_Load()
  SelectClassCombo.AddItem "2109"
  SelectClassCombo.AddItem "2111"
  If Meta.Class <> "" Then
    SelectClass.Icon = Main.Icon
    SelectClass.Caption = "程序参数设置 - 选择班级"
    Select Case Meta.Class
    Case "2109"
      SelectClassCombo.ListIndex = 0
    Case "2111"
      SelectClassCombo.ListIndex = 1
    End Select
  Else
    SelectClass.Icon = Welcome.Icon
    Unload Welcome
  End If
End Sub

