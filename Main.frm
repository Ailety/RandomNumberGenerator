VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "随机数生成器 - RNG - SNAPSHOT Ver."
   ClientHeight    =   5055
   ClientLeft      =   9255
   ClientTop       =   5610
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10215
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox MinimumProtectSwitch 
      BackColor       =   &H8000000E&
      Caption         =   "启用保护机制"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   31
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton DataStatisticsButton 
      Caption         =   " 数据统计 "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   30
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Min_Button 
      Caption         =   "最小化"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6360
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Max_Button 
      Caption         =   "最大化"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer AnnouncementLoad 
      Interval        =   500
      Left            =   9720
      Top             =   600
   End
   Begin VB.CheckBox OverwriteData 
      BackColor       =   &H8000000E&
      Caption         =   "启用强制覆盖功能"
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
      Left            =   5520
      TabIndex        =   20
      Top             =   3000
      Width           =   2115
   End
   Begin VB.CommandButton AmountDown 
      Caption         =   "↓"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   1800
      Width           =   250
   End
   Begin VB.CommandButton AmountUp 
      Caption         =   "↑"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   1800
      Width           =   250
   End
   Begin VB.CheckBox WindowWeight 
      BackColor       =   &H8000000E&
      Caption         =   "程序前置"
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
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Width           =   1125
   End
   Begin VB.CheckBox FormattingData 
      BackColor       =   &H8000000E&
      Caption         =   "格式化显示数据"
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
      Left            =   2400
      TabIndex        =   15
      Top             =   3000
      Width           =   1605
   End
   Begin VB.CheckBox AllowDuplicateData 
      BackColor       =   &H8000000E&
      Caption         =   "允许生成重复数据"
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   1845
   End
   Begin VB.CommandButton ViewLastData 
      Caption         =   "查看上次数据"
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
      Left            =   7680
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Generate 
      Caption         =   "生成"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CheckBox AllowMultiple 
      BackColor       =   &H8000000E&
      Caption         =   "启用生成多次"
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
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CheckBox NameHook 
      BackColor       =   &H8000000E&
      Caption         =   "启用学生姓名挂钩"
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
      TabIndex        =   6
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox AmountBox 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Text            =   "1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox MaxBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "46"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox MinBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame GenderControl 
      BackColor       =   &H8000000E&
      Caption         =   "性别数据筛选"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   4815
      Begin VB.OptionButton OnlyGirl 
         BackColor       =   &H8000000E&
         Caption         =   "仅筛选女生"
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
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   1245
      End
      Begin VB.OptionButton OnlyBoy 
         BackColor       =   &H8000000E&
         Caption         =   "仅筛选男生"
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
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   1245
      End
      Begin VB.OptionButton AllGender 
         BackColor       =   &H8000000E&
         Caption         =   "不筛选性别"
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
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Frame DataControl 
      BackColor       =   &H8000000E&
      Caption         =   "数据管控选项"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Frame ProgramControl 
      BackColor       =   &H8000000E&
      Caption         =   "程序管控选项"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   24
      Top             =   2640
      Width           =   4815
      Begin VB.OptionButton Window_Display_Min 
         BackColor       =   &H8000000E&
         Caption         =   "窗口最小化"
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
         Left            =   3240
         TabIndex        =   27
         Top             =   720
         Width           =   1245
      End
      Begin VB.OptionButton Window_Display_Normal 
         BackColor       =   &H8000000E&
         Caption         =   "窗口缩放化"
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
         Left            =   1800
         TabIndex        =   26
         Top             =   720
         Width           =   1245
      End
      Begin VB.OptionButton Window_Display_Max 
         BackColor       =   &H8000000E&
         Caption         =   "窗口最大化"
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
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Label Title 
      BackStyle       =   0  'Transparent
      Caption         =   "随机数生成器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   22
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Subtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "SNAPSHOT Ver."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   21
      Top             =   1100
      Width           =   3135
   End
   Begin VB.Label ClassDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "当前已载入NULL班学生数据"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Line DataLine 
      BorderWidth     =   3
      X1              =   5280
      X2              =   9840
      Y1              =   2500
      Y2              =   2500
   End
   Begin VB.Label AmountLabel 
      BackColor       =   &H8000000E&
      Caption         =   "生成次数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label MaxLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "最大值"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label MinLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "最小值"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Min As Single
Dim Max As Single
Dim lngAns As Long
Dim RNMax As Single
Dim RNCache(1 To 25000) As Single
Dim RNCacheCount(1 To 1000000) As Single
Dim MateAmount As Integer
Dim OnlyGirlValue As Boolean
Dim OnlyBoyValue As Boolean
Dim AllGenderValue As Boolean
Dim NameHookValue As Boolean
Dim OverwriteDataValue As Boolean
Dim AllowDuplicateDataValue As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Dim MinimumProtect(1 To 10) As Integer
Dim MinimumProtectCount As Integer

Sub ApplyMin()
  Call ValueGet
  If MinBox.Text <> "" And NameHookValue Then
    If Amount = Max - Min + 1 Then
      AllGender.Value = True
      OnlyBoy.Enabled = False
      OnlyGirl.Enabled = False
    Else
      OnlyBoy.Enabled = True
      OnlyGirl.Enabled = True
    End If
  End If
End Sub

Sub ApplyMax()
  Call ValueGet
  If MaxBox.Text <> "" And NameHookValue Then
    If Amount = Max - Min + 1 Then
      AllGender.Value = True
      OnlyBoy.Enabled = False
      OnlyGirl.Enabled = False
    Else
      OnlyBoy.Enabled = True
      OnlyGirl.Enabled = True
    End If
  End If
End Sub

Sub DisplayResult()
  If OverwriteDataValue Then
    Unload ResultForm
    ResultForm.Show
  Else
    ResultForm.Show
  End If
End Sub

Sub ShowControl()
  Title.Visible = True
  Subtitle.Visible = True
  
  MinLabel.Visible = True
  MinBox.Visible = True
  
  MaxLabel.Visible = True
  MaxBox.Visible = True
  
  AmountLabel.Visible = True
  AmountBox.Visible = True
  
  AmountUp.Visible = True
  AmountDown.Visible = True
  
  NameHook.Visible = True
  AllowMultiple.Visible = True
  FormattingData.Visible = True
  AllowDuplicateData.Visible = True
  MinimumProtectSwitch.Visible = True
  OverwriteData.Visible = True
  WindowWeight.Visible = True
  
  DataControl.Visible = True
  GenderControl.Visible = True
  ProgramControl.Visible = True
  
  ViewLastData.Visible = True
  
  DataLine.Visible = True
  ClassDisplay.Visible = True
    
  Max_Button.Visible = False
  Min_Button.Visible = False

  Main.Caption = "随机数生成器 - RNG - SNAPSHOT " + Meta.Version
  Main.Width = 10305
  Main.Height = 5490
  
  Generate.Top = 4080
  Generate.Left = 5160
  
  Window_Display_Max.Value = True
  
  Main.Top = Main.Top - 2000
  Main.Left = Main.Left - 3900
End Sub

Sub HideControl()
  Main.SetFocus

  Title.Visible = False
  Subtitle.Visible = False
  
  MinLabel.Visible = False
  MinBox.Visible = False
  
  MaxLabel.Visible = False
  MaxBox.Visible = False
  
  AmountLabel.Visible = False
  AmountBox.Visible = False
  
  AmountUp.Visible = False
  AmountDown.Visible = False
  
  NameHook.Visible = False
  AllowMultiple.Visible = False
  FormattingData.Visible = False
  AllowDuplicateData.Visible = False
  MinimumProtectSwitch.Visible = False
  OverwriteData.Visible = False
  WindowWeight.Visible = False
  
  DataControl.Visible = False
  GenderControl.Visible = False
  ProgramControl.Visible = False
  
  ViewLastData.Visible = False
  
  DataLine.Visible = False
  ClassDisplay.Visible = False
  
  Max_Button.Visible = True
  Min_Button.Visible = True
  
  Main.Caption = "当前状态: 窗口缩放"
  Main.Width = 2620
  Main.Height = 1880
  
  Generate.Top = 120
  Generate.Left = 120
  Max_Button.Top = 975
  Max_Button.Left = 120
  Min_Button.Top = 975
  Min_Button.Left = 1300
  
  Main.Top = Main.Top + 2000
  Main.Left = Main.Left + 3900
  
  Main.SetFocus
End Sub

Sub RandomEvent()
  Call RealRandom
  If NameHookValue Then
    If MinimumProtectSwitch.Value Then
      For i = 1 To 10 Step 1
        If RNMax = MinimumProtect(i) Then
          Meta.Protect = True
          Call RandomEvent
          Exit Sub
        End If
      Next i
      For i = 1 To 9 Step 1
        MinimumProtect(i) = MinimumProtect(i + 1)
      Next i
      MinimumProtect(10) = RNMax
    End If
  End If
End Sub

Sub RealRandom()
  Min = Val(MinBox.Text)
  Max = Val(MaxBox.Text)
  Amount = Val(AmountBox.Text)
  Call ValueGet
  If NameHook = False Then
    RNMax = Int(Rnd() * (Max - Min + 1)) + Min
    Exit Sub
  End If
  For a = 1 To 25000 Step 1
    RNCache(a) = Int(Rnd() * (Max - Min + 1)) + Min
    RNCacheCount(RNCache(a)) = RNCacheCount(RNCache(a)) + 1
  Next a
  RNMax = Min
  For b = Min To Meta.MateAmount Step 1
    If RNCacheCount(b) >= RNCacheCount(RNMax) Then
      RNMax = b
    End If
  Next b
  RNMin = Min
  For c = Min To Meta.MateAmount Step 1
    If RNCacheCount(c) <= RNCacheCount(RNMin) Then
      RNMin = c
    End If
  Next c
  If Amount <> Max - Min + 1 And Max - Min <> 0 Then
    If Rnd() * 100 + 1 <= 50 Then
      RNMax = RNMin
    End If
  End If
  For d = 1 To Meta.MateAmount Step 1
    RNCacheCount(d) = 0
  Next d
End Sub

Sub ValueGet()
  OnlyGirlValue = OnlyGirl.Value
  OnlyBoyValue = OnlyBoy.Value
  AllGenderValue = AllGender.Value
  NameHookValue = NameHook.Value
  OverwriteDataValue = OverwriteData.Value
  AllowDuplicateDataValue = AllowDuplicateData.Value
End Sub

Private Sub AllowDuplicateData_Click()
  Call ValueGet
  If AllGenderValue Then
    OnlyBoy.Enabled = True
    OnlyGirl.Enabled = True
  End If
  If Not (AllowDuplicateDataValue) Then
    If Amount > Max - Min + 1 And AmountBox.Text <> "" Then
      AmountBox.Text = Max - Min + 1
    End If
  End If
  If NameHook Then
    If OnlyBoyValue Then
      If Amount > Meta.MaleAmount And Not (AllowDuplicateDataValue) Then
        AmountBox.Text = CStr(Meta.MaleAmount)
      End If
    ElseIf OnlyGirlValue Then
      If Amount > Meta.FemaleAmount And Not (AllowDuplicateDataValue) Then
        AmountBox.Text = CStr(Meta.FemaleAmount)
      End If
    End If
  End If
End Sub

Private Sub AllowMultiple_Click()
  If AllowMultiple.Value = 1 Then
    AmountBox.Enabled = True
    AmountUp.Enabled = True
    AllowDuplicateData.Enabled = True
  Else
    AmountBox.Enabled = False
    AllowDuplicateData.Enabled = False
    AmountUp.Enabled = False
    AmountDown.Enabled = False
    AllowDuplicateData.Value = False
    AmountBox.Text = "1"
  End If
End Sub

Private Sub AmountBox_Change()
  If Len(AmountBox.Text) > 5 Then
    MsgBox "生成次数不应该使用过于庞大的数值！", vbOKOnly + vbCritical, "参数错误"
    AmountBox.Text = Left(AmountBox.Text, Len(AmountBox.Text) - 1)
  End If
  Min = Val(MinBox.Text)
  Max = Val(MaxBox.Text)
  Amount = Val(AmountBox.Text)
  Call ValueGet
  Call ApplyMin
  Call ApplyMax
  If Not (IsNumeric(AmountBox.Text)) And AmountBox.Text <> "" Then
    MsgBox "生成次数只允许键入数字！", vbOKOnly + vbCritical, "参数错误"
    AmountBox.Text = "1"
    Exit Sub
  End If
  If Amount > 1 Then
    AmountDown.Enabled = True
  End If
  If NameHook Then
    If OnlyBoyValue Then
      If Amount > Meta.MaleAmount And Not (AllowDuplicateDataValue) Then
        MsgBox "在启用姓名挂钩且勾选了[仅筛选男生]的情况下，" + vbCrLf + "生成次数不允许大于男生的总人数: " + CStr(Meta.MaleAmount) + "人！", vbOKOnly + vbCritical, "参数错误"
        AmountBox.Text = CStr(Meta.MaleAmount)
        Exit Sub
      End If
    ElseIf OnlyGirlValue Then
      If Amount > Meta.FemaleAmount And Not (AllowDuplicateDataValue) Then
        MsgBox "在启用姓名挂钩且勾选了[仅筛选女生]的情况下，" + vbCrLf + "生成次数不允许大于女生的总人数: " + CStr(Meta.FemaleAmount) + "人！", vbOKOnly + vbCritical, "参数错误"
        AmountBox.Text = CStr(Meta.FemaleAmount)
        Exit Sub
      End If
    End If
  End If
  If AmountBox.Text = "0" Or Amount < 0 Then
    MsgBox "生成次数的范围只允许在[1,∞)之间", vbOKOnly + vbExclamation, "参数错误"
    AmountBox.Text = "1"
  Else
    If Not (AllowDuplicateDataValue) Then
      If Amount > Max - Min + 1 And AmountBox.Text <> "" Then
        MsgBox "在禁用允许生成重复数据功能时，为防止数据溢出，" + vbCrLf + "已根据你的设置，计算出生成次数应为不超过 " + CStr(Max - Min + 1) + " 的整数。", vbOKOnly + vbExclamation, "参数错误"
        AmountBox.Text = Max - Min + 1
      End If
    End If
  End If
End Sub

Private Sub AmountDown_Click()
  Amount = Val(AmountBox.Text)
  If AmountBox.Text <> "1" Then
    If Amount - 1 = 1 Then
      AmountDown.Enabled = False
    End If
    AmountBox.Text = CStr(Val(AmountBox.Text) - 1)
  End If
  AmountBox.SetFocus
End Sub

Private Sub AmountUp_Click()
  AmountDown.Enabled = True
  AmountBox.Text = CStr(Val(AmountBox.Text) + 1)
  AmountBox.SetFocus
End Sub

Private Sub AnnouncementLoad_Timer()
  Dim ReadString As Long
  Dim ReadValue As String
  ReadValue = String(255, 0)
  If Dir(App.Path & "\Meta.vbd") <> "" Then
    Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
  End If
  ReadString = GetPrivateProfileString("Application_Data", "Announcement", "NULL", ReadValue, 256, App.Path & "\config.ini")
  Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
  SelectClass.PTS.Text = ReadValue
  If SelectClass.PTS.Text = "Always" Then
    Announcement.Show
  End If
  AnnouncementLoad.Enabled = False
End Sub

Private Sub ClassDisplay_DblClick()
  If Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") <> "" Then
    Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
  ElseIf Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") = "" Then
    SelectClass.Hide
    MsgBox "加载配置文件时出现错误，" + vbCrLf + "如持续出现该错误，请报告给开发者。", vbOKOnly + vbCritical, "配置文件错误"
    End
  End If
  SelectClass.Show
  Main.Hide
End Sub

Private Sub DataStatisticsButton_Click()
  DataStatistics.Show
  Main.Hide
End Sub

Private Sub Form_Load()
  Randomize
  Main.Icon = Welcome.Icon
  Meta.Version = "3.2.9"
  Unload Welcome
  Meta.WindowState = "Max"
  MinimumProtectCount = 0
  MaxBox.Text = CStr(Meta.MateAmount)
  Main.Caption = "随机数生成器 - RNG - SNAPSHOT " + Meta.Version
  Subtitle.Caption = "SNAPSHOT " + Meta.Version
  ClassDisplay.Caption = "当前已载入 " + Meta.Class + "班 学生数据"
  For i = 1 To 10 Step 1
    MinimumProtect(i) = Int(Rnd() * (Max - Min + 1)) + Min
  Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Dir(App.Path & "\config.ini") <> "" Then
    Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
  End If
  End
End Sub

Private Sub Generate_Click()
  Dim Savetime As Double
  Call ValueGet
  Meta.LastAmount = Val(AmountBox.Text)
  If MinBox.Text = "" Or MaxBox.Text = "" Or AmountBox.Text = "" Then
    If MinBox.Text = "" Then
      MsgBox "尚未填写最小值", vbOKOnly + vbCritical, "参数缺失"
      MinBox.SetFocus
    ElseIf MaxBox.Text = "" Then
      MsgBox "尚未填写最大值", vbOKOnly + vbCritical, "参数缺失"
      MaxBox.SetFocus
    ElseIf AmountBox.Text = "" Then
      MsgBox "尚未填写数量值", vbOKOnly + vbCritical, "参数缺失"
      AmountBox.SetFocus
    End If
  Else
    Min = Val(MinBox.Text)
    Max = Val(MaxBox.Text)
    Meta.Data_GenerateCount = Meta.Data_GenerateCount + Val(AmountBox.Text)
    If AmountBox.Text = "1" Then
      Meta.Protect = False
      If NameHook Then
        Call RandomEvent
      Else
        Call RealRandom
      End If
      Meta.Result(1) = RNMax
      Meta.Data_MateCount(RNMax) = Meta.Data_MateCount(RNMax) + 1
      If NameHook Then
        If OnlyBoyValue Then
          If Meta.Gender(Meta.Result(1)) = "女" Then
            Do While Meta.Gender(Meta.Result(1)) = "女"
              Call RandomEvent
              Meta.Result(1) = RNMax
            Loop
          End If
          Meta.Data_MateCount(RNMax) = Meta.Data_MateCount(RNMax) + 1
          Call DisplayResult
        ElseIf OnlyGirlValue Then
          If Meta.Gender(Meta.Result(1)) = "男" Then
            Do While Meta.Gender(Meta.Result(1)) = "男"
              Call RandomEvent
              Meta.Result(1) = RNMax
            Loop
          End If
          Meta.Data_MateCount(RNMax) = Meta.Data_MateCount(RNMax) + 1
          Call DisplayResult
        ElseIf AllGenderValue Then
          Call DisplayResult
        End If
      Else
        Call DisplayResult
      End If
    Else
      Meta.Amount = Val(AmountBox.Text)
      For i = 1 To Meta.Amount Step 1
        Call RealRandom
        Meta.Result(i) = RNMax
        Meta.Data_MateCount(RNMax) = Meta.Data_MateCount(RNMax) + 1
        If OnlyBoyValue Then
          If Meta.Gender(Meta.Result(i)) = "女" Then
            i = i - 1
          Else
            If Not (AllowDuplicateDataValue) Then
              If i > 1 Then
                For b = 1 To i - 1 Step 1
                  If Meta.Result(i) = Meta.Result(b) Then
                    i = i - 1
                    Exit For
                  End If
                Next b
              End If
            End If
          End If
        ElseIf OnlyGirlValue Then
          If Meta.Gender(Meta.Result(i)) = "男" Then
            i = i - 1
          Else
            If Not (AllowDuplicateDataValue) Then
              If i > 1 Then
                For b = 1 To i - 1 Step 1
                  If Meta.Result(i) = Meta.Result(b) Then
                    i = i - 1
                    Exit For
                  End If
                Next b
              End If
            End If
          End If
        ElseIf AllGenderValue Then
          If Not (AllowDuplicateDataValue) Then
            If i > 1 Then
              For b = 1 To i - 1 Step 1
                If Meta.Result(i) = Meta.Result(b) Then
                  i = i - 1
                  Exit For
                End If
              Next b
            End If
          End If
        End If
      Next i
      Call DisplayResult
    End If
  End If
End Sub

Private Sub Max_Button_Click()
  Meta.WindowLastState = Meta.WindowState
  Meta.WindowState = "Max"
  Call ShowControl
End Sub

Private Sub MaxBox_Change()
  Min = Val(MinBox.Text)
  Max = Val(MaxBox.Text)
  Amount = Val(AmountBox.Text)
  Call ApplyMax
  If Max < Min And MaxBox.Text <> "" Then
    MsgBox "最大值不能比最小值小！", vbOKOnly + vbExclamation, "参数错误"
    MaxBox.Text = CStr(Meta.MateAmount)
    Call ApplyMax
    Exit Sub
  End If
  If Max <= 0 Or Max > Meta.MateAmount Then
    If NameHook.Value = 1 And MaxBox.Text <> "" Then
      MsgBox "在启用姓名挂钩功能时，为防止数据溢出，" + vbCrLf + "只允许最大值设置为[1," + CStr(Meta.MateAmount) + "]的整数。", vbOKOnly + vbExclamation, "参数错误"
      MaxBox.Text = CStr(Meta.MateAmount)
    End If
  End If
  Call ApplyMax
End Sub

Private Sub Min_Button_Click()
  Meta.WindowLastState = Meta.WindowState
  Meta.WindowState = "Min"
  Main.Hide
  MinWindow.Show
End Sub

Private Sub MinBox_Change()
  Min = Val(MinBox.Text)
  Max = Val(MaxBox.Text)
  Amount = Val(AmountBox.Text)
  Call ApplyMin
  If Min > Max Then
    MsgBox "最小值不能比最大值大！", vbOKOnly + vbExclamation, "参数错误"
    MinBox = "1"
    Call ApplyMin
    Exit Sub
  End If
  If Min <= 0 Or Min > Meta.MateAmount Then
    If NameHook.Value = 1 And MinBox.Text <> "" Then
      MsgBox "在启用姓名挂钩功能时，为防止数据溢出，" + vbCrLf + "只允许最小值设置为[1," + CStr(Meta.MateAmount) + "]的整数。", vbOKOnly + vbExclamation, "参数错误"
      MinBox.Text = "1"
    End If
  End If
  Call ApplyMin
End Sub

Private Sub NameHook_Click()
  Call ValueGet
  If NameHookValue Then
    OnlyBoy.Enabled = True
    OnlyGirl.Enabled = True
    AllGender.Enabled = True
    AllGender.Value = True
    FormattingData.Enabled = True
    MinimumProtectSwitch.Enabled = True
    If Val(MinBox.Text) <= 0 Or Val(MinBox.Text) > Meta.MateAmount Then
      MinBox.Text = "1"
    End If
    If Val(MaxBox.Text) <= 0 Or Val(MaxBox.Text) > Meta.MateAmount Then
      MaxBox.Text = CStr(Meta.MateAmount)
    End If
  Else
    OnlyBoy.Enabled = False
    OnlyGirl.Enabled = False
    AllGender.Enabled = False
    OnlyBoy.Value = False
    OnlyGirl.Value = False
    AllGender.Value = False
    FormattingData.Enabled = False
    FormattingData.Value = False
    MinimumProtectSwitch.Value = False
    MinimumProtectSwitch.Enabled = False
  End If
End Sub

Private Sub OnlyBoy_Click()
  Call ValueGet
  Amount = Val(AmountBox.Text)
  If NameHook Then
    If OnlyBoyValue Then
      If Amount > Meta.MaleAmount And Not AllowDuplicateDataValue Then
        AmountBox.Text = CStr(Meta.MaleAmount)
        Exit Sub
      End If
    End If
  End If
End Sub

Private Sub OnlyGirl_Click()
  Call ValueGet
  Amount = Val(AmountBox.Text)
  If NameHook Then
    If OnlyGirlValue Then
      If Amount > Meta.FemaleAmount And Not AllowDuplicateDataValue Then
        AmountBox.Text = CStr(Meta.FemaleAmount)
        Exit Sub
      End If
    End If
  End If
End Sub

Private Sub Subtitle_DblClick()
  Announcement.Show
End Sub

Private Sub ViewLastData_Click()
  If CStr(Meta.Result(1)) = "0" Then
    MsgBox "找不到数据，请先生成一次随机数！", vbOKOnly + vbCritical, "找不到数据"
  Else
    Call DisplayResult
  End If
End Sub

Private Sub Window_Display_Min_Click()
  If Meta.WindowLastState = "" Then
    Meta.WindowLastState = "Max"
    Meta.WindowState = "Min"
  Else
    Meta.WindowLastState = Meta.WindowState
    Meta.WindowState = "Min"
  End If
  Main.Hide
  MinWindow.Show
End Sub

Private Sub Window_Display_Normal_Click()
  If Meta.WindowLastState = "" Then
    Meta.WindowLastState = "Max"
    Meta.WindowState = "Normal"
  Else
    Meta.WindowLastState = Meta.WindowState
    Meta.WindowState = "Normal"
  End If
  Call HideControl
End Sub

Private Sub WindowWeight_Click()
  If WindowWeight.Value = 1 Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
  End If
End Sub
