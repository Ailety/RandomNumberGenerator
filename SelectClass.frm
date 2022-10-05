VERSION 5.00
Begin VB.Form SelectClass 
   BackColor       =   &H8000000E&
   Caption         =   "�����ʼ��: ��ѡ����İ༶"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5205
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox PTS 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Confirm 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "SelectClass.frx":0000
      Left            =   1440
      List            =   "SelectClass.frx":000D
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
        ReadString = GetPrivateProfileString(CStr(Meta.Class), "MateMale", "NULL", ReadValue, 256, App.Path & "\config.ini")
        Meta.MaleAmount = Val(ReadValue)
        ReadString = GetPrivateProfileString(CStr(Meta.Class), "MateFemale", "NULL", ReadValue, 256, App.Path & "\config.ini")
        Meta.FemaleAmount = Val(ReadValue)
        For i = 1 To Meta.MateAmount Step 1
          ReadString = GetPrivateProfileString(Meta.Class, "MateName(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
          PTS.Text = ReadValue
          Meta.Name(i) = PTS.Text
          ReadString = GetPrivateProfileString(Meta.Class, "MateGender(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
          PTS.Text = ReadValue
          Meta.Gender(i) = PTS.Text
        Next i
        Randomize
        Main.MaxBox.Text = CStr(Meta.MateAmount)
        Main.ClassDisplay.Caption = "��ǰ������ " + Meta.Class + "�� ѧ������"
        For i = 1 To Meta.MateAmount Step 1
          Meta.Data_MateCount(i) = 0
        Next i
        Main.AmountBox.Text = "1"
        If Dir(App.Path & "\config.ini") <> "" Then
          Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
        End If
        MsgBox "������������ɣ�", vbOKOnly, "�������"
        SelectClass.Hide
        Main.Show
        Main.SetFocus
      Else
        MsgBox "���������������Կ����" + vbCrLf + "����ϵ�������Ի�ȡ���°༶������Կ��", vbOKOnly + vbCritical, "������Կ����"
      End If
    Else
      MsgBox "��ǰ���ص������Ѿ���" + Meta.Class + "�࣡", vbOKOnly + vbExclamation, "�Ѽ���" + Meta.Class + "�༶����"
    End If
  Else
    MsgBox "��ѡ��һ���༶���ڳ�������༶���ã�" + vbCrLf + "ȷ��֮��ÿ�������������м��ظð༶���á�", vbOKOnly + vbExclamation, "��ѡ��һ���༶"
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
  Else
    Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
    Cancel = -1
    SelectClass.Hide
    Main.Show
    Main.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Dim WriteString As Long
  Dim ReadString As Long
  Dim ReadValue As String
  SelectClass.Icon = Welcome.Icon
  Unload Welcome
  If Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") <> "" Then
    Name App.Path & "\Meta.vbd" As App.Path & "\config.ini"
  End If
  If Dir(App.Path & "\config.ini") = "" And Dir(App.Path & "\Meta.vbd") = "" Then
    SelectClass.Hide
    MsgBox "���������ļ�ʱ���ִ���" + vbCrLf + "��������ָô����뱨��������ߡ�", vbOKOnly + vbCritical, "�����ļ�����"
    End
  End If
  If Dir(App.Path & "\config.ini") <> "" Then
    ReadValue = String(255, 0)
    ReadString = GetPrivateProfileString("Application_Data", "DefaultClass", "NULL", ReadValue, 256, App.Path & "\config.ini")
    PTS.Text = ReadValue
    If PTS.Text = "" Then
      Exit Sub
    End If
    DefaultClass = JieMi(ReadValue)
    If Mid(DefaultClass, 1, 4) = "2008" Or Mid(DefaultClass, 1, 4) = "2009" Or Mid(DefaultClass, 1, 4) = "2024" Then
      SelectClass.Hide
      Meta.Class = CStr(Val(DefaultClass))
      ReadString = GetPrivateProfileString(DefaultClass, "MateAmount", "NULL", ReadValue, 256, App.Path & "\config.ini")
      Meta.MateAmount = Val(ReadValue)
      ReadString = GetPrivateProfileString(DefaultClass, "MateMale", "NULL", ReadValue, 256, App.Path & "\config.ini")
      Meta.MaleAmount = Val(ReadValue)
      ReadString = GetPrivateProfileString(DefaultClass, "MateFemale", "NULL", ReadValue, 256, App.Path & "\config.ini")
      Meta.FemaleAmount = Val(ReadValue)
      For i = 1 To Meta.MateAmount Step 1
        ReadString = GetPrivateProfileString(DefaultClass, "MateName(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
        PTS.Text = ReadValue
        Meta.Name(i) = PTS.Text
        ReadString = GetPrivateProfileString(DefaultClass, "MateGender(" + CStr(i) + ")", "NULL", ReadValue, 256, App.Path & "\config.ini")
        PTS.Text = ReadValue
        Meta.Gender(i) = PTS.Text
      Next i
      Name App.Path & "\config.ini" As App.Path & "\Meta.vbd"
      Savetime = timeGetTime
      While timeGetTime < Savetime + 200
      DoEvents
      Wend
      Main.Show
      Main.SetFocus
    Else
      WriteString = WritePrivateProfileString("Application_Data", "DefaultClass", "", App.Path & "\config.ini")
      MsgBox "�༶��������" + vbCrLf + "��������ʧ�ܣ�����������ļ��д���İ༶���ݡ�", vbOKOnly + vbCritical, "��������"
      Exit Sub
    End If
  End If
End Sub
