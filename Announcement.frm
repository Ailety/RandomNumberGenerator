VERSION 5.00
Begin VB.Form Announcement 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���¹���: Ver. �汾 NULLNULL"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   4455
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer AnnouncementLoad 
      Interval        =   25
      Left            =   4080
      Top             =   6960
   End
   Begin VB.CheckBox NoLongerRemind 
      BackColor       =   &H8000000E&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
  Announcement.Caption = "���¹���: " + Meta.Version + " 20221119"
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
  AnnouncementMessageLoadAmount = 1
  AnnouncementMessageAmount = 38
  AnnouncementMessage(1) = "1.1.2 - 2.7.2 ��������"
  AnnouncementMessage(2) = ""
  AnnouncementMessage(3) = "1.��д����������߼�"
  AnnouncementMessage(4) = "2.�޸���14��Ӱ�������BUG"
  AnnouncementMessage(5) = "3.�Ż����ִ����߼�"
  AnnouncementMessage(6) = "4.��д�༶���ݹ��ܣ�����֧��2008��"
  AnnouncementMessage(7) = "2009 ��2024�༶"
  AnnouncementMessage(8) = "5.����������������ʦʹ�õĹ���: ǿ�Ƹ������ݡ���С����ʾ����"
  AnnouncementMessage(9) = "6.�Խ�������˵������Ż�"
  AnnouncementMessage(10) = "7.�޸����ڳ����������л��༶ʱż��"
  AnnouncementMessage(11) = "�ļ��޷��ҵ��ķǳ�����Ĵ�����Ϣ"
  AnnouncementMessage(12) = "8.�Ż���������������ʾ�ı���С"
  AnnouncementMessage(13) = ""
  AnnouncementMessage(14) = "2.7.2 - 3.1.2 ��������"
  AnnouncementMessage(15) = ""
  AnnouncementMessage(16) = "1.��д�������Ź��ܣ���֧������ģʽ: �������(Ĭ��)���������Ż���������С��"
  AnnouncementMessage(17) = "2.�Ľ������ı�"
  AnnouncementMessage(18) = "3.������ӭ����"
  AnnouncementMessage(19) = "4.������UI�����Ż�"
  AnnouncementMessage(20) = "5.�޸��ˡ����Ź����л����濨�١�������󡱡�����������ɹ���ż�����ݴ��󡱡������ֹ���ʵ�����������������������"
  AnnouncementMessage(21) = "6.�Բ��ֹ��ܽ������߼��Ż�"
  AnnouncementMessage(22) = "7.�������ظ���������жϣ�����ֻ��������һ�������Ա������ݰ�ȫ"
  AnnouncementMessage(23) = "8.��Ե�����������ɽ������߼����Ż�����������10�κ�ÿһ�ζ����������ǰ10�α����е��ˡ�"
  AnnouncementMessage(24) = "9.����������ͳ�ƹ���"
  AnnouncementMessage(25) = ""
  AnnouncementMessage(26) = "3.1.2 - 3.2.3 ��������"
  AnnouncementMessage(27) = ""
  AnnouncementMessage(28) = "1.�޸��˰༶ѡ�񴰿�ż�ּ����쳣�������������λ������"
  AnnouncementMessage(29) = "2.�Ż��˱��׻��Ƶ����û��ƣ������������ÿ���"
  AnnouncementMessage(30) = ""
  AnnouncementMessage(31) = "3.2.3 - 3.2.9 �������� "
  AnnouncementMessage(32) = ""
  AnnouncementMessage(33) = "1.�޸��˴������Ź����򲿷ִ����߼��쳣�����л�����ʱ��ʾλ�ô�λ������ - 20221014"
  AnnouncementMessage(34) = "2.�޸��˲鿴�ϴ����ݹ�����Ϊ�������ɴ������µ��±�Խ����� - 20221022"
  AnnouncementMessage(35) = "3.�޸�������Сֵ�����ֵ��ͬ���µĳ�ȡ�����߼�����鵽��Ӧ�ó鵽��ͬѧ��- 20221108"
  AnnouncementMessage(36) = "4.�޸����Ա�ɸѡ���ִ����߼������µĲ��ֹ����쳣������ - 20221119"
  AnnouncementMessage(37) = "5.�Ե���༶���ݵĲ��ִ�������˵��� - 20221119"
  AnnouncementMessage(38) = "6.�޸��˸��¹�����ı���ʾ��ʽ��������������εĸ��� - 20221119"
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
