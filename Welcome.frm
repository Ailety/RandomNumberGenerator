VERSION 5.00
Begin VB.Form Welcome 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӭ����"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Confirm 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "QQ��"
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
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Way2 
      BackStyle       =   0  'Transparent
      Caption         =   "΢�ţ�"
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
      Left            =   600
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Way6 
      BackStyle       =   0  'Transparent
      Caption         =   "Telegram��"
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
      Left            =   600
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Way4 
      BackStyle       =   0  'Transparent
      Caption         =   "Twitter��"
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
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Way5 
      BackStyle       =   0  'Transparent
      Caption         =   "Discord��"
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
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Way3 
      BackStyle       =   0  'Transparent
      Caption         =   "���䣺"
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
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Way7 
      BackStyle       =   0  'Transparent
      Caption         =   "����Ŀ����github�Ͽ�Դ������ѭ GPL3.0 Э�顣������������Ľ�����ѧϰ�����������ͨ��github�ֿ�����Դ���롣"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
  WelcomeText.Text = "" + vbCrLf + "                                                             ��ӭʹ�������������(RNG)" + vbCrLf + " " + vbCrLf + "�� ����������ʦ�Ͽε����������������ѵ����� SNAPSHOT 3.1.2 (�������հ汾��һ�θ���+�����޶�)������Ҳ����������ơ����ο�������1�ڿΣ������ĸ��º�ά������35.7Сʱ(ʵ�ʿ���ʱ��)��" + vbCrLf + "�� ��ǰ�汾����˺ܶ�����汾�����ڵ�ʹ�㣬ͬʱҲ�޸���99%��BUG�����������ھ����ͼ��������ܴ����ż�Ϊ���Ե�©������ӭ��������Ȼ��������кõĽ��飬Ҳ����������ϵ��������������ơ�"
End Sub

Private Sub OpenBrowser_Click()
  Call OpenWebPage1
End Sub
