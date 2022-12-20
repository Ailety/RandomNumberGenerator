VERSION 5.00
Begin VB.Form DataStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ͳ��"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12810
   StartUpPosition =   2  '��Ļ����
   Begin VB.ListBox DataList_Part 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   7080
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Timer DataDisplay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   7440
   End
   Begin VB.Timer DataLoad 
      Interval        =   25
      Left            =   120
      Top             =   7920
   End
   Begin VB.CommandButton Back 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
   Begin VB.ListBox DataList 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label DataLoadLabel 
      Alignment       =   2  'Center
      Caption         =   "���ݼ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3200
      Width           =   1215
   End
End
Attribute VB_Name = "DataStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadCount As Integer

Private Sub Back_Click()
  Main.Show
  Unload Me
End Sub

Private Sub DataDisplay_Timer()
  If LoadCount <= Meta.MateAmount Then
    If Len(Meta.Name(LoadCount)) = 2 Then
      DataName = Mid(Meta.Name(LoadCount), 1, 1) + "��" + Mid(Meta.Name(LoadCount), 2, 1)
    Else
      DataName = Meta.Name(LoadCount)
    End If
    If LoadCount < 10 Then
      DataList.AddItem "0" + CStr(LoadCount) + "�� " + DataName + " " + Meta.Gender(LoadCount) + " ���Ʊ�����: " + CStr(Meta.Data_MateCount(LoadCount)) + "��"
    Else
      DataList.AddItem CStr(LoadCount) + "�� " + DataName + " " + Meta.Gender(LoadCount) + " ���Ʊ�����: " + CStr(Meta.Data_MateCount(LoadCount)) + "��"
    End If
    LoadCount = LoadCount + 1
  Else
    DataDisplay.Enabled = False
  End If
End Sub

Private Sub DataLoad_Timer()
  DataLoadLabel.Visible = False
  DataList.Visible = True
  DataList_Part.Visible = True
  DataList.AddItem "�������ڱ��������ڼ乲������ " + CStr(Meta.Data_GenerateCount) + " ��"
  DataList.AddItem "�������������ݵ����:"
  Max = 0
  If Meta.Data_GenerateCount > 0 Then
    For i = 1 To Meta.MateAmount Step 1
      If Meta.Data_MateCount(i) >= Max Then
        Max = Meta.Data_MateCount(i)
        MaxIndex = i
      End If
    Next i
    DataList.AddItem ""
    DataList.AddItem Meta.Name(MaxIndex) + "�����еĴ�����࣬Ϊ " + CStr(Max) + " ��"
    If Meta.Data_GenerateCount <= Meta.MateAmount Then
      DataList.AddItem "ע: �����ɴ�������ʱ�����ֵ���߱��ο���ֵ"
    End If
    DataList.AddItem ""
  Else
    DataList.AddItem ""
    DataList.AddItem "��δ�����������ҹ���ͬʱ��ȡ����"
    DataList.AddItem "�޷���ʾ������������ݡ�"
    DataList.AddItem ""
  End If
  DataList_Part.AddItem Meta.Class + "�༶����ͳ��"
  DataList_Part.AddItem ""
  DataList_Part.AddItem "��������: " + CStr(Meta.MaleAmount)
  DataList_Part.AddItem "Ů������: " + CStr(Meta.FemaleAmount)
  DataList_Part.AddItem ""
  DataList_Part.AddItem "�༶������: " + CStr(Meta.MateAmount)
  DataDisplay.Enabled = True
  DataLoad.Enabled = False
End Sub

Private Sub Form_Load()
  Me.Icon = Main.Icon
  LoadCount = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Main.Show
End Sub
