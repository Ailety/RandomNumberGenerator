VERSION 5.00
Begin VB.Form ResultForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4815
   StartUpPosition =   1  '����������
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
      Left            =   1560
      TabIndex        =   1
      Top             =   5160
      Width           =   1695
   End
   Begin VB.ListBox Result 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "ResultForm.frx":0000
      Left            =   360
      List            =   "ResultForm.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label TimeDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ʱ: NULL ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   4815
   End
End
Attribute VB_Name = "ResultForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim NowTime

Public Function GetUnixTime_ms() As String
    GetUnixTime_ms = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(timeGetTime, 3)
End Function

Private Sub Confirm_Click()
  Unload ResultForm
End Sub

Private Sub Form_Load()
  Dim ResultData As String
  ResultForm.Icon = Main.Icon
  SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
  If Main.AmountBox.Text = "1" Then
    If Main.NameHook.Value Then
      If Main.FormattingData.Value Then
        If Len(Meta.Name(Meta.Result(1))) = 2 Then
          ResultData = Mid(Meta.Name(Meta.Result(1)), 1, 1) + "��" + Mid(Meta.Name(Meta.Result(1)), 2, 1)
        Else
          ResultData = Meta.Name(Meta.Result(1))
        End If
        If Meta.Result(1) < 10 Then
          Result.AddItem "0" + CStr(Meta.Result(1)) + "�� " + ResultData
        Else
          Result.AddItem CStr(Meta.Result(1)) + "�� " + ResultData
        End If
        If Meta.Protect Then
          Result.AddItem "�����������ڱ��γ�ȡ����Ч"
        End If
      Else
        Result.AddItem CStr(Meta.Result(1)) + "�� " + Meta.Name(Meta.Result(1))
        If Meta.Protect Then
          Result.AddItem "�����������ڱ��γ�ȡ����Ч"
        End If
      End If
    Else
      Result.AddItem Meta.Result(1)
    End If
  Else
    If Main.NameHook.Value Then
      For i = 1 To Meta.LastAmount Step 1
        If Main.FormattingData.Value Then
          If Len(Meta.Name(Meta.Result(i))) = 2 Then
            ResultData = Mid(Meta.Name(Meta.Result(i)), 1, 1) + "��" + Mid(Meta.Name(Meta.Result(i)), 2, 1)
          Else
            ResultData = Meta.Name(Meta.Result(i))
          End If
          If Meta.Result(i) < 10 Then
            Result.AddItem "0" + CStr(Meta.Result(i)) + "�� " + ResultData
          Else
            Result.AddItem CStr(Meta.Result(i)) + "�� " + ResultData
          End If
        Else
          Result.AddItem CStr(Meta.Result(i)) + "�� " + Meta.Name(Meta.Result(i))
        End If
      Next i
    Else
      For i = 1 To Meta.Amount Step 1
        Result.AddItem CStr(Meta.Result(i))
      Next i
    End If
  End If
  NowTime = GetUnixTime_ms()
  Cache = (NowTime - Meta.GenerateTime) / 1000
  If Left(Cache, 1) = "." Then
    Dim Cache2
    Cache2 = Mid(Cache, 2, Len(Cache) - 1)
    If Left(Cache2, 1) = "0" Then
      Cache2 = Val(Cache2)
    End If
    TimeDisplay.Caption = "��ʱ: " & Cache2 & "����"
  ElseIf Left(Cache, 1) = "-" Then
    Cache2 = Val(Mid(Cache, 3, Len(Cache) - 2))
    TimeDisplay.Caption = "��ʱ: " & Cache2 & "����"
  ElseIf Cache = "0" Then
    TimeDisplay.Caption = "��ʱ: 0����"
  Else
    TimeDisplay.Caption = "��ʱ: " & Cache & "��"
  End If
End Sub
