VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  '단일 고정
   Caption         =   "님게임"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdReset 
      Caption         =   "다시"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3개"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2개"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1개"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtInfo 
      Height          =   1215
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   2520
      Width           =   6495
   End
   Begin VB.Label Label13 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "13"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "12"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "11"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "10"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "9"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "8"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "7"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "6"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "5"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "4"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "3"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "2"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "1"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   12
      Left            =   5160
      Shape           =   3  '원형
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   11
      Left            =   4320
      Shape           =   3  '원형
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   10
      Left            =   3480
      Shape           =   3  '원형
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   9
      Left            =   2640
      Shape           =   3  '원형
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   8
      Left            =   1800
      Shape           =   3  '원형
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   7
      Left            =   6360
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   6
      Left            =   5520
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   5
      Left            =   4680
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   4
      Left            =   3840
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   3
      Left            =   3000
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   2
      Left            =   2160
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000002&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   0
      Left            =   480
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000003&
      BackStyle       =   1  '투명하지 않음
      Height          =   495
      Index           =   1
      Left            =   1320
      Shape           =   3  '원형
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rim As Integer
Dim s As Integer

Private Sub cmd1_Click()
    On Error Resume Next
    Randomize Timer
    Ball(s).Visible = False
    s = s + 1
    rim = rim - 1
    txtInfo.Text = "당신이 구슬 한 개를 가져갑니다." & vbCrLf & txtInfo.Text
    If rim = 0 Then
        txtInfo.Text = "당신이 이겼읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        Exit Sub
    End If
    bot = Int(Rnd * 3 + 1)
    For i = s To s + bot - 1
        If i > 12 Then
            txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
            cmd1.Enabled = False
            cmd2.Enabled = False
            cmd3.Enabled = False
            Exit Sub
            Exit For
        End If
        Ball(i).Visible = False
    Next i
    txtInfo.Text = "컴퓨터가 구슬 " & bot & "개를 가져갑니다." & vbCrLf & txtInfo.Text
    s = s + bot
    rim = rim - bot
    If rim = 0 Then
        txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
    End If
End Sub

Private Sub cmd2_Click()
    On Error Resume Next
    Randomize Timer
    For i = s To s + 1
        Ball(i).Visible = False
    Next i
    s = s + 2
    rim = rim - 2
    txtInfo.Text = "당신이 구슬 두 개를 가져갑니다." & vbCrLf & txtInfo.Text
    If rim = 0 Then
        txtInfo.Text = "당신이 이겼읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        Exit Sub
    End If
    bot = Int(Rnd * 3 + 1)
    For i = s To s + bot - 1
        If i > 12 Then
            txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
            cmd1.Enabled = False
            cmd2.Enabled = False
            cmd3.Enabled = False
            Exit Sub
            Exit For
        End If
        Ball(i).Visible = False
    Next i
    txtInfo.Text = "컴퓨터가 구슬 " & bot & "개를 가져갑니다." & vbCrLf & txtInfo.Text
    s = s + bot
    rim = rim - bot
    If rim = 0 Then
        txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
    End If
End Sub

Private Sub cmd3_Click()
    On Error Resume Next
    Randomize Timer
    For i = s To s + 2
        Ball(i).Visible = False
    Next i
    s = s + 3
    rim = rim - 3
    txtInfo.Text = "당신이 구슬 세 개를 가져갑니다." & vbCrLf & txtInfo.Text
    If rim = 0 Then
        txtInfo.Text = "당신이 이겼읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
        Exit Sub
    End If
    bot = Int(Rnd * 3 + 1)
    For i = s To s + bot - 1
        If i > 12 Then
            txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
            cmd1.Enabled = False
            cmd2.Enabled = False
            cmd3.Enabled = False
            Exit Sub
            Exit For
        End If
        Ball(i).Visible = False
    Next i
    txtInfo.Text = "컴퓨터가 구슬 " & bot & "개를 가져갑니다." & vbCrLf & txtInfo.Text
    s = s + bot
    rim = rim - bot
    If rim = 0 Then
        txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
        cmd1.Enabled = False
        cmd2.Enabled = False
        cmd3.Enabled = False
    End If
End Sub

Private Sub cmdReset_Click()
    Randomize Timer
    cmd1.Enabled = True
    cmd2.Enabled = True
    cmd3.Enabled = True
    
    s = 0
    rim = 13
    For i = 0 To 12
        Ball(i).Visible = True
    Next i
    
    txtInfo.Text = "다시 합니다." & vbCrLf & txtInfo.Text
    
    who = Int(Rnd * 2 + 1)
    If who = 2 Then
        bot = Int(Rnd * 3 + 1)
        For i = s To s + bot - 1
            If i > 12 Then
                txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
                cmd1.Enabled = False
                cmd2.Enabled = False
                cmd3.Enabled = False
                Exit Sub
                Exit For
            End If
            Ball(i).Visible = False
        Next i
        txtInfo.Text = "컴퓨터가 구슬 " & bot & "개를 가져갑니다." & vbCrLf & txtInfo.Text
        s = s + bot
        rim = rim - bot
        If rim = 0 Then
            txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
            cmd1.Enabled = False
            cmd2.Enabled = False
            cmd3.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Randomize Timer
    rim = 13
    s = 0
    
    who = Int(Rnd * 2 + 1)
    If who = 2 Then
        bot = Int(Rnd * 3 + 1)
        For i = s To s + bot - 1
            If i > 12 Then
                txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
                cmd1.Enabled = False
                cmd2.Enabled = False
                cmd3.Enabled = False
                Exit Sub
                Exit For
            End If
            Ball(i).Visible = False
        Next i
        txtInfo.Text = "컴퓨터가 구슬 " & bot & "개를 가져갑니다." & vbCrLf & txtInfo.Text
        s = s + bot
        rim = rim - bot
        If rim = 0 Then
            txtInfo.Text = "당신이 졌읍니다." & vbCrLf & txtInfo.Text
            cmd1.Enabled = False
            cmd2.Enabled = False
            cmd3.Enabled = False
        End If
    End If
End Sub
