VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QQ炫舞模拟练习机"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   8055
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7680
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   7215
   End
   Begin VB.Label ji 
      BackColor       =   &H0000FF00&
      Height          =   975
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   15
   End
   Begin VB.Label fuhao 
      BackColor       =   &H00808080&
      Caption         =   "QQ炫舞练习机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   10335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "收"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label fuhao2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "CH制作-2014年2月7日14:12:03"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label signbiao 
      BackColor       =   &H00404040&
      Caption         =   "Label3"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "分数：0 Percent：0 great：0 good：0 bad：0 miss：0 共：0"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label sign 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   15
   End
   Begin VB.Label tiao 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim score As Single
Dim howmuch As Single
Dim bei As Long
Dim Percents As Long, Greats As Long, Goods As Long, Bads As Long, Misss As Long, Gong As Long
'自编应用函数
Private Sub aaabac()
fuhao2.Left = 0
Dim abcc
fuhao.Left = fuhao.Left + 975
abcc = fuhao.Left
fuhao2.Left = fuhao.Left - 975
fuhao2.Top = fuhao.Top
fuhao2.Height = fuhao.Height
fuhao2.Width = 975
fuhao2.Caption = Left(fuhao.Caption, 1)
fuhao.Caption = Right(fuhao.Caption, Len(fuhao.Caption) - 1)
fuhao2.Visible = True
'fuhao.Left = fuhao.Left + 960
For a = 1 To 60 Step 2
'Delay 1
fuhao.Left = fuhao.Left - 960 / 30
fuhao2.Left = fuhao2.Left - 960 / 30
DoEvents
Next a
'fuhao2.Visible = False
fuhao2.Caption = ""
fuhao2.Left = 0
fuhao.Left = abcc - 975
DoEvents
End Sub

Private Sub Delay(MillSeconds As Long)
    Dim S As Long
    S = GetTickCount + MillSeconds
    Do
    DoEvents
        If GetTickCount >= S Then Exit Sub
    Loop
End Sub

Private Function 多少个(文本, 字符)
多少个 = (Len(文本) - Len(Replace(文本, 字符, ""))) / Len(字符)
End Function
Private Sub Command1_Click()
Text2.SetFocus
For a = 0 To fuhao.Width Step 5
biao (a)
DoEvents
Next
Percents = 0
Greats = 0
Goods = 0
Bads = 0
Misss = 0
Gong = 0

fuhao.Caption = ""
score = 0
howmuch = 3
For aaa = 1 To 50
howmuch = howmuch + Rnd + 0.25
If howmuch > 15 Then howmuch = 4
    For aa = 1 To Int(howmuch)
    Dim rnds As Single
    rnds = Rnd(1) * (Rnd(1) + 0.5)
    If rnds < 0.25 Then
    fuhao.Caption = fuhao.Caption & "↑"
    ElseIf rnds <= 0.5 Then
    fuhao.Caption = fuhao.Caption & "↓"
    ElseIf rnds <= 0.75 Then
    fuhao.Caption = fuhao.Caption & "←"
    ElseIf rnds <= 1 Then
    fuhao.Caption = fuhao.Caption & "→"
    End If
    Gong = Gong + 1
    DoEvents
    Label2 = "分数：" & Int(score) & " Percent：" & Percents & " great：" & Greats & " good：" & Goods & " bad：" & Bads & " miss：" & Misss & " 共：" & Gong
    Delay 10
    Next aa

    'For a = 0 To 8055 Step 3 '50 - Int(howmuch * 1)
    'tiao.Width = a
    ''Delay 1
    'DoEvents
    'Next a
    wixzx = GetTickCount
    maxtime = howmuch * 200 + 500
    Do
    paac = GetTickCount - wixzx
    tiao.Width = paac / maxtime * 8055
    DoEvents
    Loop Until paac >= maxtime
    
    Dim ass
    ass = Left(fuhao.Caption, 1)
    If ass = "m" Or ass = "g" Or ass = "b" Or ass = "P" Then
    Else
    Misss = Misss + 1
    End If
    
tiao.Width = 0
fuhao.Caption = ""
Text2.Text = ""
If fuhao.Caption <> "" Then
    Label1 = "miss!"
    bei = 0
    Misss = Misss + 1
End If

Next aaa
Command1.Enabled = False
fuhao.Caption = "游戏结束！本游戏由曹鸿制作！"
fuhao.AutoSize = True
For a = 0 To fuhao.Width
fuhao.Left = fuhao.Left - 50
DoEvents
Delay 1
If fuhao.Left + fuhao.Width < 0 Then Exit For
Next
fuhao.AutoSize = False
fuhao.Left = 975
Command1.Enabled = True
End Sub


Private Sub Form_Load()
Randomize

'3833
'3848
tiao.Width = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Change()
fuhao.Caption = Label1.Caption
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'fuhao.Caption = Replace(fuhao.Caption, "miss!", "")
If KeyCode = 37 Then
If Left(fuhao.Caption, 1) = "←" Then
Text2.Text = Text2.Text & "←"
aaabac
ElseIf fuhao.Caption <> "" Then
Label1 = "miss!"
fuhao.Caption = Text2.Text & fuhao.Caption
Text2.Text = ""
End If
End If
If KeyCode = 38 Then
If Left(fuhao.Caption, 1) = "↑" Then
Text2.Text = Text2.Text & "↑"
aaabac
ElseIf fuhao.Caption <> "" Then
Label1 = "miss!"
fuhao.Caption = Text2.Text & fuhao.Caption
Text2.Text = ""
End If
End If
If KeyCode = 39 Then
If Left(fuhao.Caption, 1) = "→" Then
Text2.Text = Text2.Text & "→"
aaabac
ElseIf fuhao.Caption <> "" Then
Label1 = "miss!"
fuhao.Caption = Text2.Text & fuhao.Caption
Text2.Text = ""
End If
End If
If KeyCode = 40 Then
If Left(fuhao.Caption, 1) = "↓" Then
Text2.Text = Text2.Text & "↓"
aaabac
ElseIf fuhao.Caption <> "" Then
Label1 = "miss!"
fuhao.Caption = Text2.Text & fuhao.Caption
Text2.Text = ""
End If
End If
If KeyCode = 32 Then
    If fuhao.Caption = "" Then
    'Text2 = tiao.Width
    '6720
    '8055'''''''''''''''''''''''''''{6420'''''''''[6570''(6670''|6720|''6770)''6820]''''''8055}
    
        If tiao.Width <= 6770 And tiao.Width >= 6670 Then
            bei = bei + 1
            Label1 = "Percect!  X" & bei
            fuhao.Caption = Label1.Caption
            Percents = Percents + 1
        ElseIf (tiao.Width < 6670 And tiao.Width > 6570) Or (tiao.Width < 6820 And tiao.Width > 6770) Then
            Label1 = "great!"
            fuhao.Caption = Label1.Caption
            bei = 0
            Greats = Greats + 1
        ElseIf (tiao.Width <= 6570 And tiao.Width >= 6420) Or (tiao.Width <= 6920 And tiao.Width >= 6820) Then
            Label1 = "good!"
            fuhao.Caption = Label1.Caption
            bei = 0
            Goods = Goods + 1
        ElseIf (tiao.Width < 6420 And tiao.Width >= 6220) Or (tiao.Width < 7120 And tiao.Width > 6920) Then
            Label1 = "bad!"
            fuhao.Caption = Label1.Caption
            bei = 0
            Bads = Bads + 1
        Else
            Label1 = "miss!"
            fuhao.Caption = Label1.Caption
            bei = 0
            Misss = Misss + 1
        End If
        If tiao.Width > 6720 Then
            score = score + (6720 - (tiao.Width - 6720)) * howmuch / 20 * (bei + 1)
            Label2 = "分数：" & Int(score) & " Percent：" & Percents & " great：" & Greats & " good：" & Goods & " bad：" & Bads & " miss：" & Misss & " 共：" & Gong
        ElseIf tiao.Width <= 6720 Then
            score = score + (6720 - (6720 - tiao.Width)) * howmuch / 20 * (bei + 1)
            Label2 = "分数：" & Int(score) & " Percent：" & Percents & " great：" & Greats & " good：" & Goods & " bad：" & Bads & " miss：" & Misss & " 共：" & Gong
        End If
            Text2.Text = ""
        Else
    Label1 = "miss!"
    fuhao.Caption = Label1.Caption
    Misss = Misss + 1
    End If
    biao (tiao.Width)
End If
End Sub
Private Sub biao(widths)
Me.Tag = 0
signbiao.Top = 0
signbiao.Left = widths
signbiao.Visible = True
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Me.Tag = Me.Tag + 1
Select Case Me.Tag
Case 1
signbiao.BackColor = &HFF00&
Case 2
signbiao.BackColor = &HC000&
Case 3
signbiao.BackColor = &H4000&
Case 4
signbiao.BackColor = &H404040
Case 5
signbiao.BackColor = &H0&
Case 6
signbiao.BackColor = &H0&
Case 7
signbiao.BackColor = &H404040
Case 8
signbiao.BackColor = &H4000&
Case 9
signbiao.BackColor = &HC000&
Case 10
signbiao.BackColor = &HFF00&
Me.Tag = 1
signbiao.Visible = False
signbiao.BackColor = &H8000000F
Timer1.Enabled = False
End Select
End Sub
