VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "课堂提问随机产生器-513"
   ClientHeight    =   3444
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10212
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   15.6
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3444
   ScaleWidth      =   10212
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   8400
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7080
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "下一个"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Integer
Dim a(100) As String
Dim b(100) As Boolean
Private Sub Command1_Click()
    Timer2.Enabled = True
    Timer1.Enabled = True
    
End Sub

Private Sub Command3_Click()
    For i = 0 To l - 1
        b(i) = True
    Next i
End Sub

Private Sub Timer1_Timer()
    x = Int(Rnd * l)
    If b(x) Then Label1.Caption = a(x)
        
End Sub

Private Sub Form_Load()
    Randomize
    Open "student.txt" For Input As #1
    l = 0
    Do Until EOF(1)
        Line Input #1, a(l)
        l = l + 1
    Loop
    For i = 0 To l - 1
        b(i) = True
    Next i
End Sub

Private Sub Timer2_Timer()
        Timer1.Enabled = False
        For i = 0 To l - 1
            If Label1.Caption = a(i) Then b(i) = False
        Next i
        Timer2.Enabled = False
End Sub
