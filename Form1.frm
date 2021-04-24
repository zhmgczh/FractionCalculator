VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "分數計算器"
   ClientHeight    =   4575
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6615
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command6 
      Caption         =   "清零"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "繼續計算"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "要計算得數請在“紅色區域”輸入要計算的分數，點擊+、-、*、/四則運算按鈕，即可在黃色區域得到答案。"
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "自動約分，精準永恆！"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "分數計算器——蔡于飛"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   240
      Width           =   5415
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   6000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   2520
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "" Then
  Text2.Text = "1"
End If
If Text4.Text = "" Then
  Text4.Text = "1"
End If
If Text1.Text = "" Or Text3.Text = "" Then
  MsgBox "分子不能為空！", 42, "分數計算器"
  Exit Sub
End If
fm = Text2.Text * Text4.Text
Text6.Text = fm
fz1 = Text1.Text * Text4.Text
fz2 = Text3.Text * Text2.Text
fz = Val(fz1) + Val(fz2)
If fz > fm Then
  m = fz
  n = fm
Else
  m = fm
  n = fz
End If
If n = 0 Then
Text5.Text = "0"
Text6.Text = ""
Else
Do
r = m Mod n
m = n
n = r
Loop Until r = 0
End If
Text5.Text = fz / m
Text6.Text = fm / m
End Sub
Private Sub Command2_Click()
If Text2.Text = "" Then
  Text2.Text = "1"
End If
If Text4.Text = "" Then
  Text4.Text = "1"
End If
If Text1.Text = "" Or Text3.Text = "" Then
  MsgBox "分子不能為空！", 42, "分數計算器"
  Exit Sub
End If
fm = Text2.Text * Text4.Text
fz = Text1.Text * Text3.Text
If fz > fm Then
  m = fz
  n = fm
Else
  m = fm
  n = fz
End If
If n = 0 Then
Text5.Text = "0"
Text6.Text = ""
Else
Do
r = m Mod n
m = n
n = r
Loop Until r = 0
End If
Text5.Text = fz / m
Text6.Text = fm / m
End Sub
Private Sub Command3_Click()
If Text2.Text = "" Then
  Text2.Text = "1"
End If
If Text4.Text = "" Then
  Text4.Text = "1"
End If
If Text1.Text = "" Or Text3.Text = "" Then
  MsgBox "分子不能為空！", 42, "分數計算器"
  Exit Sub
End If
fm = Text2.Text * Text4.Text
fz1 = Text1.Text * Text4.Text
fz2 = Text3.Text * Text2.Text
fz = Val(fz1) - Val(fz2)
If fz > fm Then
  m = fz
  n = fm
Else
  m = fm
  n = fz
End If
If n = 0 Then
Text5.Text = "0"
Text6.Text = ""
Else
Do
r = m Mod n
m = n
n = r
Loop Until r = 0
End If
Text5.Text = fz / m
Text6.Text = fm / m
End Sub
Private Sub Command4_Click()
If Text2.Text = "" Then
  Text2.Text = "1"
End If
If Text4.Text = "" Then
  Text4.Text = "1"
End If
If Text1.Text = "" Or Text3.Text = "" Then
  MsgBox "分子不能為空！", 42, "分數計算器"
  Exit Sub
End If
fz = Text1.Text * Text4.Text
fm = Text2.Text * Text3.Text
If fz > fm Then
  m = fz
  n = fm
Else
  m = fm
  n = fz
End If
If n = 0 Then
Text5.Text = "0"
Text6.Text = ""
Else
Do
r = m Mod n
m = n
n = r
Loop Until r = 0
End If
Text5.Text = fz / m
Text6.Text = fm / m
End Sub
Private Sub Command5_Click()
Text1.Text = Text5.Text
Text2.Text = Text6.Text
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
Private Sub Command6_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
Private Sub Form_Load()
Text1.Alignment = vbCenter
Text2.Alignment = vbCenter
Text3.Alignment = vbCenter
Text4.Alignment = vbCenter
Text5.Alignment = vbCenter
Text6.Alignment = vbCenter
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
  Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
  KeyAscii = 0
End If
End Sub
