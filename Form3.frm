VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duyu - WAV或MP3音频文件_消除人声软件版权声明"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6600
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "试用wav转mp3组件"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "学校官网:https://www.lcez.cn/"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "保留所有权利"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "济南市历城第二中学55级31班 杜宇"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2385
      Left            =   240
      Picture         =   "Form3.frx":000C
      Top             =   240
      Width           =   2310
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpszOp As String, _
ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal LpszDir As String, ByVal FsShowCmd As Long) _
As Long
Private Sub Command1_Click()
Form3.Hide
End Sub

Private Sub Command2_Click()
  DoEvents
  ShellExecute Me.hwnd, "open", "Dy_WPM.exe", vbNullString, App.Path, 1
End Sub

Private Sub Form_Load()
Label4.Caption = Form1.Caption
End Sub

Private Sub Label3_Click()
On Error Resume Next
  ShellExecute Me.hwnd, "open", "https://www.lcez.cn/", vbNullString, App.Path, 1
End Sub
