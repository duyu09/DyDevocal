VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Duyu - WAV或MP3音频文件 消除人声软件"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6240
   ScaleWidth      =   8790
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1920
      Picture         =   "Form1.frx":048A
      ScaleHeight     =   1935
      ScaleWidth      =   3495
      TabIndex        =   13
      ToolTipText     =   "请将WAV或MP3格式音频文件拖放到此处"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7440
      Top             =   5520
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清除"
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
      Left            =   8040
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   7200
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "一键消除人声"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      TabIndex        =   3
      ToolTipText     =   "单击可以一键消除人声"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "请将WAV或MP3格式音频文件拖放到此处"
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   855
      Left            =   1920
      TabIndex        =   12
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2021"
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
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Duyu - WAV或MP3音频文件 消除人声软件 - 单击查看版权声明"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "错误信息："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Duyu - wav或mp3音频文件 消除人声软件"
      BeginProperty Font 
         Name            =   "华文隶书"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Duyu - WAV或MP3音频文件 消除人声软件 - 单击查看版权声明"
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "加载完毕"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "正在处理："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpszOp As String, _
ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal LpszDir As String, ByVal FsShowCmd As Long) _
As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
lStructSize As Long
hwndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Dim arg As String, a As Integer, rannum As Long
 

If List1.ListCount = 0 Then
  MsgBox "请添加文件。", 48
  Command1.Enabled = True
  Command2.Enabled = True
  Command4.Enabled = True
  Exit Sub
End If
Label2.Caption = "正在处理"
For a = 1 To List1.ListCount

 Randomize
 rannum = 10000 * Rnd()
Text1.Text = List1.List(a - 1)
  arg = Chr(34) & Text1.Text & Chr(34) & " " & rannum
  DoEvents
  ShellExecute Me.hwnd, "open", "Bridge.exe", arg, App.Path, 0
   Do While (Dir(Environ("temp") & "\" & rannum & "_finished") = "")
   DoEvents
   Loop
  DoEvents
  Sleep (1111)
  On Error Resume Next
  Kill Environ("temp") & "\" & rannum & "_finished"
  DoEvents
      Sleep (1111)
      DoEvents
      If Dir(Text1.Text & "__WAVfile.wav_Output.wav") = "" Then
      List2.AddItem "警告：" & Text1.Text & "消音失败！"
      On Error Resume Next
      Kill Environ("temp") & "\" & rannum & "_finished"
      GoTo Erro
      End If
      Sleep (1111)
      On Error Resume Next
  Kill Environ("temp") & "\" & rannum & "_finished"
  DoEvents
  
  On Error Resume Next
  Kill Text1.Text & "_WAVfile.wav"
Erro:
  On Error Resume Next
  ProgressBar1.Value = CInt(a * 100 / List1.ListCount)
  Label2.Caption = "正在处理 " & CInt(a * 100 / List1.ListCount) & "%"
Next a
Command1.Enabled = True
  Command2.Enabled = True
Command4.Enabled = True
Label2.Caption = "处理完毕"
ProgressBar1.Value = 0
List1.Clear
End Sub

Private Sub Command4_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub

Private Sub Command2_Click()
Dim ofn As OPENFILENAME
Dim rtn As String, fp As String
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Me.hwnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "所有文件(*.*)"
ofn.lpstrFile = Space(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = Environ("homepath") & "\Music"
ofn.lpstrTitle = "请选择被处理的文件"
ofn.flags = 6148
rtn = GetOpenFileName(ofn)
If rtn >= 1 Then
List1.AddItem ofn.lpstrFile
End If
End Sub

Private Sub Command3_Click()
Form2.Text1.Text = Me.Text1.Text
Form2.Show
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For Each file In Data.Files
If Err.Number > 0 Then
ss = MsgBox(Err.Description, 48)
Exit Sub
End If
List1.AddItem file
Next
End Sub
Private Sub Form_Load()
Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
Me.Height = 10000
 'If Dir("lame.exe") = "" Or Dir("Dy_DeCore.exe") = "" Or Dir("Bridge.exe") = "" Then
 'MsgBox "缺少组件，无法运行！", 48
 'End
 'End If
Load Form3
SendMessage List1.hwnd, &H194, 3000, ByVal 0
SendMessage List2.hwnd, &H194, 3000, ByVal 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell "taskkill /f /im Bridge.exe", vbHide
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Width = 9030
If Me.Height < 5300 Then
 Me.Height = 5300
End If
List1.Height = Abs(Me.Height - 2580)
Me.ProgressBar1.Top = Abs(Me.Height - 1100)
Label2.Top = Abs(Me.Height - 1100)
Command1.Top = Abs(Me.Height - 2000)
List2.Height = Abs(Me.Height - 4700)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell "taskkill /f /im Bridge.exe", vbHide
End
End Sub

Private Sub Label3_Click()
Form3.Show
End Sub

Private Sub Label5_Click()
Form3.Show
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Load Form4
Form4.Label1.Caption = List1.Text
Form4.WindowsMediaPlayer1.URL = Form4.Label1.Caption
Form4.WindowsMediaPlayer1.Controls.stop
Form4.Label2.Caption = "标题：" & Form4.WindowsMediaPlayer1.currentMedia.getItemInfo("title")
Form4.Label3.Caption = "艺术家：" & Form4.WindowsMediaPlayer1.currentMedia.getItemInfo("Author")
Form4.Label4.Caption = "媒体内容描述：" & Form4.WindowsMediaPlayer1.currentMedia.getItemInfo("Description")
Form4.Label5.Caption = "持续时间：" & Form4.WindowsMediaPlayer1.currentMedia.getItemInfo("Duration") & " 秒"
Form4.Label6.Caption = "媒体类型：" & Form4.WindowsMediaPlayer1.currentMedia.getItemInfo("filetype")
Form4.Caption = "媒体信息 - " & Form4.Label1.Caption
On Error Resume Next
Form4.WindowsMediaPlayer1.enableContextMenu = False
Form4.Show
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.Enabled = False Then
 Exit Sub
 End If
On Error Resume Next
For Each file In Data.Files
If Err.Number > 0 Then
ss = MsgBox(Err.Description, 48)
Exit Sub
End If
List1.AddItem file
Next
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.Enabled = False Then
 Exit Sub
 End If
On Error Resume Next
For Each file In Data.Files
If Err.Number > 0 Then
ss = MsgBox(Err.Description, 48)
Exit Sub
End If
List1.AddItem file
Next
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProgressBar1.ToolTipText = Label2.Caption
End Sub
Private Sub Text1_GotFocus()
Command2.SetFocus
End Sub

Private Sub Timer1_Timer()
Picture1.Top = (List1.Height - Picture1.Height) / 2 + 1440
If List1.ListCount = 0 Then
Picture1.Visible = True
Else
Picture1.Visible = False
End If
End Sub
