VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "智能聊天机器人"
   ClientHeight    =   9735
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   11205
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox ListDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4365
      Left            =   9120
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.PictureBox PicTalk 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   -20
      Picture         =   "FrmMain.frx":324A
      ScaleHeight     =   9495
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   -20
      Width           =   9135
      Begin VB.PictureBox PicTeach 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   9015
         TabIndex        =   8
         Top             =   7440
         Width           =   9015
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H80000009&
            Caption         =   "编辑数据库(&E)"
            Height          =   375
            Index           =   3
            Left            =   6840
            TabIndex        =   17
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H80000009&
            Caption         =   "退出调教(&C)"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   16
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H80000009&
            Caption         =   "保存修改(&O)"
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   15
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox TxtTeach 
            Appearance      =   0  'Flat
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
            Index           =   2
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   7455
         End
         Begin VB.TextBox TxtTeach 
            Appearance      =   0  'Flat
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
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label LabTip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "说明：多条回复请用“\”分隔."
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   1080
            Width           =   2460
         End
         Begin VB.Label LabTip2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "智能回复："
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label LabTip1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关键字词："
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.CommandButton CmdSend 
         BackColor       =   &H80000009&
         Caption         =   "发送(&Send)"
         Default         =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   6920
         Width           =   1215
      End
      Begin VB.CommandButton CmdTeach 
         BackColor       =   &H8000000B&
         Caption         =   "调教(&Teach)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3390
         TabIndex        =   4
         Top             =   6920
         Width           =   1215
      End
      Begin VB.TextBox TxtSend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   5400
         Width           =   5895
      End
      Begin VB.TextBox TxtTalk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Timer TimerActive 
         Interval        =   7000
         Left            =   6240
         Top             =   6240
      End
      Begin VB.Timer TimerRevert 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6720
         Top             =   6240
      End
      Begin VB.Label LabRevise 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "更新日期：2014/9/5"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6830
         TabIndex        =   27
         ToolTipText     =   "修改人员：5311011lmq & moqiaoduo"
         Top             =   6880
         Width           =   2040
      End
      Begin VB.Shape ShapeHead 
         BorderColor     =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   735
      End
      Begin VB.Line LCloseR 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   8640
         X2              =   8760
         Y1              =   240
         Y2              =   120
      End
      Begin VB.Line LCloseL 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   8640
         X2              =   8760
         Y1              =   120
         Y2              =   240
      End
      Begin VB.Line LMin 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   8280
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label LabClose 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
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
         Left            =   8520
         TabIndex        =   26
         ToolTipText     =   "关闭"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label LabMin 
         BackColor       =   &H002CF900&
         BackStyle       =   0  'Transparent
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
         Left            =   8160
         TabIndex        =   25
         ToolTipText     =   "最小化"
         Top             =   0
         Width           =   375
      End
      Begin VB.Image ImgPortrait 
         Height          =   735
         Left            =   120
         Picture         =   "FrmMain.frx":B8DA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LabRemark 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   6360
         TabIndex        =   24
         Top             =   3720
         Width           =   2460
      End
      Begin VB.Label LabRemarks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "备注："
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   3315
         Width           =   540
      End
      Begin VB.Label LabSmartWord 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "数据：？"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   2860
         Width           =   720
      End
      Begin VB.Label LabAuthor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "作者：？"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label LabFrom 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "来自：？"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label LabBirth 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "生日：？"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   1515
         Width           =   720
      End
      Begin VB.Label LabSex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "性别：？"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label LabRTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "[在线]"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   585
         Width           =   480
      End
      Begin VB.Label LabRName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   225
         Width           =   240
      End
      Begin VB.Label LabStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   5055
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long  '圆角矩形
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long  '圆角矩形
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long  '圆角矩形
Private Type RECT  '圆角矩形
    Left As Long  '圆角矩形
    Top As Long  '圆角矩形
    Right As Long  '圆角矩形
    Bottom As Long  '圆角矩形
End Type  '圆角矩形
Dim xx As Long, yy As Long   '移动窗体
Private Const CS_DROPSHADOW = &H20000 '窗体阴影
Private Const GCL_STYLE = (-26) '窗体阴影
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long '窗体阴影
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '窗体阴影
Dim WordCheck As Integer '遍历字词对照
Dim WithEvents Voice As SpVoice  '语音朗读
Attribute Voice.VB_VarHelpID = -1
Dim DatPath As String, FDat() As Byte, AllDat() As String '数据库内容
Dim user As String, YMNum As Integer, YMN3 As String, MN As Integer '用户昵称,消息
Dim AboutRob() As String, Hello() As String, AnsTyp() As String '关于机器人,机器人欢迎语,回复方式
Dim YourMsg As String, RobMsg As String '消息
Dim BadMsg() As String, RndUnAns1() As String, RndUnAns2() As String '不良消息,拒答的回复数据
Dim RndQus() As String, RndAns() As String, AIAns() As String '随机回答的数据,智能回答的数据,随机提问的数据
Dim Seconds As Integer '自动回复，倒计时
Dim RndDK() As String '听不懂时回复的数据
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long '播放 滴滴滴 的引用
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '功能是，让这条语句下一条语停顿 dwMilliseconds 时间后再运行，播放滴滴滴的引用
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long    '打开浏览器的引用

Private Sub Form_Load()
'On Error GoTo ErrorAI: '错误处理
DatPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Data.ait"      '释放资源文件的路径
FDat = LoadResData(101, "CUSTOM")     '释放资源文件
AboutRob = Split("|小志|公的,帅哥|08年8月诞生|地球村,中国|zhijian01|我天生比较聪明,只要你教过我的东西,我一定会记住的!|5311011lmq 和 moqiaoduo", "|") '信息
AnsTyp = Split("| 已经回复消息| 向你发送消息|不明觉厉...（需要你来调教调教）|", "|") '消息类型

If Dir(DatPath, vbHidden Or vbReadOnly Or vbSystem) = "" Then '如果数据库不存在
  user = InputBox("您好!我是智能聊天机器人,我叫 " & AboutRob(1) & vbCr & vbCr & "初次见面,请告诉我您的名字:")
  If user = "" Then '如果用户没有告知名字
UN:
    user = InputBox("喂喂,先别着急,我还不知道你的名字呢:", "Hello", "无名人士")
    If user = "" Then GoTo UN '循环，直到得到用户名位置
  End If
  Open DatPath For Output As #1 '将名字记入数据库
  Print #1, "|" & user & "|"
  Print #1, StrConv(FDat, vbUnicode)
  Close #1
End If
Initialize  '数据初始化
Me.Height = 7320 '窗体高度定位
Me.Width = 9025 '窗体宽度定位
Me.Top = (Screen.Height - Me.Height) / 2 '窗体位置定位
    Dim Rec As RECT, hRgn As Long  '圆角矩形
    GetWindowRect Me.hwnd, Rec  '圆角矩形
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 4, 4) '这里的数字表示圆角弧度  400,400 圆角的宽和高  '圆角矩形
    SetWindowRgn Me.hwnd, hRgn, True  '圆角矩形
SetClassLong Me.hwnd, GCL_STYLE, GetClassLong(Me.hwnd, GCL_STYLE) Or CS_DROPSHADOW '窗体阴影
LabRName.Caption = AboutRob(1)
LabSex.Caption = "性别：" & AboutRob(2)
LabBirth.Caption = "性别：" & AboutRob(3)
LabFrom.Caption = "来自：" & AboutRob(4)
LabAuthor.Caption = "原作：" & AboutRob(5)
LabRemark.Caption = "  " & AboutRob(6)
Exit Sub
ErrorAI:  '出现错误
If MsgBox("聊天数据系统出错,点击“确定”还原聊天数据系统!", vbExclamation + vbOKCancel, "Error on AI:") = vbOK Then
  Kill DatPath '删掉原有的数据库
  Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus '重新启动
  End
Else
  End
End If
End Sub

Sub Initialize() '初始化
Open DatPath For Binary As #1 '开始获取数据
ReDim FDat(FileLen(DatPath) - 1) As Byte
Get #1, , FDat
Close #1
AllDat = Split(StrConv(FDat, vbUnicode), "[AI_Dat]") '读取所有数据
user = Split(AllDat(0), "|")(1) '获取用户名
BadMsg = Split(AllDat(1), "|") '获取不良信息
RndUnAns1 = Split(AllDat(2), "|") '获取拒答的回复消息
RndQus = Split(AllDat(3), "|") '获取随机提问
RndAns = Split(AllDat(4), "|") '获取随机回答
RndDK = Split(AllDat(5), "|") '获取库内无答案时的回复消息
Hello = Split(AllDat(6), "|") '获取欢迎语
AIAns = Split(AllDat(7), "|") '获取智能回答
LabSmartWord.Caption = "数据：" & UBound(AIAns) & "个" '显示智能数据条数
Set Voice = New SpVoice    '语音阅读初始化
Set Voice.Voice = Voice.GetVoices("", "Language=804").Item(0) '语音阅读初始化
Me.Visible = True '窗体可见
Randomize '随机函数生成
RobSay Hello(Int(Rnd * (UBound(Hello) - 1)) + 1), 2 '欢迎~
End Sub
Private Sub CmdSend_Click() '发送信息
If Trim(TxtSend.Text) = "" Then '如果输入为空白
  Voice.Speak vbNullString, SVSFPurgeBeforeSpeak  '停止未完成的语音
  Voice.Speak "注意：请不要发送空白内容", SVSFlagsAsync   '语音播报
  Exit Sub
End If
YourMsg = LCase(Replace(TxtSend.Text, " ", "")) '提取内容
TxtTalk.Text = TxtTalk.Text & user & " " & Date & " " & Time & vbNewLine '显示发送文字时间
TxtTalk.Text = TxtTalk.Text & TxtSend.Text & vbNewLine & vbNewLine '显示发送的文字
TxtTalk.SelStart = Len(TxtTalk.Text)     '文本框显示至末尾
TxtSend.Text = "" '清空输入框
TxtSend.SetFocus '输入框获取焦点
YMNum = YMNum + 1      '消息数目累加
LabStatus.Caption = "您已发送 " & YMNum & " 条信息" '提示信息
TimerRevert.Enabled = True
LabRTip.Caption = "[回复中。。]" '更改状态
Seconds = 0
End Sub

Private Sub ListDate_Click() '单击调教列表框
  If ListDate.ListIndex = -1 Then Exit Sub '没有选中任何项，退出
  TxtTeach(1).Text = Split(ListDate.Text, "\")(0)
End Sub
Private Sub TxtSend_Change()  '检测到输入时关闭自动发言
Seconds = 0
End Sub
Private Sub TxtSend_KeyDown(KeyCode As Integer, Shift As Integer) '检测到输入时关闭自动发言
Seconds = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) '自动将焦点置于输入框
If PicTeach.Top <> 5400 Then TxtSend.SetFocus
End Sub

Private Sub TimerActive_Timer() '自动发言             '这段有点啰嗦，要改~不过我懒得改了。。
Seconds = Seconds + 1

If Seconds = 3 Then
Randomize
b = Int((2 - 1 + 1) * Rnd + 1)

If b = 1 Then
Randomize
RobSay RndQus(Int(Rnd * (UBound(RndQus) - 1)) + 1), 2
ElseIf b = 2 Then
Randomize
RobSay RndAns(Int(Rnd * (UBound(RndAns) - 1)) + 1), 2
End If
Seconds = 0
End If
End Sub

Private Sub RobSay(Msg As String, AT As Integer) '机器人回复，加个计时器是为了制造聊天效果 注：RobSay 后面的数字：1代表回复，2代表向你发消息，3 代表不明觉厉
'On Error Resume Next
Msg = Replace(Msg, "[rob]", AboutRob(1)) '信息替代
Msg = Replace(Msg, "[msg]", YourMsg)
Msg = Replace(Msg, "[day]", Date)
Msg = Replace(Msg, "[time]", Time)
RobMsg = Msg
TxtTalk.Text = TxtTalk.Text & AboutRob(1) & " " & Date & " " & Time & vbNewLine '显示回复文字时间
TxtTalk.Text = TxtTalk.Text & RobMsg & vbNewLine & vbNewLine '显示回复的文字
Voice.Speak vbNullString, SVSFPurgeBeforeSpeak  '停止未完成的语音
Voice.Speak RobMsg, SVSFlagsAsync   '语音播报
TxtTalk.SelStart = Len(TxtTalk.Text)  '文本框显示至末尾
TxtSend.SetFocus '输入框获取焦点
LabStatus.Caption = AboutRob(1) & AnsTyp(AT) '显示回复状态
If AT = 3 Then '如果需要调教
CmdTeach.Enabled = True '调教打开
Else
CmdTeach.Enabled = False
TimerActive.Enabled = True '自动回复计时器打开
End If
TimerRevert.Enabled = False '延时发消息计时器关闭
'PlayWavFile App.Path + "\msg.wav", 1, 0   '滴滴滴滴滴滴滴~
Seconds = 0 '自动回复计时清零
End Sub
Private Sub CmdEdit_Click(Index As Integer) '教机器人说话
Dim Qus As String, Ans As String, AIA() As String, HaveMsg As Boolean
   Dim WordSave As Integer '保存数据
Select Case Index
  Case 1
  If Trim(TxtTeach(1).Text) = "" Or Trim(TxtTeach(2).Text) = "" Or Qus = LCase(Replace(TxtTeach(1).Text, " ", "")) Or _
InStr(TxtTeach(1).Text, "\") <> 0 Or InStr(TxtTeach(1).Text, "|") <> 0 Then '判断是否为空,或重复教导,或含有分隔符
MsgBox "本次调教未能成功,可能的原因:" & vbCr & _
"1,关键字或智能回复为空;" & vbCr & "2,关键字含有“|”或“\”符号.", vbExclamation, "ERR"
Exit Sub
End If
Qus = LCase(Replace(TxtTeach(1).Text, " ", "")) & "\" '注：LCase是大写转小写
Ans = TxtTeach(2).Text & IIf(Right(TxtTeach(2).Text, 1) = "\", "", "\") '如果后面有\则不加\，如果没有，则加上
For i = 1 To UBound(BadMsg) - 1 '判断是否为不良消息
If InStr(Qus & Ans, BadMsg(i)) <> 0 Then
MsgBox "本次调教未能成功,可能的原因:" & vbCr & "调教内容含有不良字眼.", vbExclamation, "ERR"
TxtTeach(1).Text = ""
TxtTeach(2).Text = ""
Qus = ""
Ans = ""
Exit Sub
Exit For
End If
Next i

For WordCheck = 1 To UBound(AIAns) - 1 '判断是否已存在库中
AIA = Split(ListDate.List(WordCheck - 1), "\")
If LCase(Replace(TxtTeach(1).Text, " ", "")) <> AIA(0) Then
HaveMsg = False '不存在
Else
HaveMsg = True '存在
Exit For
End If
Next WordCheck

If HaveMsg = False Then '如果不存在
  ListDate.AddItem Qus & Ans & vbCrLf '添加新的内容

ElseIf HaveMsg = True Then  '如果已经存在
  ListDate.List(WordCheck - 1) = Qus & Ans & vbCrLf '替换新的内容
End If
MsgBox "调教成功!" & vbCr & "机器人会好好学习滴^_^", vbInformation + vbYesNo, "调教成功!"
TxtTeach(1).Text = ""
TxtTeach(2).Text = ""
   Case 2 '取消调教
   Open App.Path & "\Data.ait" For Output As #1 '保存数据
   AllDat = Split(StrConv(FDat, vbUnicode), "[AI_Dat]") '读取所有数据
   For WordSave = 0 To 6
   Print #1, AllDat(WordSave);
   Print #1, "[AI_Dat]";
   Next WordSave
   Print #1, "" '用于提出一行
   If ListDate.ListCount > 0 Then
   For WordSave = 0 To ListDate.ListCount - 1
   Print #1, "|" & Left(ListDate.List(WordSave), Len(ListDate.List(WordSave)) - 2)
   Next WordSave
   End If
   Close #1
   ListDate.Clear
   ListDate.Visible = False
   RobSay "    退出调教模式", 1
   Initialize '重新加载数据
   LabRTip.Caption = "[在线]" '更改状态
   TimerActive.Enabled = True '打开主动发言
   PicTeach.Top = 7440
   TxtTeach(1).Text = ""
   TxtTeach(2).Text = ""
   Case 3 '编辑数据库
   ListDate.Left = 0
   ListDate.Visible = True
End Select
End Sub
Private Sub TimerRevert_Timer() '正常消息智能应答  注：robotsay 后面的数字：1代表回复，2代表向你发消息，3代表不明觉厉
Dim AIA() As String
LabRTip.Caption = "[在线]"

Select Case YourMsg
  Case "/help" '帮助
    RobSay "[rob]告诉你怎么用：" + vbCrLf + "/help    显示帮助" + vbCrLf + "/teach   教[rob]说话" + vbCrLf + "/exit    跟[rob]Say Goodbye" _
    + vbCrLf + "/qqtalk  跟我的主人聊天←_←" + vbCrLf + "/getcode 索要我的源码0_0" + vbCrLf + "计算XX 进行四则混合运算" + vbCrLf + _
    "/clear 清屏", 1

  Case "/qqt" '坑爹
    RobSay "不经意间你又被我玩了，哇哈哈哈哈o(^▽^)o", 1

  Case "/clear" '清屏
    TxtTalk.Text = ""

  Case "/teach" '调教
    CmdTeach_Click
    TxtTeach(1).Text = ""
    TxtTeach(1).SetFocus
    TimerRevert.Enabled = False

  Case "/exit" '退出
    Unload Me

  Case "/qqtalk" '与作者聊天
    RobSay "程序猿都很忙啦，你直接更我聊天就行啦~ ←_←", 1
    'ShellExecute Me.hwnd, "open", "http://wpa.qq.com/msgrd?V=1&Uin=1163540807", vbNullString, vbNullString, SW_SHOWNORMAL

  Case "/getcode" '获取源码
    RobSay "这个，你自己到VB吧里去找吧←_←'", 1
    ShellExecute Me.hwnd, "open", "http://tieba.baidu.com/vb", vbNullString, vbNullString, SW_SHOWNORMAL '打开浏览器浏览VB吧

  Case Else
    If Left(Replace(YourMsg, " ", ""), 2) = "计算" Then '检测是否是计算模式
    YourMsg = Replace(YourMsg, "计算", "")
    YourMsg = Replace(YourMsg, "等于", "")
    YourMsg = Replace(YourMsg, "？", "")
    YourMsg = Replace(YourMsg, "=", "")
      If YourMsg <> "" Then '判断内容是否为空
        If JS(YourMsg) <> "" Then '判断计算是否正确
        Randomize
         RobSay Replace(Split("| 等于[JS]呀,拜托来点更高难度的| 这么简单你也不会呀,等于[JS]啊|我没算错的话,结果应该是[JS]|等于[JS],[rob]的计算绝对准确|", "|")(Int(Rnd * (4) + 1)), "[JS]", JS(YourMsg)), 1 '消息类型
        Else
        RobSay "太难了，把[rob]都算晕了还算不出来~换一个吧", 1
        End If
      Else
        RobSay "拜托，您得先告诉[rob]要算啥啊！", 1
      End If
    Exit Sub
    End If
    
    For WordCheck = 1 To UBound(BadMsg) - 1 '判断是否为不良消息
     YourMsg = Replace(YourMsg, AboutRob(1), "[rob]")
     If InStr(YourMsg, BadMsg(WordCheck)) <> 0 Then
      Randomize
      RobSay RndUnAns1(Int(Rnd * (UBound(RndUnAns1) - 1)) + 1), 1 '回复 请不要说脏话。。
      Exit Sub
     End If
    Next WordCheck

For WordCheck = 1 To UBound(AIAns) '智能应答
     AIA = Split(AIAns(WordCheck), "\")
     If InStr(YourMsg, AIA(0)) <> 0 Then
      Randomize
      RobSay AIA(Int(Rnd * (UBound(AIA) - 1)) + 1), 1
      Exit Sub
     ElseIf WordCheck = UBound(AIAns) Then '不在库里
      Randomize
      RobSay RndDK(Int(Rnd * (UBound(RndDK) - 1)) + 1), 3
     Exit Sub
     End If
   Next WordCheck
End Select
End Sub
Function JS(ByVal Expressions As String) As String              '智能计算（其实不智能。）
   Dim Mssc As Object
   Set Mssc = CreateObject("MSScriptControl.ScriptControl")
   Mssc.Language = "vbscript"
   On Error GoTo EvalErr
   JS = Mssc.Eval(Expressions)
   Exit Function
EvalErr:
JS = ""
End Function
Sub PlayWavFile(strFileName As String, PlayCount As Long, JianGe As Long)    '播放滴滴滴的消息音
    If Len(Dir(strFileName)) = 0 Then Exit Sub  'strFileName 要播放的文件名(带路径) 'playCount 播放的次数 'JianGe  多次播放时,每次的时间间隔
    If PlayCount = 0 Then Exit Sub
    If JianGe < 1000 Then JianGe = 1000
    DoEvents
    sndPlaySound strFileName, 16 + 1
    Sleep JianGe
    Call PlayWavFile(strFileName, PlayCount - 1, JianGe)
End Sub
Private Sub CmdTeach_Click()      '调教打开
PicTeach.Top = 5400
RobSay "    进入调教模式", 1
LabRTip.Caption = "[学习中。。]" '更改状态
      For WordCheck = 1 To UBound(AIAns)  '循环添加
      ListDate.AddItem AIAns(WordCheck)
      Next WordCheck
TxtTeach(1).Text = YourMsg
TxtTeach(2).SetFocus
TimerActive.Enabled = False
End Sub

Private Sub TxtTeach_Change(Index As Integer)
If Index = 1 Then
  For WordCheck = 0 To ListDate.ListCount - 1 '检测是否存在
     AIA = Split(ListDate.List(WordCheck), "\")
     If LCase(Replace(TxtTeach(1).Text, " ", "")) = AIA(0) Then
      TxtTeach(2).Text = Mid(ListDate.List(WordCheck), Len(TxtTeach(1).Text) + 2, Len(ListDate.List(WordCheck)) - Len(TxtTeach(1).Text) - 3)
      Exit Sub
      Else
      TxtTeach(2).Text = ""
      End If
   Next WordCheck
End If
End Sub
Private Sub PicTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '移动窗体
xx = X '移动窗体
yy = Y '移动窗体
End Sub '移动窗体
Private Sub PicTalk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '移动窗体
LabClose.BackStyle = 0
LabMin.BackStyle = 0
If Button = 1 Then Me.Move Me.Left + X - xx, Me.Top + Y - yy '移动窗体
End Sub '移动窗体
Private Sub LabClose_Click() '关闭
End
End Sub
Private Sub LabClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabClose.BackStyle = 1
LabMin.BackStyle = 0
End Sub
Private Sub LabMin_Click() '最小化
Me.WindowState = vbMinimized
LabMin.BackStyle = 0
End Sub
Private Sub LabMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabClose.BackStyle = 0
LabMin.BackStyle = 1
End Sub
