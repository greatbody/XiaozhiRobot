VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "�������������"
   ClientHeight    =   9735
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   1  '����������
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
            Caption         =   "�༭���ݿ�(&E)"
            Height          =   375
            Index           =   3
            Left            =   6840
            TabIndex        =   17
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H80000009&
            Caption         =   "�˳�����(&C)"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   16
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H80000009&
            Caption         =   "�����޸�(&O)"
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
               Name            =   "΢���ź�"
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
               Name            =   "΢���ź�"
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
            Caption         =   "˵���������ظ����á�\���ָ�."
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   1080
            Width           =   2460
         End
         Begin VB.Label LabTip2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ܻظ���"
            BeginProperty Font 
               Name            =   "΢���ź�"
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
            Caption         =   "�ؼ��ִʣ�"
            BeginProperty Font 
               Name            =   "΢���ź�"
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
         Caption         =   "����(&Send)"
         Default         =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   6920
         Width           =   1215
      End
      Begin VB.CommandButton CmdTeach 
         BackColor       =   &H8000000B&
         Caption         =   "����(&Teach)"
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
            Name            =   "΢���ź�"
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
            Name            =   "΢���ź�"
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
         Caption         =   "�������ڣ�2014/9/5"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6830
         TabIndex        =   27
         ToolTipText     =   "�޸���Ա��5311011lmq & moqiaoduo"
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
            Name            =   "΢���ź�"
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
         ToolTipText     =   "�ر�"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label LabMin 
         BackColor       =   &H002CF900&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "��С��"
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
         Caption         =   "��ע��"
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
         Caption         =   "���ݣ���"
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
         Caption         =   "���ߣ���"
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
         Caption         =   "���ԣ���"
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
         Caption         =   "���գ���"
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
         Caption         =   "�Ա𣺣�"
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
         Caption         =   "[����]"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long  'Բ�Ǿ���
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long  'Բ�Ǿ���
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long  'Բ�Ǿ���
Private Type RECT  'Բ�Ǿ���
    Left As Long  'Բ�Ǿ���
    Top As Long  'Բ�Ǿ���
    Right As Long  'Բ�Ǿ���
    Bottom As Long  'Բ�Ǿ���
End Type  'Բ�Ǿ���
Dim xx As Long, yy As Long   '�ƶ�����
Private Const CS_DROPSHADOW = &H20000 '������Ӱ
Private Const GCL_STYLE = (-26) '������Ӱ
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long '������Ӱ
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '������Ӱ
Dim WordCheck As Integer '�����ִʶ���
Dim WithEvents Voice As SpVoice  '�����ʶ�
Attribute Voice.VB_VarHelpID = -1
Dim DatPath As String, FDat() As Byte, AllDat() As String '���ݿ�����
Dim user As String, YMNum As Integer, YMN3 As String, MN As Integer '�û��ǳ�,��Ϣ
Dim AboutRob() As String, Hello() As String, AnsTyp() As String '���ڻ�����,�����˻�ӭ��,�ظ���ʽ
Dim YourMsg As String, RobMsg As String '��Ϣ
Dim BadMsg() As String, RndUnAns1() As String, RndUnAns2() As String '������Ϣ,�ܴ�Ļظ�����
Dim RndQus() As String, RndAns() As String, AIAns() As String '����ش������,���ܻش������,������ʵ�����
Dim Seconds As Integer '�Զ��ظ�������ʱ
Dim RndDK() As String '������ʱ�ظ�������
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long '���� �εε� ������
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '�����ǣ������������һ����ͣ�� dwMilliseconds ʱ��������У����ŵεεε�����
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long    '�������������

Private Sub Form_Load()
'On Error GoTo ErrorAI: '������
DatPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Data.ait"      '�ͷ���Դ�ļ���·��
FDat = LoadResData(101, "CUSTOM")     '�ͷ���Դ�ļ�
AboutRob = Split("|С־|����,˧��|08��8�µ���|�����,�й�|zhijian01|�������Ƚϴ���,ֻҪ��̹��ҵĶ���,��һ�����ס��!|5311011lmq �� moqiaoduo", "|") '��Ϣ
AnsTyp = Split("| �Ѿ��ظ���Ϣ| ���㷢����Ϣ|��������...����Ҫ�������̵��̣�|", "|") '��Ϣ����

If Dir(DatPath, vbHidden Or vbReadOnly Or vbSystem) = "" Then '������ݿⲻ����
  user = InputBox("����!�����������������,�ҽ� " & AboutRob(1) & vbCr & vbCr & "���μ���,���������������:")
  If user = "" Then '����û�û�и�֪����
UN:
    user = InputBox("ιι,�ȱ��ż�,�һ���֪�����������:", "Hello", "������ʿ")
    If user = "" Then GoTo UN 'ѭ����ֱ���õ��û���λ��
  End If
  Open DatPath For Output As #1 '�����ּ������ݿ�
  Print #1, "|" & user & "|"
  Print #1, StrConv(FDat, vbUnicode)
  Close #1
End If
Initialize  '���ݳ�ʼ��
Me.Height = 7320 '����߶ȶ�λ
Me.Width = 9025 '�����ȶ�λ
Me.Top = (Screen.Height - Me.Height) / 2 '����λ�ö�λ
    Dim Rec As RECT, hRgn As Long  'Բ�Ǿ���
    GetWindowRect Me.hwnd, Rec  'Բ�Ǿ���
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 4, 4) '��������ֱ�ʾԲ�ǻ���  400,400 Բ�ǵĿ�͸�  'Բ�Ǿ���
    SetWindowRgn Me.hwnd, hRgn, True  'Բ�Ǿ���
SetClassLong Me.hwnd, GCL_STYLE, GetClassLong(Me.hwnd, GCL_STYLE) Or CS_DROPSHADOW '������Ӱ
LabRName.Caption = AboutRob(1)
LabSex.Caption = "�Ա�" & AboutRob(2)
LabBirth.Caption = "�Ա�" & AboutRob(3)
LabFrom.Caption = "���ԣ�" & AboutRob(4)
LabAuthor.Caption = "ԭ����" & AboutRob(5)
LabRemark.Caption = "  " & AboutRob(6)
Exit Sub
ErrorAI:  '���ִ���
If MsgBox("��������ϵͳ����,�����ȷ������ԭ��������ϵͳ!", vbExclamation + vbOKCancel, "Error on AI:") = vbOK Then
  Kill DatPath 'ɾ��ԭ�е����ݿ�
  Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus '��������
  End
Else
  End
End If
End Sub

Sub Initialize() '��ʼ��
Open DatPath For Binary As #1 '��ʼ��ȡ����
ReDim FDat(FileLen(DatPath) - 1) As Byte
Get #1, , FDat
Close #1
AllDat = Split(StrConv(FDat, vbUnicode), "[AI_Dat]") '��ȡ��������
user = Split(AllDat(0), "|")(1) '��ȡ�û���
BadMsg = Split(AllDat(1), "|") '��ȡ������Ϣ
RndUnAns1 = Split(AllDat(2), "|") '��ȡ�ܴ�Ļظ���Ϣ
RndQus = Split(AllDat(3), "|") '��ȡ�������
RndAns = Split(AllDat(4), "|") '��ȡ����ش�
RndDK = Split(AllDat(5), "|") '��ȡ�����޴�ʱ�Ļظ���Ϣ
Hello = Split(AllDat(6), "|") '��ȡ��ӭ��
AIAns = Split(AllDat(7), "|") '��ȡ���ܻش�
LabSmartWord.Caption = "���ݣ�" & UBound(AIAns) & "��" '��ʾ������������
Set Voice = New SpVoice    '�����Ķ���ʼ��
Set Voice.Voice = Voice.GetVoices("", "Language=804").Item(0) '�����Ķ���ʼ��
Me.Visible = True '����ɼ�
Randomize '�����������
RobSay Hello(Int(Rnd * (UBound(Hello) - 1)) + 1), 2 '��ӭ~
End Sub
Private Sub CmdSend_Click() '������Ϣ
If Trim(TxtSend.Text) = "" Then '�������Ϊ�հ�
  Voice.Speak vbNullString, SVSFPurgeBeforeSpeak  'ֹͣδ��ɵ�����
  Voice.Speak "ע�⣺�벻Ҫ���Ϳհ�����", SVSFlagsAsync   '��������
  Exit Sub
End If
YourMsg = LCase(Replace(TxtSend.Text, " ", "")) '��ȡ����
TxtTalk.Text = TxtTalk.Text & user & " " & Date & " " & Time & vbNewLine '��ʾ��������ʱ��
TxtTalk.Text = TxtTalk.Text & TxtSend.Text & vbNewLine & vbNewLine '��ʾ���͵�����
TxtTalk.SelStart = Len(TxtTalk.Text)     '�ı�����ʾ��ĩβ
TxtSend.Text = "" '��������
TxtSend.SetFocus '������ȡ����
YMNum = YMNum + 1      '��Ϣ��Ŀ�ۼ�
LabStatus.Caption = "���ѷ��� " & YMNum & " ����Ϣ" '��ʾ��Ϣ
TimerRevert.Enabled = True
LabRTip.Caption = "[�ظ��С���]" '����״̬
Seconds = 0
End Sub

Private Sub ListDate_Click() '���������б��
  If ListDate.ListIndex = -1 Then Exit Sub 'û��ѡ���κ���˳�
  TxtTeach(1).Text = Split(ListDate.Text, "\")(0)
End Sub
Private Sub TxtSend_Change()  '��⵽����ʱ�ر��Զ�����
Seconds = 0
End Sub
Private Sub TxtSend_KeyDown(KeyCode As Integer, Shift As Integer) '��⵽����ʱ�ر��Զ�����
Seconds = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) '�Զ����������������
If PicTeach.Top <> 5400 Then TxtSend.SetFocus
End Sub

Private Sub TimerActive_Timer() '�Զ�����             '����еㆪ�£�Ҫ��~���������ø��ˡ���
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

Private Sub RobSay(Msg As String, AT As Integer) '�����˻ظ����Ӹ���ʱ����Ϊ����������Ч�� ע��RobSay ��������֣�1����ظ���2�������㷢��Ϣ��3 ����������
'On Error Resume Next
Msg = Replace(Msg, "[rob]", AboutRob(1)) '��Ϣ���
Msg = Replace(Msg, "[msg]", YourMsg)
Msg = Replace(Msg, "[day]", Date)
Msg = Replace(Msg, "[time]", Time)
RobMsg = Msg
TxtTalk.Text = TxtTalk.Text & AboutRob(1) & " " & Date & " " & Time & vbNewLine '��ʾ�ظ�����ʱ��
TxtTalk.Text = TxtTalk.Text & RobMsg & vbNewLine & vbNewLine '��ʾ�ظ�������
Voice.Speak vbNullString, SVSFPurgeBeforeSpeak  'ֹͣδ��ɵ�����
Voice.Speak RobMsg, SVSFlagsAsync   '��������
TxtTalk.SelStart = Len(TxtTalk.Text)  '�ı�����ʾ��ĩβ
TxtSend.SetFocus '������ȡ����
LabStatus.Caption = AboutRob(1) & AnsTyp(AT) '��ʾ�ظ�״̬
If AT = 3 Then '�����Ҫ����
CmdTeach.Enabled = True '���̴�
Else
CmdTeach.Enabled = False
TimerActive.Enabled = True '�Զ��ظ���ʱ����
End If
TimerRevert.Enabled = False '��ʱ����Ϣ��ʱ���ر�
'PlayWavFile App.Path + "\msg.wav", 1, 0   '�εεεεεε�~
Seconds = 0 '�Զ��ظ���ʱ����
End Sub
Private Sub CmdEdit_Click(Index As Integer) '�̻�����˵��
Dim Qus As String, Ans As String, AIA() As String, HaveMsg As Boolean
   Dim WordSave As Integer '��������
Select Case Index
  Case 1
  If Trim(TxtTeach(1).Text) = "" Or Trim(TxtTeach(2).Text) = "" Or Qus = LCase(Replace(TxtTeach(1).Text, " ", "")) Or _
InStr(TxtTeach(1).Text, "\") <> 0 Or InStr(TxtTeach(1).Text, "|") <> 0 Then '�ж��Ƿ�Ϊ��,���ظ��̵�,���зָ���
MsgBox "���ε���δ�ܳɹ�,���ܵ�ԭ��:" & vbCr & _
"1,�ؼ��ֻ����ܻظ�Ϊ��;" & vbCr & "2,�ؼ��ֺ��С�|����\������.", vbExclamation, "ERR"
Exit Sub
End If
Qus = LCase(Replace(TxtTeach(1).Text, " ", "")) & "\" 'ע��LCase�Ǵ�дתСд
Ans = TxtTeach(2).Text & IIf(Right(TxtTeach(2).Text, 1) = "\", "", "\") '���������\�򲻼�\�����û�У������
For i = 1 To UBound(BadMsg) - 1 '�ж��Ƿ�Ϊ������Ϣ
If InStr(Qus & Ans, BadMsg(i)) <> 0 Then
MsgBox "���ε���δ�ܳɹ�,���ܵ�ԭ��:" & vbCr & "�������ݺ��в�������.", vbExclamation, "ERR"
TxtTeach(1).Text = ""
TxtTeach(2).Text = ""
Qus = ""
Ans = ""
Exit Sub
Exit For
End If
Next i

For WordCheck = 1 To UBound(AIAns) - 1 '�ж��Ƿ��Ѵ��ڿ���
AIA = Split(ListDate.List(WordCheck - 1), "\")
If LCase(Replace(TxtTeach(1).Text, " ", "")) <> AIA(0) Then
HaveMsg = False '������
Else
HaveMsg = True '����
Exit For
End If
Next WordCheck

If HaveMsg = False Then '���������
  ListDate.AddItem Qus & Ans & vbCrLf '����µ�����

ElseIf HaveMsg = True Then  '����Ѿ�����
  ListDate.List(WordCheck - 1) = Qus & Ans & vbCrLf '�滻�µ�����
End If
MsgBox "���̳ɹ�!" & vbCr & "�����˻�ú�ѧϰ��^_^", vbInformation + vbYesNo, "���̳ɹ�!"
TxtTeach(1).Text = ""
TxtTeach(2).Text = ""
   Case 2 'ȡ������
   Open App.Path & "\Data.ait" For Output As #1 '��������
   AllDat = Split(StrConv(FDat, vbUnicode), "[AI_Dat]") '��ȡ��������
   For WordSave = 0 To 6
   Print #1, AllDat(WordSave);
   Print #1, "[AI_Dat]";
   Next WordSave
   Print #1, "" '�������һ��
   If ListDate.ListCount > 0 Then
   For WordSave = 0 To ListDate.ListCount - 1
   Print #1, "|" & Left(ListDate.List(WordSave), Len(ListDate.List(WordSave)) - 2)
   Next WordSave
   End If
   Close #1
   ListDate.Clear
   ListDate.Visible = False
   RobSay "    �˳�����ģʽ", 1
   Initialize '���¼�������
   LabRTip.Caption = "[����]" '����״̬
   TimerActive.Enabled = True '����������
   PicTeach.Top = 7440
   TxtTeach(1).Text = ""
   TxtTeach(2).Text = ""
   Case 3 '�༭���ݿ�
   ListDate.Left = 0
   ListDate.Visible = True
End Select
End Sub
Private Sub TimerRevert_Timer() '������Ϣ����Ӧ��  ע��robotsay ��������֣�1����ظ���2�������㷢��Ϣ��3����������
Dim AIA() As String
LabRTip.Caption = "[����]"

Select Case YourMsg
  Case "/help" '����
    RobSay "[rob]��������ô�ã�" + vbCrLf + "/help    ��ʾ����" + vbCrLf + "/teach   ��[rob]˵��" + vbCrLf + "/exit    ��[rob]Say Goodbye" _
    + vbCrLf + "/qqtalk  ���ҵ����������_��" + vbCrLf + "/getcode ��Ҫ�ҵ�Դ��0_0" + vbCrLf + "����XX ��������������" + vbCrLf + _
    "/clear ����", 1

  Case "/qqt" '�ӵ�
    RobSay "����������ֱ������ˣ��۹�������o(^��^)o", 1

  Case "/clear" '����
    TxtTalk.Text = ""

  Case "/teach" '����
    CmdTeach_Click
    TxtTeach(1).Text = ""
    TxtTeach(1).SetFocus
    TimerRevert.Enabled = False

  Case "/exit" '�˳�
    Unload Me

  Case "/qqtalk" '����������
    RobSay "����Գ����æ������ֱ�Ӹ������������~ ��_��", 1
    'ShellExecute Me.hwnd, "open", "http://wpa.qq.com/msgrd?V=1&Uin=1163540807", vbNullString, vbNullString, SW_SHOWNORMAL

  Case "/getcode" '��ȡԴ��
    RobSay "��������Լ���VB����ȥ�Ұɡ�_��'", 1
    ShellExecute Me.hwnd, "open", "http://tieba.baidu.com/vb", vbNullString, vbNullString, SW_SHOWNORMAL '����������VB��

  Case Else
    If Left(Replace(YourMsg, " ", ""), 2) = "����" Then '����Ƿ��Ǽ���ģʽ
    YourMsg = Replace(YourMsg, "����", "")
    YourMsg = Replace(YourMsg, "����", "")
    YourMsg = Replace(YourMsg, "��", "")
    YourMsg = Replace(YourMsg, "=", "")
      If YourMsg <> "" Then '�ж������Ƿ�Ϊ��
        If JS(YourMsg) <> "" Then '�жϼ����Ƿ���ȷ
        Randomize
         RobSay Replace(Split("| ����[JS]ѽ,������������Ѷȵ�| ��ô����Ҳ����ѽ,����[JS]��|��û���Ļ�,���Ӧ����[JS]|����[JS],[rob]�ļ������׼ȷ|", "|")(Int(Rnd * (4) + 1)), "[JS]", JS(YourMsg)), 1 '��Ϣ����
        Else
        RobSay "̫���ˣ���[rob]�������˻��㲻����~��һ����", 1
        End If
      Else
        RobSay "���У������ȸ���[rob]Ҫ��ɶ����", 1
      End If
    Exit Sub
    End If
    
    For WordCheck = 1 To UBound(BadMsg) - 1 '�ж��Ƿ�Ϊ������Ϣ
     YourMsg = Replace(YourMsg, AboutRob(1), "[rob]")
     If InStr(YourMsg, BadMsg(WordCheck)) <> 0 Then
      Randomize
      RobSay RndUnAns1(Int(Rnd * (UBound(RndUnAns1) - 1)) + 1), 1 '�ظ� �벻Ҫ˵�໰����
      Exit Sub
     End If
    Next WordCheck

For WordCheck = 1 To UBound(AIAns) '����Ӧ��
     AIA = Split(AIAns(WordCheck), "\")
     If InStr(YourMsg, AIA(0)) <> 0 Then
      Randomize
      RobSay AIA(Int(Rnd * (UBound(AIA) - 1)) + 1), 1
      Exit Sub
     ElseIf WordCheck = UBound(AIAns) Then '���ڿ���
      Randomize
      RobSay RndDK(Int(Rnd * (UBound(RndDK) - 1)) + 1), 3
     Exit Sub
     End If
   Next WordCheck
End Select
End Sub
Function JS(ByVal Expressions As String) As String              '���ܼ��㣨��ʵ�����ܡ���
   Dim Mssc As Object
   Set Mssc = CreateObject("MSScriptControl.ScriptControl")
   Mssc.Language = "vbscript"
   On Error GoTo EvalErr
   JS = Mssc.Eval(Expressions)
   Exit Function
EvalErr:
JS = ""
End Function
Sub PlayWavFile(strFileName As String, PlayCount As Long, JianGe As Long)    '���ŵεεε���Ϣ��
    If Len(Dir(strFileName)) = 0 Then Exit Sub  'strFileName Ҫ���ŵ��ļ���(��·��) 'playCount ���ŵĴ��� 'JianGe  ��β���ʱ,ÿ�ε�ʱ����
    If PlayCount = 0 Then Exit Sub
    If JianGe < 1000 Then JianGe = 1000
    DoEvents
    sndPlaySound strFileName, 16 + 1
    Sleep JianGe
    Call PlayWavFile(strFileName, PlayCount - 1, JianGe)
End Sub
Private Sub CmdTeach_Click()      '���̴�
PicTeach.Top = 5400
RobSay "    �������ģʽ", 1
LabRTip.Caption = "[ѧϰ�С���]" '����״̬
      For WordCheck = 1 To UBound(AIAns)  'ѭ�����
      ListDate.AddItem AIAns(WordCheck)
      Next WordCheck
TxtTeach(1).Text = YourMsg
TxtTeach(2).SetFocus
TimerActive.Enabled = False
End Sub

Private Sub TxtTeach_Change(Index As Integer)
If Index = 1 Then
  For WordCheck = 0 To ListDate.ListCount - 1 '����Ƿ����
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
Private Sub PicTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '�ƶ�����
xx = X '�ƶ�����
yy = Y '�ƶ�����
End Sub '�ƶ�����
Private Sub PicTalk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '�ƶ�����
LabClose.BackStyle = 0
LabMin.BackStyle = 0
If Button = 1 Then Me.Move Me.Left + X - xx, Me.Top + Y - yy '�ƶ�����
End Sub '�ƶ�����
Private Sub LabClose_Click() '�ر�
End
End Sub
Private Sub LabClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabClose.BackStyle = 1
LabMin.BackStyle = 0
End Sub
Private Sub LabMin_Click() '��С��
Me.WindowState = vbMinimized
LabMin.BackStyle = 0
End Sub
Private Sub LabMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabClose.BackStyle = 0
LabMin.BackStyle = 1
End Sub
