VERSION 5.00
Begin VB.Form FrmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "±à¼­Êý¾Ý¿â"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8520
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox TxtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5880
      Width           =   8295
   End
   Begin VB.ListBox ListDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   5640
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Data.ait" For Input As #1
'While Not EOF(1)
'Line Input #1, a
'ListDate.AddItem a
'Wend
'Close #1
End Sub
