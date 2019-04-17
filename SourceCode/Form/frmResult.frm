VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmResult 
   Caption         =   "控制台"
   ClientHeight    =   9600
   ClientLeft      =   4350
   ClientTop       =   855
   ClientWidth     =   9300
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   9300
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5160
      Top             =   1440
   End
   Begin RichTextLib.RichTextBox txtResult 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   16748
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmResult.frx":06EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   540
   End
   Begin VB.Menu 控制 
      Caption         =   "程序控制(&C)"
      Begin VB.Menu 暂停 
         Caption         =   "暂停(&P)"
      End
      Begin VB.Menu 结束 
         Caption         =   "结束(&X)"
      End
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Append(ByVal s As String)
    Static a As Integer
    txtResult.Text = txtResult.Text & s
    If a = 0 Then
        txtResult.SelStart = 0
        txtResult.SelLength = Len(txtResult.Text)
        txtResult.SelColor = vbWhite
        a = 1
    End If
    txtResult.SelStart = Len(txtResult)
End Sub

Public Function InputWindow(ByVal Info As String)
    n = ""
    While n = ""
        n = InputBox(Info, "数据输入")
    Wend
    InputWindow = n
End Function

Public Function StopWindow()
    frmVars.Show 1
End Function

Public Sub Clear()
    txtResult.Text = ""
End Sub

 
Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Cancel = 1
    Unload frmVars
End Sub

Private Sub Timer1_Timer()
    If InStr(Me.Caption, "中止") > 0 Then
        暂停.Enabled = False
    Else
        暂停.Enabled = True
    End If
End Sub

Private Sub 结束_Click()
    Me.Hide
End Sub

Private Sub 暂停_Click()
    this_Graphic.Code.runFlag = False
End Sub
