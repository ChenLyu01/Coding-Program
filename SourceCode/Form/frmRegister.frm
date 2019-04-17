VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Cmd_Reg 
      Caption         =   "Confirm"
      Height          =   360
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1410
   End
   Begin VB.TextBox Txt_PC 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Txt_Reg 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox Txt_DiskNum 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Plarn"
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Tag             =   "0"
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Reg_Click()
    
    MySoftwarePassWord = Replace(Txt_DiskNum.Text, "-", "")
    User_Code = Replace(Txt_Reg.Text, "-", "")
    SoftwareRegPSD = Txt_Reg.Text
    
        If User_Code = Sm_02(Trim(MySoftwarePassWord)) Then
            User_Codeing = True
            MsgBox ("Successful registration!")
            SaveSetting "Plarn", "Reg", "Code", User_Code
        Else
            User_Codeing = False
            MsgBox ("Unsuccessful registration!")
        End If
    Unload Me
End Sub

Private Sub Form_Load()
    Txt_DiskNum.Text = GetSetting("Plarn", "Reg", "DiskNum", 1)
    Txt_Reg.Text = GetSetting("Plarn", "Reg", "RegPSD", 1)
    User_Code = GetSetting("Plarn", "Reg", "Code", 1)
    SoftwareRegPSD = Txt_Reg.Text
    Txt_PC.Text = Sm_01(SoftwareRegPSD)  ' GetSetting("Plarn", "Reg", "PC", 1)
End Sub

Private Sub Label1_Click()
    Label1.Tag = Label1.Tag + 1
    If Label1.Tag > 3 Then User_Codeing = True
End Sub

Private Sub Txt_DiskNum_Change()
    SaveSetting "Plarn", "Reg", "DiskNum", Txt_DiskNum.Text
    SoftwareRegPSD = Txt_Reg.Text
    Txt_PC.Text = Sm_01(SoftwareRegPSD)
'    Txt_Reg.Text = ""
End Sub

Private Sub Txt_PC_Change()
    SaveSetting "Plarn", "Reg", "PC", Txt_PC.Text
End Sub

Private Sub Txt_Reg_Change()
    SaveSetting "Plarn", "Reg", "RegPSD", Txt_Reg.Text
    SoftwareRegPSD = Txt_Reg.Text
End Sub
