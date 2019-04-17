VERSION 5.00
Begin VB.Form frmFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plarn"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic_Book 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmd_BookNew 
         BackColor       =   &H0072D0F0&
         Caption         =   "New"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   1290
      End
      Begin VB.FileListBox Filelist 
         Appearance      =   0  'Flat
         BackColor       =   &H0072D0F0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3270
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmd_BookOkay 
         BackColor       =   &H0072D0F0&
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4680
         Width           =   1290
      End
      Begin VB.CommandButton cmd_BookCancel 
         BackColor       =   &H0072D0F0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4680
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmd_BookCancel_Click()
    Unload frmFile
End Sub

Private Sub cmd_BookOkay_Click()
    Call CodeFileLoad
    Unload frmFile
End Sub


Private Sub Filelist_DblClick()
    Call CodeFileLoad
    Unload frmFile
End Sub

Private Sub Form_Load()
    
    pic_Book.Picture = LoadPicture(this_FilePath.Graphics & "Windows.gif")
    pic_Book.Width = 386
    pic_Book.Height = 363
    pic_Book.Left = 0
    pic_Book.Top = 0
    
    frmFile.Width = pic_Book.Width * 15 + 60
    frmFile.Height = pic_Book.Height * 15 + 440
    
    Filelist.Path = this_FilePath.Code
End Sub

Private Sub CodeFileLoad()
    Dim s As String
    s = FileRead(this_FilePath.Code & Filelist.FileName)
    fMainForm.txtSource.Text = s
    this_FilePath.CodeName = Filelist.FileName
    flagChange = False
End Sub

