VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plarn Coding Game"
   ClientHeight    =   16215
   ClientLeft      =   3885
   ClientTop       =   2475
   ClientWidth     =   15585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1081
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1039
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_Editor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   10560
      ScaleHeight     =   623
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Txt_letter 
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   6600
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmd_letter 
         Height          =   165
         Left            =   2520
         TabIndex        =   25
         Top             =   3120
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   3240
         Max             =   31
         TabIndex        =   24
         Top             =   3120
         Value           =   31
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3720
         TabIndex        =   23
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Pathway"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   20
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Effect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   19
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Object 
         Appearance      =   0  'Flat
         BackColor       =   &H0072D0F0&
         Height          =   240
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   240
      End
      Begin VB.CommandButton cmd_EditorCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H0072D0F0&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8760
         Width           =   1770
      End
      Begin VB.CommandButton cmd_EditorSave 
         Appearance      =   0  'Flat
         BackColor       =   &H0072D0F0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8760
         Width           =   1770
      End
      Begin VB.ListBox lst_Object 
         Appearance      =   0  'Flat
         BackColor       =   &H0072D0F0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         ItemData        =   "frmMain.frx":4492
         Left            =   120
         List            =   "frmMain.frx":4494
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   4080
         Width           =   3855
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Hero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Object"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Event"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chk_Editor 
         BackColor       =   &H0072D0F0&
         Caption         =   "Block"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_Debug 
      Appearance      =   0  'Flat
      BackColor       =   &H0072D0F0&
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10800
      MouseIcon       =   "frmMain.frx":4496
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10320
      Width           =   1770
   End
   Begin VB.TextBox txt_Order 
      Appearance      =   0  'Flat
      BackColor       =   &H00122331&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   360
      TabIndex        =   21
      Top             =   10320
      Visible         =   0   'False
      Width           =   9975
   End
   Begin VB.CommandButton cmd_Run 
      Appearance      =   0  'Flat
      BackColor       =   &H0072D0F0&
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12720
      MouseIcon       =   "frmMain.frx":57C0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10320
      Width           =   1890
   End
   Begin VB.CommandButton cmd_MoveDown 
      BackColor       =   &H0072D0F0&
      Height          =   930
      Left            =   8160
      MouseIcon       =   "frmMain.frx":6AEA
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7E14
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8880
      Width           =   930
   End
   Begin VB.CommandButton cmd_MoveUp 
      BackColor       =   &H0072D0F0&
      Height          =   930
      Left            =   8160
      MouseIcon       =   "frmMain.frx":8591
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":98BB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   930
   End
   Begin VB.CommandButton cmd_MoveRight 
      BackColor       =   &H0072D0F0&
      Height          =   930
      Left            =   9120
      MouseIcon       =   "frmMain.frx":9FFE
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":B328
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   930
   End
   Begin VB.CommandButton cmd_MoveLeft 
      BackColor       =   &H0072D0F0&
      Height          =   930
      Left            =   7200
      MouseIcon       =   "frmMain.frx":BA2B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":CD55
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   930
   End
   Begin VB.TextBox txtSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00122331&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10560
      Tag             =   "0"
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10080
      Top             =   360
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   15945
      Visible         =   0   'False
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   476
      SimpleText      =   "2"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22304
            Text            =   "状态"
            TextSave        =   "状态"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2019/4/17"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:37"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   12600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txt_Message 
      Appearance      =   0  'Flat
      BackColor       =   &H0072D0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   3135
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   1440
      Width           =   4215
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   9975
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
         Height          =   3810
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Txt_Input 
         Appearance      =   0  'Flat
         BackColor       =   &H00122331&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "frmMain.frx":D46E
         Top             =   4080
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.Image img_Font 
      Height          =   7680
      Left            =   4320
      Picture         =   "frmMain.frx":D474
      Top             =   14520
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.Image img_Event 
      Height          =   13440
      Left            =   720
      Picture         =   "frmMain.frx":15623
      Top             =   14040
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save as(&A)..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View(&V)"
      Begin VB.Menu mnuViewAIEditorBar 
         Caption         =   "AIEditor(&A)"
      End
      Begin VB.Menu mnuViewScriptEditorBar 
         Caption         =   "ScriptEditor(&S)"
      End
      Begin VB.Menu mnuViewStorytellingEditorBar 
         Caption         =   "StorytellingEditor(&T)"
      End
      Begin VB.Menu mnuViewMapEditorBar 
         Caption         =   "MapEditor(&M)"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "GameStatus(&B)"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run(&R)"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ffb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnubreak 
         Caption         =   "Break"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help(&H)"
      Begin VB.Menu mnuRegistration 
         Caption         =   "Registration(&R)"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Content(&C)"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About(&A) "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

'Functions of the main form

'====================================================================================================================================



Option Explicit


Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private cCode As New clsCode
Private m_EditorIndex As Byte
Private this_LetterID As Byte
Private this_StepsTimer As Byte

'=============================================================
'Describe:Running function of script
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Runit(this_Graphic As m_Graphic, Optional this_Order As Boolean)
    Dim ts() As String
    Dim i As Integer
    Dim s1, s2 As String
    With this_Graphic
         
        If this_Order = False Then
            
                .Code.Text = cCode.Init(.Code.Text)
        
                BanClear
                On Error Resume Next
                Call cCode.Run(.Code.Text)
        Else
                .Code.Text = cCode.Init(.Code.Text)
                .Code.Order = cCode.Init(.Code.Order)
        
                BanClear
                On Error Resume Next
                Call cCode.Run(.Code.Text & vbCrLf & .Code.Order)
        End If
    End With

    If Err.Number > 0 Then
        MsgBox Err.Description & "(The line " & ErrLine & " has some errors.)", vbCritical, "Error" & Err.Number
        Err.Clear
        ts = Split(txtSource.Text, vbCrLf)
        For i = 0 To UBound(ts)
            If i = ErrLine - 1 Then
                s2 = Len(ts(i))
                Exit For
            Else
                s1 = s1 + Len(ts(i)) + 2
            End If
        Next
        If s2 = "" Then s2 = 0
        txtSource.SelStart = s1
        txtSource.SelLength = s2
        txtSource.SetFocus
        SendMessage txtSource.hWnd, EM_SCROLLCARET, 0, 0
        
    End If
End Sub


'=============================================================
'Describe:Functions of the Editor's Checkbox
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub chk_Editor_Click(Index As Integer)
Select Case Index

    Case 0
        If this_Switch.Block = True Then
            chk_Editor(0).Value = 0
            this_Switch.Block = False
        Else
            chk_Editor(0).Value = 1
            this_Switch.Block = True
            
        End If
        
    Case 1
        If this_Switch.Event = True Then
            chk_Editor(1).Value = 0
            this_Switch.Event = False
        Else
            chk_Editor(1).Value = 1
            this_Switch.Event = True
        End If
        
    Case 2
        If this_Switch.Object = True Then
            chk_Editor(2).Value = 0
            this_Switch.Object = False
        Else
            chk_Editor(2).Value = 1
            this_Switch.Object = True
        End If
        
    Case 3
        If this_Switch.Player = True Then
            chk_Editor(3).Value = 0
            this_Switch.Player = False
        Else
            chk_Editor(3).Value = 1
            this_Switch.Player = True
        End If
        
     Case 4
        If this_Switch.Effect = True Then
            chk_Editor(4).Value = 0
            this_Switch.Effect = False
        Else
            chk_Editor(4).Value = 1
            this_Switch.Effect = True
        End If
        
     Case 5
        If this_Switch.Pathway = True Then
            chk_Editor(5).Value = 0
            this_Switch.Pathway = False
        Else
            chk_Editor(5).Value = 1
            this_Switch.Pathway = True
        End If
        
      Case 6
        If this_Switch.Letters = True Then
            chk_Editor(6).Value = 0
            this_Switch.Letters = False
        Else
            chk_Editor(6).Value = 1
            this_Switch.Letters = True
        End If
        
    Call Add_lst_Object(CByte(Index), this_Graphic)
    m_EditorIndex = CByte(Index)
End Select

End Sub


'=============================================================
'Describe:Function of Object Checking
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Add_lst_Object(m_index As Byte, this_Graphic As m_Graphic)
    Dim a, i, j, M As Integer
    M = 0
    
    
    Select Case m_index
    
        Case 0
           For j = 0 To 13
                For i = 0 To 15
                    If this_Graphic.Map.Block(i, j) = 1 Then
                        Cmd_Object(M).BackColor = &H0&
                    Else
                        Cmd_Object(M).BackColor = &H72D0F0
                    End If
                    M = M + 1
                Next i
           Next j
            
        Case 6
            For i = 0 To 223
                Cmd_Object(i).BackColor = &H72D0F0
            Next i

           
           For a = 42 To 51
            For j = 0 To 13
                For i = 0 To 15
                    If M < 32 Then
                        If this_Graphic.Map.Letter(a).MapPosition(M).x = i And this_Graphic.Map.Letter(51).MapPosition(M).y = j Then
                            Cmd_Object(M).BackColor = &H0&
                        Else
                            Cmd_Object(M).BackColor = &H72D0F0
                        End If
                        M = M + 1
                    End If
                Next i
           Next j
           Next a
           
            lst_Object.Clear
            For a = 42 To 51
                lst_Object.AddItem a
            Next a
            lst_Object.Selected(1) = True
            lst_Object.Selected(2) = True
             
    End Select
    
    m_EditorIndex = m_index

End Sub

'=============================================================
'Describe:Functions for Program Debug
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub cmd_Debug_Click()
'     this_Graphic.Code.runFlag = False
'    Call Runit(this_Graphic)
    
    
If mnuViewStatusBar.Checked = True Then
    this_Graphic.Code.runFlag = False
    Call Runit(this_Graphic, False)
Else
    this_Graphic.Code.Order = txtSource.Text
    
        If Head(txtSource.Text) = True Then
            this_Graphic.Code.runFlag = False
            Call Runit(this_Graphic, True)

        Else
            If this_Graphic.AILoaded = True Then
                Call Type_Check(txtSource.Text, MUD_ValROM)
            Else
                this_Graphic.Code.runFlag = False
                Call Runit(this_Graphic, True)
            End If
        End If
        
            txtSource.SelStart = 0
            txtSource.SelLength = Len(txtSource.Text)
            txtSource.SetFocus
    
End If

End Sub

Private Sub cmd_EditorCancel_Click()
    Dim i As Byte
    mnuViewMapEditorBar.Checked = Not mnuViewMapEditorBar.Checked
    Pic_Editor.Visible = mnuViewMapEditorBar.Checked
'    cmd_MoveUp.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveDown.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveLeft.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveRight.Visible = Not mnuViewMapEditorBar.Checked
    
    For i = 0 To 2
        chk_Editor(i).Value = 0
        this_Switch.Block = False
        this_Switch.Event = False
    Next i
End Sub

Private Sub cmd_EditorSave_Click()
    Call Files_Data_Map_Save(this_Graphic, this_FilePath, "Map")
End Sub

Private Sub cmd_letter_Click()
Dim a, b As Integer

    For b = 42 To 51
        For a = 0 To 31
            If this_Graphic.Map.Letter(b).MapPosition(a).y <> -1 And this_Graphic.Map.Letter(b).MapPosition(a).y <> -1 Then
                Txt_letter.Text = Txt_letter.Text & vbNewLine & "this_Graphic.Map.Letter(" & b & ").MapPosition(" & a & ").X = " & this_Graphic.Map.Letter(b).MapPosition(a).x & vbNewLine & "this_Graphic.Map.Letter(" & b & ").MapPosition(" & a & ").y = " & this_Graphic.Map.Letter(b).MapPosition(a).y
            End If
        Next a
    Next b
End Sub

'=============================================================
'Describe:Functions for item editing
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Cmd_Object_Click(Index As Integer)
    Dim i, j, a, b As Integer
    i = Index Mod 16
    j = Int(Index / 16)
    Debug.Print j & "-" & i
    Debug.Print Index
    Select Case m_EditorIndex
    
        Case 0
            If this_Graphic.Map.Block(i, j) = 0 Then
                this_Graphic.Map.Block(i, j) = 1
                Cmd_Object(Index).BackColor = &H0&
            Else
                this_Graphic.Map.Block(i, j) = 0
                Cmd_Object(Index).BackColor = &H72D0F0
            End If
        Case 6
            If this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).x = i And this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).y = j Then
                this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).x = -1
                this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).y = -1
                Cmd_Object(Index).BackColor = &H72D0F0
                If HScroll1.Value > 0 Then HScroll1.Value = HScroll1.Value - 1
            Else
                this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).x = i
                this_Graphic.Map.Letter(this_LetterID).MapPosition(HScroll1.Value).y = j
                Cmd_Object(Index).BackColor = &H0&
                If HScroll1.Value < 31 Then HScroll1.Value = HScroll1.Value + 1
                
            End If
    End Select
End Sub

'=============================================================
'Describe:Functions that run programs
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub cmd_Run_Click()

    
If mnuViewStatusBar.Checked = True Then
    this_Graphic.Code.runFlag = True
    Call Runit(this_Graphic, False)
Else
    this_Graphic.Code.Order = txtSource.Text
    
        If Head(txtSource.Text) = True Then
            this_Graphic.Code.runFlag = True
            Call Runit(this_Graphic, True)

        Else
            If this_Graphic.AILoaded = True Then
                Call Type_Check(txtSource.Text, MUD_ValROM)
            Else
                this_Graphic.Code.runFlag = True
                Call Runit(this_Graphic, True)
            End If
        End If
        
            txtSource.SelStart = 0
            txtSource.SelLength = Len(txtSource.Text)
            txtSource.SetFocus
    
End If
    
    
End Sub

'=============================================================
'Describe:Function to open a file after double-clicking the file list
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Filelist_DblClick()
    Call CodeFileLoad(this_Graphic)
    Call File_thisLoad
End Sub

Public Sub File_thisLoad()
    Dim ts() As String
    Dim i As Integer
    Dim s1, s2 As String
    
    txtSource.Text = "" 'this_Graphic.Code.Text
    
'        ts = Split(txtSource.Text, vbCrLf)
'        For i = 0 To UBound(ts)
'            If i = ErrLine - 1 Then
'                s2 = Len(ts(i))
'
'                Exit For
'            Else
'                s1 = s1 + Len(ts(i)) + 2
'            End If
'        Next
'        If s2 = "" Then s2 = 0
'
'        txtSource.SelStart = s1
'        txtSource.SelLength = s2
'        txtSource.SetFocus
'        SendMessage txtSource.hWnd, EM_SCROLLCARET, 0, 0
    
'    CurrentGameEvents = 63
'    Call Event_Get(this_Graphic, CurrentGameEvents)
    If this_Graphic.Code.Autorun = True Then
        Call mnuStart_Click
        this_Graphic.Code.Autorun = False
    End If
    flagChange = False
End Sub

        
'=============================================================
'Describe:Form Loading Function
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Form_Load()
    Pic_Editor.Picture = LoadPicture(this_FilePath.Graphics & "Editor.gif")
'     Left(this_FilePath.Code, Len(this_FilePath.Code) - 1)
    
    Dim i As Integer
    
    
    User_Codeing = False
    cmd_MoveUp.Visible = False
    cmd_MoveDown.Visible = False
    cmd_MoveRight.Visible = False
    cmd_MoveLeft.Visible = False
    
'    MySoftwarePassWord = Replace(GetSetting("Plarn", "Reg", "DiskNum", 1), "-", "")
    User_Code = Replace(GetSetting("Plarn", "Reg", "RegPSD", 1), "-", "")
    User_Codeing = True
'    If User_Code = Sm_02(Trim(MySoftwarePassWord)) Then
'        User_Codeing = True
'        'MsgBox ("Successful registration!")
'    Else
'        User_Codeing = False
'        MsgBox ("The software was not registered successfully")
'        frmRegistration.Show 1
''        Txt_Message.Visible = False
''        txtSource.Visible = False
''        cmd_Debug.Visible = False
''        cmd_Run.Visible = False
'    End If

End Sub

Private Sub Add_Buttons()
    Dim i As Integer
    Dim j As Integer
    Dim M As Integer
    
        M = 1
        For j = 0 To 13
            For i = 0 To 15
                Load Cmd_Object(M)
                Cmd_Object(M - 1).Visible = True
                Cmd_Object(M - 1).Left = 8 + i * 16
                Cmd_Object(M - 1).Top = 8 + j * 14
                M = M + 1
            Next i
        Next j
        
End Sub


'=============================================================
'Describe:Functions after Form Size Reset
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Form_Resize()
    On Error Resume Next

    PicMain.Width = this_Graphic.Screen.Width
    PicMain.Height = this_Graphic.Screen.Height
    Me.Width = this_Graphic.Screen.Width * 15 + 90
    Me.Height = this_Graphic.Screen.Height * 15 + 740
End Sub


'=============================================================
'Describe:Functions for Form Exit
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Form_Unload(Cancel As Integer)
     Dim i As Integer
    With this_Graphic.Buffer
        For i = 0 To 15
            SelectObject this_Graphic.Buffer.TileSetBmp(i), this_Graphic.Buffer.OldTilesetBmpDC(i)
            DeleteDC .TileSetBmp(i)
        Next i
        
        SelectObject .BackBuffer, .OldBackBufferDC
        DeleteDC .BackBuffer
        DeleteObject .BackBufferBmp
            
    End With

'    For i = 0 To Cmd_Object.UBound
'        Unload Cmd_Object(i)
'    Next i
    AI_Unload
    Unload Me
    End
End Sub


Private Sub lvCode_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        '加入断点列表
        BanAdd Int(Item.SubItems(1))
    Else
        '从列表删除
        BanDel Int(Item.SubItems(1))
    End If
End Sub

Private Sub lst_Object_Click()
    Select Case m_EditorIndex
    
        Case 6
            this_LetterID = lst_Object.ListIndex + 42
    
    End Select
End Sub

Private Sub mnubreak_Click()
    If EventBreak = True Then
        mnubreak.Checked = False
        EventBreak = False
    Else
        mnubreak.Checked = True
        EventBreak = True
    End If
End Sub

Private Sub mnuDebug_Click()
    this_Graphic.Code.runFlag = False
    Call Runit(this_Graphic)
End Sub

Public Sub mnuStart_Click()
    this_Graphic.Code.runFlag = True
    Call Runit(this_Graphic)
End Sub


Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    '如果这个工程没有帮助文件，显示消息给用户
    '可以在“工程属性”对话框中为应用程序设置帮助文件
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display help document.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

 

'=============================================================
'Describe:Game Editor Buttons
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub mnuViewMapEditorBar_Click()
    mnuViewMapEditorBar.Checked = Not mnuViewMapEditorBar.Checked
    Pic_Editor.Visible = mnuViewMapEditorBar.Checked
'    cmd_MoveUp.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveDown.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveLeft.Visible = Not mnuViewMapEditorBar.Checked
'    cmd_MoveRight.Visible = Not mnuViewMapEditorBar.Checked
    

    

    Dim i As Integer
    
    If mnuViewMapEditorBar.Checked = False Then
        For i = 0 To 3
            chk_Editor(i).Value = 0
            this_Switch.Block = False
            this_Switch.Event = False
        Next i
    Else

        
        If Cmd_Object.UBound = 0 Then
            Add_Buttons
        End If
        
            chk_Editor(0).Value = 1
            this_Switch.Block = True
        Call Add_lst_Object(0, this_Graphic)
    End If
End Sub



Private Sub mnuFileExit_Click()
    '
    mnuFileNew_Click
    End
End Sub



Private Sub mnuFileSaveAs_Click()
    If User_Codeing = True Then
        FileSaveAs
 
     Else

    End If
    
    
End Sub

Private Sub mnuFileSave_Click()
 
    
    If User_Codeing = True Then
        FileSave
 
     Else
 
    End If


End Sub

Private Sub mnuFileClose_Click()
    Call mnuFileNew_Click
End Sub

'=============================================================
'Describe:Functions to open files
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub mnuFileOpen_Click()

    
      If User_Codeing = True Then
            Dim sFile As String
            Call mnuFileNew_Click
            With dlgCommonDialog
                .DialogTitle = "Open"
                .CancelError = False
                'ToDo: 设置 common dialog 控件的标志和属性
                .Filter = "Plarn Files (*.pl,*.bas)|*.pl;*.bas|All Files(*.*)|*.*"
                .ShowOpen
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
                sFile = .FileName
            End With
            'ToDo: 添加处理打开的文件的代码
            this_FilePath.CodeName = sFile
            sFile = Replace(sFile, this_FilePath.Code, "")
            If sFile <> "" Then
                Call CodeFileLoad(this_Graphic, sFile)
                Call File_thisLoad
                flagChange = False
            End If
     Else

    End If
    
End Sub

'=============================================================
'Describe:Functions for new files
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub mnuFileNew_Click()
    If User_Codeing = True Then
        Dim R
        If flagChange = True Then
            R = MsgBox("This file has been changed, do you need to save it? ", vbQuestion + vbYesNoCancel, "Tips")
            If R = vbYes Then
                Call FileSave
            ElseIf R = vbNo Then
                '新建
                Call FileNew
            Else
                Exit Sub
            End If
        End If
        FileNew
     Else

    End If


End Sub


'=============================================================
'Describe:Buttons on the Game Status Bar
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub mnuViewStatusBar_Click()
    If User_Codeing = True Then
        If mnuViewStatusBar.Checked = True Then
            mnuViewStatusBar.Checked = False
            fMainForm.txtSource.Text = ""
            txt_Order.Visible = False
            cmd_MoveUp.Visible = False
            cmd_MoveDown.Visible = False
            cmd_MoveRight.Visible = False
            cmd_MoveLeft.Visible = False
        
        Else
            
            mnuViewStatusBar.Checked = True
            fMainForm.txtSource.Text = this_Graphic.Code.Text
            txt_Order.Visible = True
            cmd_MoveUp.Visible = True
            cmd_MoveDown.Visible = True
            cmd_MoveRight.Visible = True
            cmd_MoveLeft.Visible = True
        
        End If
     Else
         
    End If
    
End Sub

'=============================================================
'Describe:Test of Form after Key Click
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub PicMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
End Sub


'=============================================================
'Describe:Test of Textbox after Key Click
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Timer2_Timer()
    Static myln As Long
    If myln <> globalln And txtSource.SelLength = 0 Then
        myln = globalln
        Call cCode.CodeFormat(txtSource)
    End If
End Sub


'=============================================================
'Describe:Game Graphic Drawing Timer
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Timer3_Timer()
    
    Dim i, j, a As Byte
    Dim s As String
    Dim M As m_Position
    M.x = 2
    M.y = 0
    For i = 0 To 99
        Steps(i).Visible = False
    Next i
    
    
    For i = 0 To 7
        Randomize
        this_Graphic.EffectTimer(i) = this_Graphic.EffectTimer(i) + Rnd * 0.01
        
        If this_Graphic.EffectTimer(i) > 1000 Then this_Graphic.EffectTimer(i) = 0
    Next

    If CurrentSteps = 0 Then CurrentSteps = 1
    a = this_StepsTimer Mod CurrentSteps
    If Steps(a).Direction > -1 Then
        Steps(a).Visible = True
            If Steps(a).Direction > -1 And a > 0 Then
                Steps(a - 1).Visible = False
                If Steps(a - 1).Direction > -1 And a > 1 Then
                    Steps(a - 2).Visible = True
                    If Steps(a - 2).Direction > -1 And a > 2 Then
                        Steps(a - 3).Visible = False
                        If Steps(a - 3).Direction > -1 And a > 3 Then
                            Steps(a - 4).Visible = True
                            If Steps(a - 4).Direction > -1 And a > 4 Then
                                Steps(a - 5).Visible = False '\
                                If Steps(a - 5).Direction > -1 And a > 5 Then
                                    Steps(a - 6).Visible = True '\
                                    If Steps(a - 6).Direction > -1 And a > 6 Then
                                        Steps(a - 7).Visible = False '\
                                        If Steps(a - 7).Direction > -1 And a > 7 Then
                                            Steps(a - 8).Visible = True '\
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    End If
    this_StepsTimer = this_StepsTimer + 1
    If this_StepsTimer > 100 Then this_StepsTimer = 0
 
    DoEvents
    s = "Steps:" & CurrentEventTimer & "/" & this_Graphic.Player(CurrentPlayerID).Info.EventTimer
    'Call Draw_TilesBoardNote(this_Graphic, s, 0, 0, m, 3)
    
    M.x = 30
    M.y = 34
    Call Draw_TilesBoardNote(this_Graphic, s, 99, 40, M, 4)
    M.x = 570
    M.y = 34
    Call Draw_TilesBoardNote(this_Graphic, ShowLine, 98, 80, M, 5)
    
    
    s = "X:" & this_Graphic.Player(CurrentPlayerID).Info.C_Position.x & "  Y:" & this_Graphic.Player(CurrentPlayerID).Info.C_Position.y
    M.x = 800
    M.y = 46
    Call Draw_TilesBoardNote(this_Graphic, s, 97, 200, M, 5)
    
    
    DoEvents
    ShowLine
    
'    If User_Codeing = True Then
'        'Call GameDraw(PicMain.hDC, this_Graphic, this_Switch)
'     Else
'        frmRegistration.Show 1
'    End If
    Call GameDraw(PicMain.hDC, this_Graphic, this_Switch)
End Sub

Private Sub Txt_Message_DblClick()
    txtSource.Text = this_Graphic.Code.Text
End Sub

Private Sub txt_Order_KeyPress(KeyAscii As Integer)
    this_Graphic.Code.Order = txt_Order.Text
    
    If KeyAscii = 13 And Trim(this_Graphic.Code.Order) <> "" Then
        If Head(txt_Order.Text) = True Then
            this_Graphic.Code.runFlag = True
            Call Runit(this_Graphic, True)

        Else
            If this_Graphic.AILoaded = True Then
                Call Type_Check(txt_Order.Text, MUD_ValROM)
            Else
                this_Graphic.Code.runFlag = True
                Call Runit(this_Graphic, True)
            End If
        End If
        
            txt_Order.SelStart = 0
            txt_Order.SelLength = Len(txt_Order.Text)
            txt_Order.SetFocus
    End If
End Sub

Private Function Head(m_text As String) As Boolean
    Dim h As String
    Dim i
    If m_text = "" Then
        Head = False
        Exit Function
    End If
    h = Split(LCase(m_text), " ")(0)
    Dim keys() As String
    keys = Split(m_KeyWords, ",")
    For Each i In keys
        If Left(m_text & " ", Len(i)) = i Then
            If i = Trim(i) Then
                If Trim(m_text) <> Trim(i) Then
                    Head = False
                    Exit Function
                End If
            End If
            Head = True
            Exit Function
        End If
    Next
    Head = False
End Function

'=============================================================
'Describe:Keyboard function to enter keywords in the code edit box of the game
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub txtSource_KeyPress(KeyAscii As Integer)
    Dim rKeyAscii As Integer
    Dim s1, s2, flag, st As String
    Dim n As Integer
    Dim sts() As String
    Dim s As String
    
    rKeyAscii = -1
    '自动缩进/突出
    'KeyAscii = 0
    If KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Then
        With txtSource
            s1 = Left(.Text, .SelStart)
            s2 = Mid(.Text, .SelStart + 1, Len(.Text))
            If KeyAscii = 8 Then
                If Right(s1, 4) = "    " Then
                    s1 = Left(s1, Len(s1) - 4)
                    .Text = s1 & s2
                    .SelStart = Len(s1)
                Else
                    flag = 1
                End If
            ElseIf KeyAscii = 9 Then
                s1 = s1 & "    "
                .Text = s1 & s2
                .SelStart = Len(s1)
            Else
                If InStr(s1, vbCrLf) = 0 Then
                    st = s1
                Else
                    sts = Split(s1, vbCrLf)
                    st = sts(UBound(sts))
                End If
                If st <> "" Then
                    Do While Left(st, 1) = " "
                        st = Right(st, Len(st) - 1)
                        n = n + 1
                        If st = "" Then Exit Do
                    Loop
                End If
                Dim tl As New clsLine
                tl.strText = Trim(st)
                Select Case tl.Head
                    Case "if", "while", "do while", "select case", "for", "do", "case", "case else", "elseif", "else"
                        n = n + 4
                End Select
                If Right(LCase(tl.Body), 5) <> " then" And tl.Head = "if" Then n = n - 4
                If n < 0 Then n = 0
                s = s1 & vbCrLf & String(n, " ") & s2
                .Text = s
                .SelStart = Len(s1) + 2 + n
            End If
        End With
        rKeyAscii = 0
        If flag = 1 Then rKeyAscii = 8
    End If
    If rKeyAscii <> -1 Then KeyAscii = rKeyAscii
End Sub

Private Sub txtSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer2.Enabled = False
End Sub

Private Sub txtSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer2.Enabled = True
End Sub


Private Sub cmd_MoveDown_Click()
    Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
End Sub

Private Sub cmd_MoveLeft_Click()
    Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
End Sub

Private Sub cmd_MoveRight_Click()
    Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
End Sub

Private Sub cmd_MoveUp_Click()
    Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
End Sub


Private Sub txtSource_Click()
    ShowLine
End Sub


'=============================================================
'Describe:Keyboard function to enter keywords in the code edit box of the game
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub txtSource_Change()
    
    flagChange = True

If mnuViewStatusBar.Checked = True Then
    this_Graphic.Code.Text = txtSource.Text
Else
    this_Graphic.Code.Order = txtSource.Text
End If
    
    
 

 
End Sub

'=============================================================
'Describe:Number of rows in a string
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function ShowLine() As String
    Dim s1, ln As String
    s1 = Left(txtSource.Text, txtSource.SelStart)
    ln = Len(s1) - Len(Replace(s1, Chr(13), "")) + 1
    globalln = ln
    ShowLine = "Line:" & CStr(ln)
End Function

