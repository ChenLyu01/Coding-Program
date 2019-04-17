VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVars 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   9885
   ClientLeft      =   4335
   ClientTop       =   4530
   ClientWidth     =   4335
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmVars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVars.frx":06EA
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStep 
      BackColor       =   &H0072D0F0&
      Caption         =   "Step(Enter)"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9360
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvCode 
      Height          =   7455
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13150
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   7524592
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Break"
         Object.Width           =   661
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Line"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H0072D0F0&
      Cancel          =   -1  'True
      Caption         =   "Run(ESC)"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1815
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483642
      BackColor       =   7524592
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Variable"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   794
      EndProperty
   End
End
Attribute VB_Name = "frmVars"
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

'Functions of Forms for Debugging Commands

'====================================================================================================================================



'=============================================================
'Describe:Step-by-step debugging button
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub cmdContinue_Click()
    this_Graphic.Code.runFlag = True
    Unload Me
End Sub

'=============================================================
'Describe:Step-by-step debugging button
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub cmdStep_Click()
    this_Graphic.Code.runFlag = False
    Unload Me
End Sub

Private Sub Form_Load()
    If frmVars_top = 0 And frmVars_left = 0 Then Exit Sub
    Me.Top = frmVars_top
    Me.Left = frmVars_left
End Sub

Private Sub Form_Resize()
    Call Reflush(this_Graphic)
End Sub

'=============================================================
'Describe:Read all the data into the control of the form
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function Reflush(this_Graphic As m_Graphic, Optional Order_index As Boolean)
    Dim a As ListItem
    '读入变量
    For i = 1 To globalCV.VarNames.Count
        Set a = lv.ListItems.Add(, , globalCV.VarNames(i))
        a.SubItems(1) = globalCV.getVar(globalCV.VarNames(i))
        a.SubItems(2) = "Double"
    Next
    '读入代码
    
    iss = Split(this_Graphic.Code.Text, vbCrLf)
    For i = 0 To UBound(iss)
        'lvCode.ListItems.Add , , i
        Set a = lvCode.ListItems.Add(, , "")
        a.SubItems(1) = i + 1
        a.SubItems(2) = iss(i)
        a.Checked = BanHave(i + 1)
        If ErrLine = i + 1 Then
            a.Bold = True
            a.ListSubItems(2).ForeColor = vbBlue
            a.Selected = True
        End If
    Next
    If ErrLine <> 0 Then
        '控制滚动条，让被执行的那行总可见
        lvCode.ListItems(ErrLine).EnsureVisible
        lvCode.SetFocus
    End If
End Function


'=============================================================
'Describe:Form Unloading Function
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Sub Form_Unload(Cancel As Integer)
    frmVars_top = Me.Top
    frmVars_left = Me.Left
    
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

