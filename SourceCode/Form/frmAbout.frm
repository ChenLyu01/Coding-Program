VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Plarn"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":4492
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   320
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   2640
      Width           =   1467
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      Caption         =   "Boston MA US."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009CCA00&
      Height          =   450
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Chen Lyu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      MousePointer    =   4  'Icon
      TabIndex        =   4
      Top             =   360
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   5672
      Y1              =   2020
      Y2              =   2020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "www.plarn.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFA935&
      Height          =   270
      Left            =   3210
      TabIndex        =   3
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Chen Lyu"
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
      Height          =   825
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
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

'The Function of About us

'====================================================================================================================================


Const KEY_ALL_ACCESS = &H2003F
                                          

'
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         '
Const REG_DWORD = 4                      '


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const SW_SHOW = 5
Private Sub Form_Load()
    lblVersion.Caption = "Version:" & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        '  .
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        '  ..
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                '
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                '  .
                Else
                        GoTo SysInfoErr
                End If
        '  .
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System information is unavailable at this time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           '
        Dim rc As Long                                          '
        Dim hKey As Long                                        '
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  '
        Dim tmpVal As String                                    '
        Dim KeyValSize As Long                                  '
        '------------------------------------------------------------
        '
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) '
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          '  ..
        

        tmpVal = String$(1024, 0)                             '
        KeyValSize = 1024                                       '
        

        '------------------------------------------------------------
        '
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    '
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          '
      

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)


        '------------------------------------------------------------
        '
        '------------------------------------------------------------
        Select Case KeyValType                                  '
        Case REG_SZ                                             '
                KeyVal = tmpVal                                     '
        Case REG_DWORD                                          '
                For i = Len(tmpVal) To 1 Step -1
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   '
                Next
                KeyVal = Format$("&h" + KeyVal)                     '
        End Select
        

        GetKeyValue = True                                      '
        rc = RegCloseKey(hKey)                                  '
        Exit Function                                           '
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             '
        GetKeyValue = False                                     '
        rc = RegCloseKey(hKey)                                  '
End Function

Private Sub lblDescription_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.plarn.com/", "", "", SW_SHOW)
End Sub

Private Sub lblTitle_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.plarn.com/", "", "", SW_SHOW)
End Sub

Private Sub lblVersion_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.plarn.com/", "", "", SW_SHOW)
End Sub
