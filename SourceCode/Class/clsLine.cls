VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and Cuiwei
'E-mail:  plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Module declaration in the code section

'====================================================================================================================================


Public strText As String

Public strType As String
' print,input and so on

'-------if ---------
Public if_else_id As Integer
Public if_endif_id As Integer
Public else_if_id As Integer

Public for_next_id As Integer
Public for_min As Integer
Public for_max As Integer
Public for_step As Integer
Public for_var_name As String
Public exitfor_for_id As Integer
Public next_for_id As Integer

Public dowhile_loop As String
Public loop_dowhile As String
Public do_loopuntil_id As String
Public loopuntile_do_id As String

Public exitdo_dowhile As String
Public exitdo_do As String

Public vLine As Integer     'Logical line number, goto
Public phLine As Integer    'Physical line number, indicating misuse


'=============================================================
'Describe: Processing functions for each line of code, extracting keywords
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Property Get Head() As Variant
    Dim h As String
    If strText = "" Then
        Head = ""
        Exit Property
    End If
    h = Split(LCase(strText), " ")(0)
    Dim keys() As String
    keys = Split(m_KeyWords, ",")
    For Each i In keys
        If Left(strText & " ", Len(i)) = i Then
            If i = Trim(i) Then
                If Trim(strText) <> Trim(i) Then
                    Head = ""
                    Exit Property
                End If
            End If
            Head = Trim(i)
            Exit Property
        End If
    Next
    Head = ""
End Property

Public Property Let Head(ByVal vNewValue As Variant)
    'do nothing
End Property

Public Property Get Body() As Variant
    Body = Trim(Right(strText, Len(strText) - Len(Head)))
End Property

Public Property Let Body(ByVal vNewValue As Variant)

End Property
