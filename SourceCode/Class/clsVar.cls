VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVar"
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



Option Explicit

Private Vars As New Collection
Public VarNames As New Collection

Public Sub Clear()
     '
    Dim V As New Collection
    Set Vars = V
    Dim vn As New Collection
    Set VarNames = New Collection
End Sub

'=============================================================
'Describe: Variable settings
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function setVar(ByVal Name As String, ByVal Value As Double)
    '
    On Error Resume Next
    Dim V() As String
    Dim t As String
    V = Split(m_Keyword, ",")
    If inArray(V, Name) Then
        Err.Raise 59, , "Variable names " & Name & "is illegal"
        Exit Function
    End If
    t = Vars(Name)
    AddVarName Name
    If Err.Number > 0 Then
        Err.Clear
        Vars.Add Value, Name
    Else
        Vars.Remove (Name)
        Vars.Add Value, Name
    End If
End Function

Public Function getVar(ByVal Name As String)
    Dim t As String
    '
    On Error Resume Next
    Dim V() As String
    V = Split(m_Keyword, ",")
    If inArray(V, Name) Then
        Err.Raise 59, , "Variable names " & Name & "is illegal"
        Exit Function
    End If
    t = Vars(Name)
    AddVarName Name
    If Err.Number > 0 Then
        Vars.Add 0, Name
        Err.Clear
        getVar = 0
    Else
        getVar = t
    End If
End Function

Private Function AddVarName(ByVal Name As String)
    Dim i As Integer
    For i = 1 To VarNames.Count
        If VarNames(i) = Name Then
            Exit Function
        End If
    Next
    VarNames.Add Name
End Function

