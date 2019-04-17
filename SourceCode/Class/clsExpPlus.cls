VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpPlus"
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

Private Oper() As String
Private sf As New clsVar
Private skop As New clsStack
Private skexp As New clsStack

'=============================================================
'Describe: Initialization class module
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Sub Class_Initialize()
    Dim i, j As Integer
    ReDim Oper(6)
    Oper(1) = "#,szqzr,azbzs,aznzd,ozr,xzozr,nzozt"
    Oper(2) = "<,>,=,<->,<-=,>-="
    Oper(3) = "+,-,mzozd"
    Oper(4) = "*,/"
    Oper(5) = "^"
    Oper(6) = "(,)"
    For i = 1 To 6
        For j = 0 To UBound(Split(Oper(i), ","))
            sf.setVar Split(Oper(i), ",")(j), i
        Next
    Next
End Sub


Public Function logic(ByVal s As String)
    logic = CBool(logic_it(s))
End Function

'=============================================================
'Describe: Processing strings
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function Exp(ByVal s As String)
    Dim a As String
    Dim e1 As String
    Dim e2 As String
    Dim i As Integer
    Dim ch As String
    Dim le2  As String
    On Error Resume Next
    
    a = InstrPlus(s, "=")
    If a = 0 Or (InstrPlus(s, "(") < InstrPlus(s, "=") And InstrPlus(s, "(") <> 0) Then
        Exp = logic_it(s)
        Exit Function
    End If
    e1 = Trim(Left(s, a - 1))
    e2 = Trim(Mid(s, a + 1, Len(s)))
    If e1 = "" Then
        Err.Raise 102, , "The variable name cannot be empty"
        Exit Function
    End If
    If IsNumeric(Left(e1, 1)) Then
        Err.Raise 103, , "Variable names cannot begin with numbers"
        Exit Function
    End If
    On Error Resume Next
    If Len(e1) > 1 Then
        For i = 2 To Len(e1)
            ch = Mid(e1, i, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or IsNumeric(ch)) Then
                Err.Raise 104, , "Illegal variable name"
                Exit Function
            End If
        Next
    End If
    le2 = logic_it(e2)
    If Err.Number > 0 Then Exit Function
    globalCV.setVar e1, le2
    Exp = globalCV.getVar(e1)
End Function

'=============================================================
'Describe: Processing strings
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Function logic_it(ByVal s As String)
    '初始化s
    'Debug.Print "logic_it:" & s
    Dim tmp As String
    
    s = s & "+0"
    If Left(s, 1) = "-" Then s = "0" & s
    s = ReplacePlus(s, "(-", "(0-")
    s = ReplacePlus(s, "( -", "(0-")
    s = ReplacePlus(s, "and", "aznzd")
    s = ReplacePlus(s, "xor", "xzozr")
    s = ReplacePlus(s, "or", "ozr")
    s = ReplacePlus(s, "not", "nzozt")
    s = ReplacePlus(s, "<>", "<z>")
    s = ReplacePlus(s, ">=", ">z=")
    s = ReplacePlus(s, "<=", "<z=")
    s = ReplacePlus(s, "mod", "mzozd")
    s = ReplacePlus(s, "abs", "azbzs")
    s = ReplacePlus(s, "sqr", "szqzr")
    Dim se As New clsSen
    skop.Clear
    skexp.Clear
    skop.Push "#"
    skexp.Push "#"
    se.strText = s
    While Not se.Eof
        tmp = se.ReadWord
        If isKey(tmp) Then
            If sf.getVar(tmp) > sf.getVar(skop.Item(0)) Or (tmp = ")" And skop.Item(0) = "(") Or (tmp = "nzozt" And skop.Item(0) = "nzozt") Then
                '外>内
                If tmp <> ")" Then
                    skop.Push tmp
                Else
                    While skop.Count > 1 And skop.Item(0) <> "("
                        skexp.Push skop.Pop
                        Calculate
                        If Err.Number > 0 Then Exit Function
                        '----------need to do something
                    Wend
                    skop.Pop
                    'skop.Push tmp
                End If
            Else
                '外<=内
                While skop.Count > 1 And skop.Item(0) <> "(" And sf.getVar(tmp) <= sf.getVar(skop.Item(0))
                    skexp.Push skop.Pop
                    Calculate
                    If Err.Number > 0 Then Exit Function
                    '----------need to do something
                Wend
                skop.Push tmp
            End If
        Else
            '数字/变量
            skexp.Push tmp
        End If
    Wend
    While skop.Count > 1
        skexp.Push skop.Pop
        Calculate
        If Err.Number > 0 Then Exit Function
        '---need to do something
    Wend
    On Error Resume Next
    If skexp.Count <> 2 Then
        Err.Raise 120, , "Function call missing something :!"
        Exit Function
    End If
    If isKey(skexp.Item(0)) Then
        Err.Raise 121, , "Expression syntax error"
        Exit Function
    End If
    logic_it = skexp.Pop
End Function

'=============================================================
'Describe: Compute the number in the script
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Function Calculate()
    On Error Resume Next
    Dim t As Double
    Dim opt As String
    Dim op1, op2
    
    While isKey(skexp.Item(0)) And skexp.Count >= 2
        opt = skexp.Pop
        If opt = "nzozt" Then
            skexp.Push IIf((Not CBool(myVar(skexp.Pop))), 1, 0)
        ElseIf opt = "szqzr" Then
            skexp.Push Sqr(myVar(skexp.Pop))
        ElseIf opt = "azbzs" Then
            skexp.Push Abs(myVar(skexp.Pop))
        Else
            op2 = myVar(skexp.Pop)
            op1 = myVar(skexp.Pop)
            Select Case opt
                Case "+"
                    t = op1 + op2
                Case "-"
                    t = op1 - op2
                Case "*"
                    t = op1 * op2
                Case "/"
                    t = op1 / op2
                Case "aznzd"
                    t = IIf((CBool(op1) And CBool(op2)), 1, 0)
                Case "ozr"
                    t = IIf((CBool(op1) Or CBool(op2)), 1, 0)
                Case "^"
                    t = op1 ^ op2
                Case "<"
                    t = IIf((op1 < op2), 1, 0)
                Case ">"
                    t = IIf((op1 > op2), 1, 0)
                Case "<z>"
                    t = IIf((op1 <> op2), 1, 0)
                Case ">z="
                    t = IIf((op1 >= op2), 1, 0)
                Case "<z="
                    t = IIf((op1 <= op2), 1, 0)
                Case "mzozd"
                    t = op1 Mod op2
                Case "="
                    t = IIf((op1 = op2), 1, 0)
                Case "xzozr"
                    t = IIf((CBool(op1) Xor CBool(op2)), 1, 0)
            End Select
            skexp.Push t
        End If
    Wend
End Function

'=============================================================
'Describe: Read variables
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Function myVar(ByVal Name As String) As Double
    If IsNumeric(Name) Then
        myVar = Name
    Else
        myVar = globalCV.getVar(Name)
    End If
End Function

Public Function isKey(ByVal s As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To 6
        For j = 0 To UBound(Split(Oper(i), ","))
            If Split(Oper(i), ",")(j) = Trim(s) Then
                isKey = True
                Exit Function
            End If
        Next
    Next
    isKey = False
End Function
