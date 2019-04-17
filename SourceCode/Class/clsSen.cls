VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and Cuiwei(Based on BasicPlus)
'E-mail:  plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Module declaration in the code section

'====================================================================================================================================


Option Explicit
Private strText1 As String
Private pos As Integer

'=============================================================
'Describe: Functions that handle sentences, read each character, and capitalize the handwritten letters of keywords
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function ReadWord()
    Dim ch1, ch As String
    Dim pp As String
    Dim i As Integer
    Dim ttp As String
    
   
'ReadWord = ""
'    pos = 1
'ch = ""
    
    On Error Resume Next
    If Me.Eof Then ReadWord = ""
    Do While True
        ch1 = Mid(strText1, pos, 1)
        If ch1 = """" Then
            pp = ""
            For i = pos + 1 To Len(strText1)
                If Mid(strText1, i, 1) = """" Then
                    pos = i + 1
                    ReadWord = """" & pp & """"
                    Exit Function
                Else
                    pp = pp & Mid(strText1, i, 1)
                End If
            Next
            Err.Raise 208, , "Lack of double quotation marks"
            Exit Function
        ElseIf ch1 = " " And Me.Eof = False Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    i = 0
    For i = pos To Len(strText1)
        ttp = Mid(strText1, i, 1)
        If GiveType(ttp) <> GiveType(ch1) And (GiveType(ttp) <> 1 Or GiveType(ch1) <> 2) Then
            Exit For
        Else
            pos = i
            ch = ch & ttp
        End If
    Next
    If Not Me.Eof Then pos = pos + 1
    ReadWord = ch
End Function

Public Function LookNextWord()
    Dim tpos As String
    Dim tword As String
    tpos = pos
    tword = ReadWord
    pos = tpos
    LookNextWord = tword
End Function

Public Property Get Eof() As Variant
    If pos = Len(strText1) Then
        Eof = True
    Else
        Eof = False
    End If
End Property

Public Property Let Eof(ByVal vNewValue As Variant)

End Property

Private Sub Class_Initialize()
    pos = 1
End Sub

Private Function GiveType(ByVal s As String)
    If IsNumeric(s) Or s = "." Then
        GiveType = 1
    ElseIf s >= "a" And s <= "z" Then
        GiveType = 2
    ElseIf s = " " Then
        GiveType = 3
    Else
        GiveType = Asc(s)
    End If
End Function

Public Property Get strText() As Variant
    strText = Left(strText1, Len(strText1) - 1)
End Property

Public Property Let strText(ByVal vNewValue As Variant)
    strText1 = vNewValue & " "
End Property
