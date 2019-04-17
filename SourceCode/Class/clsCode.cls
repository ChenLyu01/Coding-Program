VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and James L. Dean
'E-mail:  plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Module declaration in the code section

'====================================================================================================================================



Option Explicit
Private Lines() As New clsLine              '按行存放源代码的数组
Private iline As Integer                    '存入数组时用的临时变量
Private ce As New clsExpPlus                    '表达式处理模块
Private this_Code_Notes As String
Private this_Code_Space As String
Private this_Code_If As String
Private this_Code_Then As String
Private this_Code_Else As String
Private this_Code_EndIf As String
Private this_Code_ElseIf As String
Private this_Code_Input As String
Private this_Code_While As String
Private this_Code_DoWhile As String
Private this_Code_Wend As String
Private this_Code_Loop As String
Private this_Code_SelectCase As String
Private this_Code_EndSelect As String
Private this_Code_Case As String
Private this_Code_CaseElse As String
Private this_Code_To As String
Private this_Code_Or As String
Private this_Code_And As String
Private this_Code_For As String
Private this_Code_ExitFor As String
Private this_Code_Next As String
Private this_Code_Do As String
Private this_Code_LoopUntil As String
Private this_Code_ExitDo As String
Private this_Code_Print As String
Private this_Code_Step As String
Private this_Code_End As String
Private this_Code_Goto As String
Private this_Code_GoSub As String
Private this_Code_Return As String
Private this_Code_Stop As String
Private this_Game_Autorun As String
Private this_Game_MessageShow As String
Private this_Game_MessageClear As String
Private this_Game_HeroCreate As String
Private this_Game_HeroMoveLeft As String
Private this_Game_HeroMoveRight As String
Private this_Game_HeroMoveUp As String
Private this_Game_HeroMoveDown As String
Private this_Game_Tile As String
Private this_Game_Pathway As String
Private this_Game_PathwayDraw As String
Private this_Game_PathwayClear As String
Private this_Game_BlockShow As String
Private this_Game_BlockHide As String
Private this_Game_HeroShow As String
Private this_Game_HeroHide As String
Private this_Game_HeroSteps As String
Private this_Game_HeroMoveXY As String


Private this_Game_EventShow As String
Private this_Game_EventHide As String
Private this_Game_PathwayShow As String
Private this_Game_PathwayHide As String
Private this_Game_ObjectHide As String
Private this_Game_ObjectShow As String
Private this_Game_EffectHide As String
Private this_Game_EffectShow As String
Private this_Game_EventDestroy As String
Private this_Game_EventCreate As String
Private this_Game_EventClear As String
Private this_Game_ObjectCreate As String
Private this_Game_ObjectDestroy As String
Private this_Game_ObjectClear As String

Private this_Game_EffectCreate As String
Private this_Game_EffectDestroy As String
Private this_Game_EffectClear As String
Private this_Game_EffectEnable As String
Private this_Game_EffectDisable As String

Private this_Game_BlockDestroy As String
Private this_Game_BlockCreate As String
Private this_Game_BlockClear As String

Private this_Game_FileOpen As String
Private this_Game_FileAI As String

Private DefauleEffectTimer As Integer
Private this_Game_PathwayDestroy As String
Private this_Game_PathwayWithBlock As String
Private this_Game_HeroTalk As String
Private this_Game_HeroOpen As String
Private this_Game_HeroCollect As String
Private this_Game_GameTask As String
Private this_Game_MessageClose As String
Private this_Game_FaceUp As String
Private this_Game_FaceDown As String
Private this_Game_FaceLeft As String
Private this_Game_FaceRight As String


'=============================================================
'Describe: Key words for processing functions
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function Init(ByVal this_Code As String) As String
    Init = LcasePlus(this_Code)          '
    DefauleEffectTimer = 140

    this_Code_Notes = "#"
    this_Code_Space = " "
    this_Code_If = "if"
    this_Code_Then = "then"
    this_Code_Else = "else"
    this_Code_ElseIf = "elseif"
    this_Code_EndIf = "end if"
    this_Code_Input = "input"
    this_Code_While = "while"
    this_Code_DoWhile = "do while"
    this_Code_Wend = "wend"
    this_Code_Loop = "loop"
    this_Code_SelectCase = "select case"
    this_Code_Case = "case"
    this_Code_To = "to"
    this_Code_Or = "or"
    this_Code_CaseElse = "case else"
    this_Code_EndSelect = "end select"
    this_Code_And = "and"
    this_Code_For = "for"
    this_Code_ExitFor = "exit for"
    this_Code_Next = "next"
    this_Code_Do = "do"
    this_Code_LoopUntil = "loop until"
    this_Code_ExitDo = "exit do"
    this_Code_Print = "print"
    this_Code_Step = "step"
    this_Code_End = "end"
    this_Code_Goto = "goto"
    this_Code_GoSub = "gosub"
    this_Code_Return = "return"
    this_Code_Stop = "stop"
    
    this_Game_Autorun = "game.autorun"
    this_Game_MessageShow = "message.show"
    this_Game_MessageClear = "message.clear"
    this_Game_HeroCreate = "hero.create"
    this_Game_HeroMoveLeft = "hero.moveleft"
    this_Game_HeroMoveRight = "hero.moveright"
    this_Game_HeroMoveUp = "hero.moveup"
    this_Game_HeroMoveDown = "hero.movedown"
    this_Game_Tile = "tile.style"
    this_Game_Pathway = "pathway.style"
    this_Game_PathwayDraw = "pathway.draw"
    this_Game_PathwayClear = "pathway.clear"
    this_Game_BlockShow = "block.show"
    this_Game_BlockHide = "block.hide"
    this_Game_HeroShow = "hero.show"
    this_Game_HeroHide = "hero.hide"
    this_Game_EventShow = "event.show"
    this_Game_EventHide = "event.hide"
    this_Game_PathwayShow = "pathway.show"
    this_Game_PathwayHide = "pathway.hide"
    this_Game_ObjectShow = "object.show"
    this_Game_ObjectHide = "object.hide"
    this_Game_EffectShow = "effect.show"
    this_Game_EffectHide = "effect.hide"
    this_Game_EventCreate = "event.create"
    this_Game_EventDestroy = "event.destroy"
    this_Game_EventCreate = "event.create"
    this_Game_EventDestroy = "event.destroy"
    this_Game_BlockCreate = "block.create"
    this_Game_BlockDestroy = "block.destroy"
    
    this_Game_ObjectCreate = "object.create"
    this_Game_ObjectDestroy = "object.destroy"
    this_Game_ObjectClear = "object.clear"
    
    this_Game_EffectCreate = "effect.create"
    this_Game_EffectDestroy = "effect.destroy"
    this_Game_EffectClear = "effect.clear"
    this_Game_EffectEnable = "effect.enable"
    this_Game_EffectDisable = "effect.disable"
    
    this_Game_FileOpen = "file.open"
    this_Game_FileAI = "file.ai"
    this_Game_HeroSteps = "hero.steps"
    this_Game_HeroMoveXY = "hero.movexy"
    
    this_Game_PathwayDestroy = "pathway.destroy"
    this_Game_PathwayWithBlock = "pathway.block"
    this_Game_HeroTalk = "hero.say"
    this_Game_EventClear = "event.clear"
    this_Game_BlockClear = "block.clear"
    
    this_Game_HeroOpen = "hero.open"
    this_Game_HeroCollect = "hero.collect"
    this_Game_GameTask = "game.task"
    this_Game_MessageClose = "message.close"
    this_Game_FaceUp = "hero.faceup"
    this_Game_FaceDown = "hero.facedown"
    this_Game_FaceLeft = "hero.faceleft"
    this_Game_FaceRight = "hero.faceright"
End Function

 '
 '=============================================================
'Describe: Automatically convert keywords to uppercase
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub CodeFormat(this_txtSource As Object)
    Dim s1 As String
    Dim s2 As String
    Dim s1s() As String
    Dim s2s() As String
    Dim ns1 As String
    Dim ns2 As String
    Dim i As Integer
    Dim tmp As String
    Dim spacenum As Integer
    With this_txtSource
        s1 = Left(.Text, .SelStart)
        s2 = Mid(.Text, .SelStart + 1, Len(.Text))
        s1s = Split(s1, vbCrLf)
        s2s = Split(s2, vbCrLf)
        If UBound(s1s) > 0 Then
            For i = 0 To UBound(s1s) - 1
                tmp = s1s(i)
                spacenum = Len(tmp) - Len(LTrim(tmp))
                tmp = String(spacenum, " ") & FormatPlus(tmp)
                ns1 = ns1 & tmp & vbCrLf
            Next
            ns1 = ns1 & s1s(UBound(s1s))
        Else
            ns1 = ns1 & s1
        End If
        
        If UBound(s2s) > 0 Then
            ns2 = ns2 & s2s(0) & vbCrLf
            For i = 1 To UBound(s2s)
                tmp = s2s(i)
                spacenum = Len(tmp) - Len(LTrim(tmp))
                If Trim(tmp) = "" Then
                    tmp = FormatPlus(tmp)
                Else
                    tmp = String(spacenum, " ") & FormatPlus(tmp)
                End If
                ns2 = ns2 & tmp & vbCrLf
            Next
            ns2 = Left(ns2, Len(ns2) - 2)
        Else
            ns2 = ns2 & s2
        End If
        .Text = ns1 & ns2
        .SelStart = Len(ns1)
    End With
End Sub

'=============================================================
'Describe: Automatically convert keywords to uppercase
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function FormatPlus(ByVal s As String)
    Dim cs As New clsSen
    Dim rw1, rw, ssi, R As String
    Dim kw() As String
    Dim i, j As Integer
    cs.strText = LcasePlus(s)
    If Trim(s) = "" Then
        FormatPlus = s
        Exit Function
    End If
    
    Do
        rw1 = cs.ReadWord
        If rw1 = "(" Or cs.LookNextWord = ")" Or cs.LookNextWord = ";" Or cs.LookNextWord = "," Then
            rw = rw & rw1
        Else
            rw = rw & rw1 & " "
        End If
        If Err.Number > 0 Then
            FormatPlus = s
            Err.Clear
            Exit Function
        End If
    Loop While Not cs.Eof
    If rw <> "" Then rw = Left(rw, Len(rw) - 1)
    rw = ReplacePlus(rw, "< >", "<>")
    rw = ReplacePlus(rw, "< =", "<=")
    rw = ReplacePlus(rw, "> =", ">=")
    s = rw
    Dim ss() As String
    ss = Split(s, """")
    If UBound(ss) Mod 2 = 1 Then
        FormatPlus = s
    Else
        For i = 0 To UBound(ss)
            kw = Split(m_Keyword, ",")
            ssi = LCase(ss(i))
            For j = 0 To UBound(kw)
                ssi = " " & ssi & " "
                ssi = Replace(ssi, " " & kw(j) & " ", " " & UCase(Left(kw(j), 1)) & Right(kw(j), Len(kw(j)) - 1) & " ")
                ssi = Mid(ssi, 2, Len(ssi) - 2)
            Next
            R = R & IIf(i Mod 2 = 0, ssi, ss(i))
            If i <> UBound(ss) Then R = R & """"
        Next
        FormatPlus = R
    End If
End Function

'
'=============================================================
'Describe: All converted to lowercase characters
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function LcasePlus(s As String)
    Dim ss() As String
    Dim R As String
    Dim i As Integer
     'Debug.Print s
    ss = Split(s, """")
    If UBound(ss) Mod 2 = 1 Then
        LcasePlus = s
    Else
        For i = 0 To UBound(ss)
            R = R & IIf(i Mod 2 = 0, LCase(ss(i)), ss(i))
            If i <> UBound(ss) Then R = R & """"
        Next
        LcasePlus = R
    End If
    'Debug.Print LcasePlus
End Function

'
'=============================================================
'Describe: Here we deal mainly with the grammar with if.
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Sub Code_a(this_SourceCode As String)
    'First time: Source code read-in arrays
    Dim tmpExp As String
    Dim iCount As Integer
    Dim sen As String
    Dim Exp(4) As String
    Dim tmplogic As String
    Dim s
    Dim bd As String
    
    For Each s In Split(this_SourceCode, vbCrLf)
        iCount = iCount + 1
        If Trim(s) <> "" Then
            iline = iline + 1
            ReDim Preserve Lines(iline)
            Lines(iline).phLine = iCount
            Lines(iline).strText = Trim(s)
            'Processing internal line numbers
            If IsNumeric(Split(Lines(iline).strText, this_Code_Space)(0)) Then
                Lines(iline).vLine = Split(Lines(iline).strText, " ")(0)
                Lines(iline).strText = Trim(Mid(Lines(iline).strText, InStr(s, this_Code_Space), Len(Lines(iline).strText)))
            End If
            'Ignore Comments
            If InstrPlus(Lines(iline).strText, this_Code_Notes) > 0 Then
                Lines(iline).strText = Left(Lines(iline).strText, InstrPlus(Lines(iline).strText, this_Code_Notes) - 1)
            End If
            If Err.Number > 0 Then
                ErrLine = iCount
                Exit Sub
            End If
            'Replace tabs with spaces
            Lines(iline).strText = Trim(Replace(Lines(iline).strText, Chr(9), "    "))
            '预加工逻辑上不规范的语句
            'if logic then exp1
            'if logic then exp1 else exp2
            
            '如果有If但是没有then的话，应该在同一行才对
            If Lines(iline).Head = this_Code_If Or Lines(iline).Body = this_Code_If Then
            
                sen = Lines(iline).strText
                '如果这一行没有then的话，则提示错误！
                If InstrPlus(LCase(sen), this_Code_Then) = 0 Then
                    ErrLine = iCount
                    Err.Raise 301, , "Bad if directive synatax"
'在这里应该退出所有的函数，但是我先标记在这里
                    Exit Sub
                    
                Else '如果有Then的话，则对语法进行格式化
                
                    tmpExp = Lines(iline).Body
                    tmplogic = Trim(Left(tmpExp, InstrPlus(tmpExp, this_Code_Then) - 1))
                    If Len(sen) - InstrPlus(sen, this_Code_Then) = 3 Then
                        'if logic then
                        '其实本来就应该有then，所以什么也不要做
                    '如果判断有Else的话
                    ElseIf InstrPlus(sen, this_Code_Else) > 0 Then
                        'if logic then exp1 else exp2
                        Exp(0) = Trim(Mid(tmpExp, InstrPlus(tmpExp, this_Code_Then) + 4, Len(tmpExp)))
                        '如果有else的话，需要把剩下在字符给Exp（1）
                        Exp(1) = Trim(Right(Exp(0), Len(Exp(0)) - InstrPlus(Exp(0), this_Code_Else) - 3))
                        Exp(0) = Trim(Left(Exp(1), InstrPlus(Exp(0), this_Code_Else) - 1))
                        '格式化为带一个空格的标准语句
                        Lines(iline).strText = this_Code_If & " " & tmplogic & " " & this_Code_Then
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = Exp(0)
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = this_Code_Else
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = Exp(1)
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = this_Code_EndIf
                    Else
                        '如果没有else的话
                        'if logic then exp1
                        Exp(0) = Trim(Mid(tmpExp, InstrPlus(tmpExp, this_Code_Then) + 4, Len(tmpExp)))
                        Lines(iline).strText = this_Code_If & " " & tmplogic & " " & this_Code_Then
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = Exp(0)
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = this_Code_EndIf
                    End If
                
                End If
                
            '如果有输入提示符的话
            ElseIf Lines(iline).Head = this_Code_Input Then
                'input "x",x;y
                tmpExp = Lines(iline).Body
                
                If Right(tmpExp, 1) = ";" Or Right(tmpExp, 1) = "," Then
                    bd = Right(tmpExp, 1)
                    tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                Else
                    bd = vbCrLf
                End If
                
                If FindBorder(tmpExp) > 0 Then
                    'Exit Do
                    Lines(iline).strText = this_Code_Input & " " & Left(tmpExp, FindBorder(tmpExp) - 1)
                    tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))

                    Do While FindBorder(tmpExp) > 0
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = this_Code_Input & " " & Left(tmpExp, FindBorder(tmpExp) - 1)
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                    Loop
                    
                    If tmpExp <> "" Then
                        iline = iline + 1
                        ReDim Preserve Lines(iline)
                        Lines(iline).phLine = iCount
                        Lines(iline).strText = this_Code_Input & " " & tmpExp
                    End If
                End If
            ElseIf Lines(iline).Head = this_Code_While Then
                    tmpExp = Lines(iline).Body
                    Lines(iline).strText = this_Code_DoWhile & " " & tmpExp
            ElseIf Lines(iline).Head = this_Code_Wend Then
                    Lines(iline).strText = this_Code_Loop
            End If
        End If
    Next
End Sub

'=============================================================
'Describe: Second time: Preprocess - Select Case statement to change if.. ElseIf.. Format
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Sub Code_b(this_SourceCode As String)
    '
    Dim skSelect As New clsStack        'select case ... (case ...) ...end select 栈
    Dim i As Integer
    Dim tmpExp As String
    Dim tl As String
    Dim ad As String
    Dim Exp(4) As String
    Dim rexp As String
    Dim ei
    For i = 1 To UBound(Lines)
        ErrLine = Lines(i).phLine
        If Lines(i).Head = this_Code_SelectCase Then
            If i = UBound(Lines) Then
                Err.Raise 350, , "After 'Select Case', 'Case'is needed"
                Exit Sub
            End If
            If Lines(i + 1).Head <> this_Code_Case Then
                Err.Raise 351, , "After 'Select Case', 'Case'is needed"
                Exit Sub
            End If
            skSelect.Push Lines(i).Body
            Lines(i).strText = this_Code_SelectCase & " "
        ElseIf Lines(i).Head = this_Code_Case Then
            tmpExp = Lines(i).Body
            If skSelect.Count = 0 Then
                Err.Raise 352, , "Before 'Case', 'Select Case'is needed"
                Exit Sub
            End If
            tl = skSelect.Item(0)
            If Lines(i - 1).Head = this_Code_SelectCase Then
                ad = this_Code_If & " "
            Else
                ad = this_Code_ElseIf & " "
            End If
            If InstrPlus(tmpExp, " " & this_Code_To & " ") > 0 Then
                'exp1 to exp2
                Exp(0) = Trim(Left(tmpExp, InstrPlus(tmpExp, " " & this_Code_To & " ")))
                Exp(1) = Trim(Mid(tmpExp, InstrPlus(tmpExp, " " & this_Code_To & " ") + 3, Len(tmpExp)))
                rexp = ad & tl & ">=" & Exp(0) & " and " & tl & "<=" & Exp(1)
            Else
                'exp[,exp2,...expn]
                If InStr(tmpExp, """") > 0 Then
                    Err.Raise 353, , "String is not supported after 'Case' "
                    Exit Sub
                End If
                rexp = ad
                For Each ei In Split(tmpExp, ",")
                    rexp = rexp & tl & "=" & ei & " " & this_Code_Or & " "
                Next
                rexp = Left(rexp, Len(rexp) - 4)
            End If
            Lines(i).strText = rexp & " " & this_Code_Then
        ElseIf Lines(i).Head = this_Code_CaseElse Then
            Lines(i).strText = this_Code_Else
        ElseIf Lines(i).Head = this_Code_EndSelect Then
            Lines(i).strText = this_Code_EndIf
            If skSelect.Count = 0 Then
                Err.Raise 352, , "End Select need 'Select Case'"
                Exit Sub
            End If
            skSelect.Pop
        End If
    Next
    If skSelect.Count > 0 Then
        Err.Raise 353, , "After 'Select' 'End Select' is needed"
        Exit Sub
    End If

End Sub

'=============================================================
'Describe: Third time: Preprocessing - Expansion if... ElseIf... End If becomes nested
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Sub Code_c(this_SourceCode As String)
    '
    Dim Lines1() As New clsLine
    Dim skElseif As New clsStack
    ReDim Preserve Lines1(0)
    Dim i1 As Integer
    Dim i, k As Integer
    Dim ts() As String
    
    i1 = 0
    For i = 1 To UBound(Lines)
        With Lines(i)
            ErrLine = .phLine
            If Trim(.strText <> "") Then
                i1 = i1 + 1
                Select Case .Head
                    Case this_Code_If
                        skElseif.Push i & " 1"
                        
                        ReDim Preserve Lines1(i1)
                        Lines1(i1).strText = .strText
                        Lines1(i1).vLine = .vLine
                        Lines1(i1).phLine = .phLine
                        
                    Case this_Code_EndIf
                        If skElseif.Count = 0 Then
                            Err.Raise 362, , "Else If needs 'If'"
                            Exit Sub
                        End If
                        ts = Split(skElseif.Item(0), " ")
                        i1 = i1 - 1
                        For k = 1 To Int(ts(1))
                            i1 = i1 + 1
                            ReDim Preserve Lines1(i1)
                            Lines1(i1).strText = this_Code_EndIf
                            Lines1(i1).vLine = .vLine
                            Lines1(i1).phLine = .phLine
                        Next
                    Case this_Code_ElseIf
                        If skElseif.Count = 0 Then
                            Err.Raise 361, , "Else If needs 'If'"
                            Exit Sub
                        End If
                        ts = Split(skElseif.Pop, " ")
                        skElseif.Push ts(0) & " " & CStr(CInt(ts(1)) + 1)
                        
                        ReDim Preserve Lines1(i1)
                        Lines1(i1).strText = this_Code_Else
                        Lines(i1).vLine = .vLine
                        Lines(i1).phLine = .phLine
                        i1 = i1 + 1
                        ReDim Preserve Lines1(i1)
                        Lines1(i1).strText = this_Code_If & " " & .Body
                        Lines1(i1).vLine = .vLine
                        Lines1(i1).phLine = .phLine
                    Case Else
                        ReDim Preserve Lines1(i1)
                        Lines1(i1).strText = .strText
                        Lines1(i1).vLine = .vLine
                        Lines1(i1).phLine = .phLine
                End Select
            End If
        End With
    Next
    Lines = Lines1

End Sub

'=============================================================
'Describe: Fourth time: Preprocessing - abstracting lines of code into graphed nodes and linking them
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Code_d(this_SourceCode As String)
    '
    Dim sk1 As New clsStack             'if ...(else)...end if栈
    Dim skFor As New clsStack           'for ...(exit for)...next 栈
    Dim skWhile As New clsStack         'while ...(exit do)... wend 栈
    Dim skDo As New clsStack            'do ...(exit do)... loop until 栈
    Dim tmp  As String
    Dim i As Integer
    
    For i = 1 To UBound(Lines)
        ErrLine = Lines(i).phLine
        Select Case Lines(i).Head
            Case ""
                'ce.exp Lines(i).strText
            Case this_Code_If
                'if_else_id
                sk1.Push i
            Case this_Code_Else
                'else_if
                If sk1.Count > 0 Then
                    tmp = sk1.Item(0)
                    Lines(i).else_if_id = tmp
                    Lines(tmp).if_else_id = i
                Else
                    Err.Raise 601, , "Else needs If"
                    Exit Sub
                End If
            Case this_Code_EndIf
                If sk1.Count > 0 Then
                    tmp = sk1.Pop
                    Lines(tmp).if_endif_id = i
                Else
                    Err.Raise 602, , "End If needs If"
                    Exit Sub
                End If
            Case this_Code_For
                skFor.Push i
            Case this_Code_ExitFor
                If skFor.Count > 0 Then
                    tmp = skFor.Item(0)
                    Lines(i).exitfor_for_id = tmp
                Else
                    Err.Raise 300, , "Exit For error"
                    Exit Sub
                End If
            Case this_Code_Next
                If skFor.Count > 0 Then
                    tmp = skFor.Pop
                    Lines(i).next_for_id = tmp
                    Lines(tmp).for_next_id = i
                Else
                    Err.Raise 302, , "Next needs For"
                    Exit Sub
                End If
            Case this_Code_DoWhile
                'dowhile_loop
                skWhile.Push i
            Case this_Code_Loop
                'loop_dowhile
                If skWhile.Count > 0 Then
                    tmp = skWhile.Pop
                    Lines(i).loop_dowhile = tmp
                    Lines(tmp).dowhile_loop = i
                Else
                    Err.Raise 303, , "Loop needs Do While"
                    Exit Sub
                End If
            Case this_Code_Do
                'do_loopuntil_id
                skDo.Push i
            Case this_Code_LoopUntil
                'loopuntile_do_id
                If skDo.Count > 0 Then
                    tmp = skDo.Pop
                    Lines(i).loopuntile_do_id = tmp
                    Lines(tmp).do_loopuntil_id = i
                Else
                    Err.Raise 305, , "Loop Until needs Do"
                    Exit Sub
                End If
            Case this_Code_ExitDo
                'exitdo_dowhile
                'exitdo_do
                If skWhile.Count > 0 And skDo.Count > 0 Then
                    If skWhile.Item(0) > skDo.Item(0) Then
                        tmp = skWhile.Item(0)
                        Lines(i).exitdo_dowhile = tmp
                        Lines(tmp).dowhile_loop = i
                    Else
                        tmp = skDo.Item(0)
                        Lines(i).exitdo_do = tmp
                        Lines(tmp).do_loopuntil_id = i
                    End If
                ElseIf skWhile.Count > 0 Then
                    tmp = skWhile.Item(0)
                    Lines(i).exitdo_dowhile = tmp
                    Lines(tmp).dowhile_loop = i
                ElseIf skDo.Count > 0 Then
                    tmp = skDo.Item(0)
                    Lines(i).exitdo_do = tmp
                    Lines(tmp).do_loopuntil_id = i
                Else
                    Err.Raise 304, , "Exit Do error"
                    Exit Sub
                End If
            Case ""
            
            Case this_Code_Print
                
        End Select
    Next
    
    If skFor.Count > 0 Then
        Err.Raise 303, , "For needs Next"
        Exit Sub
    End If
    
    If sk1.Count > 0 Then
        Err.Raise 304, , "If needs End If"
        Exit Sub
    End If
    
    If skDo.Count > 0 Then
        Err.Raise 305, , "Do needs Loop Until"
        Exit Sub
    End If
    
    If skWhile.Count > 0 Then
        Err.Raise 306, , "While needs Wend or Do Until needs Loop"
        Exit Sub
    End If
End Sub



'=============================================================
'Describe: Fifth time: Explain and execute procedures
'Author:   Chen Lyu
'Parameter:
'=============================================================
Private Sub Code_e(this_SourceCode As String)

'
    Dim sksub As New clsStack
    Dim runId As Long
    Dim ct As String
    Dim var1 As String
    Dim exp3 As String
    Dim varn  As String
    Dim ivar  As String
    Dim tftf As Long
    Dim T1 As String
    Dim rt As String
    Dim cline As Integer
    Dim tmpExp As String
    Dim Exp(4) As String
    Dim bd As String
    Dim i As Long
    Dim t() As String
    Dim s As String
    Dim n As String
    Dim a, b, c, d As Integer
    Dim M As m_Position

    
    Static lastStopLine As Integer
    runId = 1
'    this_Graphic.SpriteFont(0).G_String = ""
'    this_Graphic.SpriteFont(1).G_String = ""
    s = ""
    fMainForm.Txt_Message.Text = ""
    
    Do While runId <= UBound(Lines)

    
        DoEvents
        ErrLine = Lines(runId).phLine
        If BanHave(ErrLine) Or this_Graphic.Code.runFlag = False Then
            If lastStopLine <> ErrLine Then             '防止冗余的逻辑行导致物理行中断多次。
                frmVars.Show 1
            End If
        End If
        lastStopLine = ErrLine
        Select Case Lines(runId).Head
            Case this_Game_Autorun
                'Debug.Print this_Game_Autorun
                Err.Number = 0
            Case this_Game_MessageShow
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    'tmpExp = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        bd = Mid(tmpExp, FindBorder(tmpExp), 1)
                        If bd = "," Then
                            bd = Chr(9)
                        Else
                            bd = ""
                        End If
:                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        
                        If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                            s = Mid(T1, 2, Len(T1) - 2)
                            
                            '输出到控制台
                        Else
                            s = ce.Exp(T1) & bd
                        End If
                        
                        
                        If InStr(s, "\n ") > 0 Then
                            s = ReplacePlus(s, "\n ", vbNewLine)
                        End If

                        
                        Select Case tmpExp
                        
                            Case 1
                                this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                                this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                                this_Graphic.SpriteFont(0).G_WordWraped = False
                                'fMainForm.Txt_Message.Text = this_Graphic.SpriteFont(0).G_String & vbNewLine '& s
                                Call Event_Get(this_Graphic, 1)
                            Case 6 '作为任务显示在这里
                                this_Graphic.SpriteFont(4).G_String = s
                                this_Graphic.SpriteFont(4).G_MaxLineLen = 500
                                this_Graphic.SpriteFont(4).G_WordWraped = False
                                fMainForm.Txt_Message.Text = this_Graphic.SpriteFont(4).G_String & vbNewLine
                                Call Event_Get(this_Graphic, 6)
                                
                            Case 7 '作为结束语言显示在这里
                                this_Graphic.SpriteFont(5).G_String = s
                                this_Graphic.SpriteFont(5).G_MaxLineLen = 500
                                this_Graphic.SpriteFont(5).G_WordWraped = False
                                'fMainForm.Txt_Message.Text = this_Graphic.SpriteFont(5).G_String & vbNewLine
                                Call Event_Get(this_Graphic, 7)
                                
                            Case Else
                                If InStr(Trim(tmpExp), ",") > 0 Then
                                    t = Split(LCase(tmpExp), ",")
                                    If UBound(t) = 3 Then
                                        If t(3) = 9 Then
                                            If CInt(t(0)) >= 0 And CInt(t(0)) <= 16 And CInt(t(1)) >= 0 And CInt(t(1)) <= 14 Then
                                                M.x = CInt(t(0))
                                                M.y = CInt(t(1))
                                                If CInt(t(2)) > 0 And CInt(t(2)) < 4 Then
                                                Else
                                                    t(2) = 0
                                                End If
                                                s = Replace(Trim(Left(s, 16 - M.x)), Chr(9), "")
                                                
'                                                Call Event_Get(this_Graphic, 63)
                                                Call Draw_TilesBoardNote(this_Graphic, s, 0, 0, M, CInt(t(2)))
                                            End If
                                        Else
                                        
                                        End If
                                    Else
                                    
                                    End If
                                End If
                            
                            
                        End Select
                    Else
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        If tmpExp = "" Then
                            Call Event_Get(this_Graphic, 1)
                        Else
                            If Left(tmpExp, 1) = """" And Right(tmpExp, 1) = """" Then
                                s = Mid(tmpExp, 2, Len(tmpExp) - 2)
                                
                                '输出到控制台
                            Else
                                s = ce.Exp(T1) & bd
                            End If
                            
                            
                            If InStr(s, "\n ") > 0 Then
                                s = ReplacePlus(s, "\n ", vbNewLine)
                            End If

'                            Call Event_Get(this_Graphic, 63)
                            this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                            this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                            this_Graphic.SpriteFont(0).G_WordWraped = False
                            fMainForm.Txt_Message.Text = this_Graphic.SpriteFont(0).G_String & vbNewLine '& s
                            Call Event_Get(this_Graphic, 1)
 
                        End If
                        
                        

                        
                    End If
                    
                    
                    
            Case this_Game_FileOpen
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    'tmpExp = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        bd = Mid(tmpExp, FindBorder(tmpExp), 1)
                        If bd = "," Then
                            bd = Chr(9)
                        Else
                            bd = ""
                        End If
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        
                        If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                            s = Mid(T1, 2, Len(T1) - 2)
                            
                            '输出到控制台
                        Else
                            s = ce.Exp(T1) & bd
                        End If
                        
                        
                        Select Case tmpExp
                        
                            Case 0
                                Call CodeFileLoad(this_Graphic, s & ".pl")
                            Case Else
                                 Err.Raise 209, , "File open error"
                        End Select
                    Else
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        If Left(tmpExp, 1) = """" And Right(tmpExp, 1) = """" Then
                            s = Mid(tmpExp, 2, Len(tmpExp) - 2)
                            
                            '输出到控制台
                        Else
                            s = ce.Exp(tmpExp) & bd
                        End If
                    
                        If s <> "" Then
                            Call CodeFileLoad(this_Graphic, s & ".pl")
                        Else
                            Err.Raise 209, , "File open error"
                        End If

                    End If
                    
            Case this_Game_Tile
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Game_Tile setup error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 4 Then
                            .Map.TileID = CInt(tftf)
                        Else
                                    Err.Raise 209, , "Game_Tile setup error"
                        End If
                        
                    Else
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 4 Then
                            .Map.TileID = CInt(tftf)
                        Else
                                    Err.Raise 209, , "Game_Tile setup error"
                        End If
                    End If
                End With
                
            Case this_Game_FileAI
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "AI setup error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 250 Then
                            Call AI_Main(CByte(tftf))
                        Else
                                    Err.Raise 209, , "AI setup error"
                        End If
                        
                    Else
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 4 Then
                            .Map.TileID = CInt(tftf)
                        Else
                                    Err.Raise 209, , "AI setup error"
                        End If
                    End If
                End With
                
                
             Case this_Game_PathwayWithBlock
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        this_Switch.PathwayWithBlock = True
                    ElseIf IsNumeric(tmpExp) = True Then
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf >= 0 And tftf < 2 Then
                            If tftf = 0 Then
                                this_Switch.PathwayWithBlock = False
                                
                            Else
                                this_Switch.PathwayWithBlock = True
                                
                            End If
                            Call Event_Get(this_Graphic, 63)
                        Else
                                    Err.Raise 209, , "Pathway With Block setup error"
                        End If
                        
                    Else
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf >= 0 And tftf < 2 Then
                            If tftf = 0 Then
                                this_Switch.PathwayWithBlock = False
                            Else
                                this_Switch.PathwayWithBlock = True
                               
                            End If
                            Call Event_Get(this_Graphic, 63)
                        Else
                                    Err.Raise 209, , "Pathway With Block  setup error"
                        End If
                    End If
                End With
                
                
                                
             Case this_Game_HeroTalk
'                With this_Graphic
'                    tmpExp = Lines(runId).Body
'                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
'                    tftf = FindBorder(tmpExp)
'                    If Err.Number > 0 Then Exit Sub
'                    If tmpExp = "" Then
'                        this_Switch.Talk = True
'                        Call Event_Get(this_Graphic, 3)
'                    ElseIf IsNumeric(tmpExp) = True Then
'                        tftf = FindNumber(tmpExp)
'                        If Err.Number > 0 Then Exit Sub
'                        If tftf >= 0 And tftf < 2 Then
'                            If tftf = 0 Then
'                                this_Switch.Talk = False
'                                Call Event_Get(this_Graphic, 63)
'                            Else
'                                this_Switch.Talk = True
'                                Call Event_Get(this_Graphic, 3)
'                            End If
'                        Else
'                                    Err.Raise 209, , "Hero speak setup error"
'                        End If
'
'                    Else
'                        tftf = FindNumber(tmpExp)
'                        If Err.Number > 0 Then Exit Sub
'                        If tftf >= 0 And tftf < 2 Then
'                            If tftf = 0 Then
'                                this_Switch.Talk = False
'                                Call Event_Get(this_Graphic, 63)
'                            Else
'                                this_Switch.Talk = True
'                                Call Event_Get(this_Graphic, 3)
'                            End If
'                        Else
'                                    Err.Raise 209, , "Hero speak setup error"
'                        End If
'                    End If
'                End With
                
                
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    'tmpExp = FindNumber(tmpExp)
                    
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        bd = Mid(tmpExp, FindBorder(tmpExp), 1)
                        If bd = "," Then
                            bd = Chr(9)
                        Else
                            bd = ""
                        End If
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        
                        If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                            s = Mid(T1, 2, Len(T1) - 2)
                            
                            '输出到控制台
                        Else
                            s = ce.Exp(T1) & bd
                        End If
                        
                        
                        If InStr(s, "\n ") > 0 Then
                            s = ReplacePlus(s, "\n ", vbNewLine)
                        End If

                        
                        Select Case tmpExp
                            Case 0
                                this_Switch.Talk = False
                                
                            Case 1
                                this_Graphic.SpriteFont(2).G_String = s
                                this_Graphic.SpriteFont(2).G_MaxLineLen = 160
                                this_Graphic.SpriteFont(2).G_WordWraped = False
'                                fMainForm.Txt_Message.Text = fMainForm.Txt_Message.Text & vbNewLine & this_Graphic.SpriteFont(2).G_String
                                Call Event_Get(this_Graphic, 3)
                                this_Switch.Talk = True
                            Case Else
                                Err.Raise 209, , "Hero speak setup error"
                        End Select
                    
                    
                    Else
                        If IsNumeric(tmpExp) = True Then
                            tftf = FindNumber(tmpExp)
                            If Err.Number > 0 Then Exit Sub
                            If tftf >= 0 And tftf < 2 Then
                                If tftf = 0 Then
                                    this_Switch.Talk = False
                                    Call Event_Get(this_Graphic, 63)
                                Else
                                    this_Switch.Talk = True
                                    Call Event_Get(this_Graphic, 3)
                                End If
                            End If
                        Else
                        
                            Err.Raise 209, , "Hero speak setup error"
                        End If
                    End If
                     
                    
                
             Case this_Game_Pathway
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Game_Tile setup error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 4 Then
                            .Map.BlockTilesID = CInt(tftf) - 1
                            b = .Map.BlockTilesID
                            a = 0
                            For c = 0 To 3
                                For d = 0 To 2
                                    .Map.BlockTilesInfo(a).GraphicPosition.x = .Map.Tiles_Posi(b).x + .Map.BlockTilesInfo(a).GraphicPosition.Width * d
                                    If c < 3 Then
                                        .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(a).GraphicPosition.Height * c
                                    Else
                                        .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(0).GraphicPosition.Height * c
                                    End If
                                    a = a + 1
                                Next d
                            Next c
                        Else
                                    Err.Raise 209, , "Game_Tile setup error"
                        End If
                        
                    Else
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 And tftf < 4 Then
                            .Map.BlockTilesID = CInt(tftf) - 1
                            
                            b = .Map.BlockTilesID
                            a = 0
                            For c = 0 To 3
                                For d = 0 To 2
                                    .Map.BlockTilesInfo(a).GraphicPosition.x = .Map.Tiles_Posi(b).x + .Map.BlockTilesInfo(a).GraphicPosition.Width * d
                                    If c < 3 Then
                                        .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(a).GraphicPosition.Height * c
                                    Else
                                        .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(0).GraphicPosition.Height * c
                                    End If
                                    a = a + 1
                                Next d
                            Next c
                            
                        Else
                                    Err.Raise 209, , "Game_Tile setup error"
                        End If
                    End If
                End With
                
                
            Case this_Game_PathwayDraw

                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 3 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                t(2) = FindNumber(t(2))
                                t(3) = FindNumber(t(3))
                                If t(0) >= 0 And t(0) <= 16 And t(2) >= 0 And t(2) <= 16 And t(1) >= 0 And t(1) <= 14 And t(3) >= 0 And t(3) <= 14 Then
                                    For a = t(0) To t(2)
                                        For b = t(1) To t(3)
                                            .Map.BlockTiles(a, b) = 1
                                        Next b
                                    Next a
                                Else
                                    Err.Raise 209, , "Draw pathway error"
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Draw pathway error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 3 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                t(2) = FindNumber(t(2))
                                t(3) = FindNumber(t(3))
                                    If t(0) >= 0 And t(0) <= 16 And t(2) >= 0 And t(2) <= 16 And t(1) >= 0 And t(1) <= 14 And t(3) >= 0 And t(3) <= 14 Then
                                        For a = t(0) To CByte(t(2)) + CByte(t(0)) - 1
                                            For b = t(1) To CByte(t(1)) + CByte(t(3)) - 1
                                                .Map.BlockTiles(a, b) = 1
                                            Next b
                                        Next a
                                    Else
                                        Err.Raise 209, , "Draw pathway error"
                                    End If
                            Else
                                Err.Raise 209, , "Draw pathway error"
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Draw pathway error"
                        End If
                    End If
                End With
                
                
            Case this_Game_BlockShow
                this_Switch.Block = True

            Case this_Game_BlockHide
                this_Switch.Block = False
                
                
            Case this_Game_HeroShow
                this_Switch.Player = True

            Case this_Game_HeroHide
                this_Switch.Player = False
                
            Case this_Game_EventShow
                this_Switch.Event = True

            Case this_Game_EventHide
                this_Switch.Event = False
                
            Case this_Game_PathwayShow
                this_Switch.Pathway = True

            Case this_Game_PathwayHide
                this_Switch.Pathway = False
                
            Case this_Game_ObjectShow
                this_Switch.Object = True

            Case this_Game_ObjectHide
                this_Switch.Object = False
                
            Case this_Game_EffectShow
                this_Switch.Effect = True

            Case this_Game_EffectHide
                this_Switch.Effect = False

                
            Case this_Game_PathwayClear
                For a = 0 To 15
                    For b = 0 To 13
                        this_Graphic.Map.BlockTiles(a, b) = 0
                    Next b
                Next a

            Case this_Game_ObjectClear
                For a = 0 To 31
                    For b = 0 To 31
                        this_Graphic.Map.Objects(a).MapPosition(b).x = -1
                        this_Graphic.Map.Objects(a).MapPosition(b).y = -1
                    Next b
                Next a
                        
            Case this_Game_EventClear
                For a = 0 To 15
                    For b = 0 To 13
                        this_Graphic.Map.Events(a, b) = 0
                        this_Graphic.Map.Events(a, b) = 0
                    Next b
                Next a
                
            Case this_Game_BlockClear
                For a = 0 To 15
                    For b = 0 To 13
                        this_Graphic.Map.Block(a, b) = 0
                        this_Graphic.Map.Block(a, b) = 0
                    Next b
                Next a
                
             Case this_Game_EffectClear
                For a = 0 To 7
                    For b = 0 To 1
                        this_Graphic.Map.Effect(a).DrawPosition(b).x = -1
                        this_Graphic.Map.Effect(a).DrawPosition(b).y = -1
                        this_Graphic.Map.Effect(a).FrameRun(b) = False
                        this_Graphic.Map.Effect(a).Visible(b) = False
                    Next b
                Next a
                                       
                                       
            Case this_Game_MessageClear
                With this_Graphic
                
                    For i = 0 To 1
                        .SpriteFont(i).G_String = ""
                        .SpriteFont(i).G_MaxLineLen = 500
                        .SpriteFont(i).G_WordWraped = False
                    Next i
                    
                     For b = 0 To 51
                        For a = 0 To 31
                            .Map.Letter(b).MapPosition(a).x = -1
                            .Map.Letter(b).MapPosition(a).y = -1
                            If b < 26 Then
                                For c = 0 To 3
                                    .Map.Letters(c, b).MapPosition(a).x = -1
                                    .Map.Letters(c, b).MapPosition(a).y = -1
                                Next c
                            End If
                        Next a
                    Next b
                End With
'                Call Event_Get(this_Graphic, 63)
            
            Case this_Game_MessageClose
 
                Call Event_Get(this_Graphic, 63)
                
            Case this_Game_HeroCreate
            
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                With this_Graphic
                                    .Player(CurrentPlayerID).Info.Alive = True
                                    .Player(CurrentPlayerID).Info.MoveDirection = MoveDown
                                    .Player(CurrentPlayerID).Info.MoveSpeed = 0.2
                                    .Player(CurrentPlayerID).Info.PositionTimer = 0
                                    .Player(CurrentPlayerID).Info.G_Position.x = FindNumber(t(0))
                                    .Player(CurrentPlayerID).Info.G_Position.y = FindNumber(t(1))
                                    .Player(CurrentPlayerID).Info.C_Position.x = FindNumber(t(0))
                                    .Player(CurrentPlayerID).Info.C_Position.y = FindNumber(t(1))
                                    Steps(0).Position.x = FindNumber(t(0))
                                    Steps(0).Position.y = FindNumber(t(1))
                                    .Player(CurrentPlayerID).Info.MagicBall = False
                                    .Player(CurrentPlayerID).Info.MagicBallTimer = 8
                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                        .Player(CurrentPlayerID).mName = Mid(T1, 2, Len(T1) - 2)
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "The Hero needs a coordinate"
                            End If
                        End If
                        
                    Else
                    
                    End If
                    
             Case this_Game_HeroSteps
            
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do

                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) > 0 Then
                                With this_Graphic
                                    If UBound(t) > 0 And UBound(t) < 100 Then
                                        For a = 0 To 99
                                            Steps(a).Direction = -1
                                        Next a
                                        For a = 0 To UBound(t)
                                            Steps(a).Direction = t(a)
                                        Next
                                        CurrentSteps = UBound(t) + 1
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "The Hero steps needs a coordinate"
                            End If
                        Else
                             Err.Raise 209, , "The Hero steps needs a coordinate"
                        End If
                    
              Case this_Game_EventCreate
                CurrentGameEvents = 0
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                With this_Graphic
                                   
                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                        .Map.GameEvents(t(2)).m_Description = Mid(T1, 2, Len(T1) - 2)
                                    ElseIf IsNumeric(T1) = True Then
                                        s = FindNumber(T1)
                                         .Map.Events(t(0), t(1)) = s
                                    Else
                                        Err.Raise 209, , "Create Event error"
                                    End If
                                End With
                            ElseIf UBound(t) = 2 Then
                                With this_Graphic
                                   
                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                        s = FindNumber(Trim(t(0)))
                                        If s >= 0 And s <= 63 Then
                                            t(1) = FindNumber(Trim(t(1)))
                                            t(2) = FindNumber(Trim(t(2)))
                                             .Map.Events(t(1), t(2)) = s
                                            .Map.GameEvents(s).m_Description = Mid(T1, 2, Len(T1) - 2)
                                         End If
                                    ElseIf IsNumeric(T1) = True Then
                                        s = FindNumber(T1)
                                         .Map.Events(t(0), t(1)) = s
                                    Else
                                        Err.Raise 209, , "Create Event error"
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "Create Event error"
                            End If
                        End If
                        
                    Else
                        Err.Raise 209, , "Create Event error"
                    End If
 
             Case this_Game_EventDestroy
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                    .Map.Events(t(0), t(1)) = 0
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Event destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                    .Map.Events(t(0), t(1)) = 0
                                End If
                            Else
                                Err.Raise 209, , "Event destroy error"
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Event destroy error"
                        End If
                    End If
                End With
                    
                    
             Case this_Game_PathwayDestroy
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Pathway destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                    .Map.BlockTiles(t(0), t(1)) = 0
                                    .Map.Block(t(0), t(1)) = 0
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Pathway destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                t(0) = FindNumber(t(0))
                                t(1) = FindNumber(t(1))
                                If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                    .Map.BlockTiles(t(0), t(1)) = 0
                                    .Map.Block(t(0), t(1)) = 0
                                End If
                            Else
                                Err.Raise 209, , "Pathway destroy error"
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Pathway destroy error"
                        End If
                    End If
                End With
                
                    
                    
                    
                    
                    
              Case this_Game_HeroMoveXY
                    With this_Graphic
                        tmpExp = Lines(runId).Body
                        If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                        tftf = FindBorder(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tmpExp = "" Then
                            Err.Raise 209, , "Block create error"
                        ElseIf IsNumeric(tmpExp) = True Then
                            s = FindNumber(tmpExp)
                            If Err.Number > 0 Then Exit Sub
                            If InStr(Trim(s), ",") > 0 Then
                                t = Split(LCase(s), ",")
                                If UBound(t) = 1 Then
                                    If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                        this_Graphic.Player(CurrentPlayerID).Info.G_Position.x = t(0)
                                        this_Graphic.Player(CurrentPlayerID).Info.G_Position.y = t(1)
                                        this_Graphic.Player(CurrentPlayerID).Info.Next_Position.x = -1
                                        this_Graphic.Player(CurrentPlayerID).Info.Next_Position.y = -1
                                        Call Player_Move_Test(this_Graphic, CurrentPlayerID)
                                        this_Graphic.Player(CurrentPlayerID).Info.RoadXY = True
                                        Call Event_Get(this_Graphic, 63)
                                    End If
                                End If
                            Else
                                        Err.Raise 209, , "Block create error"
                            End If
                            
                        Else
                            s = FindNumber(tmpExp)
                            If InStr(Trim(s), ",") > 0 Then
                                t = Split(LCase(s), ",")
                                If UBound(t) = 1 Then
                                    If t(0) >= 0 And t(0) <= 16 And t(1) >= 0 And t(1) <= 14 Then
                                        this_Graphic.Player(CurrentPlayerID).Info.G_Position.x = t(0)
                                        this_Graphic.Player(CurrentPlayerID).Info.G_Position.y = t(1)
                                        this_Graphic.Player(CurrentPlayerID).Info.Next_Position.x = -1
                                        this_Graphic.Player(CurrentPlayerID).Info.Next_Position.y = -1
                                        Call Player_Move_Test(this_Graphic, CurrentPlayerID)
                                        this_Graphic.Player(CurrentPlayerID).Info.RoadXY = True
                                        Call Event_Get(this_Graphic, 63)
                                    End If
                                Else
                                    Err.Raise 209, , "Block create error"
                                End If
                                If Err.Number > 0 Then Exit Sub
                            Else
                                Err.Raise 209, , "Block create error"
                            End If
                        End If
                    End With
                    
              Case this_Game_BlockCreate
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Block create error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                .Map.Block(t(0), t(1)) = 1
                            End If
                        Else
                                    Err.Raise 209, , "Block create error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                .Map.Block(t(0), t(1)) = 1
                            Else
                                Err.Raise 209, , "Block create error"
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Block create error"
                        End If
                    End If
                End With
 
             Case this_Game_BlockDestroy
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                .Map.Block(t(0), t(1)) = 0
                            End If
                        Else
                                    Err.Raise 209, , "Event destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                .Map.Block(t(0), t(1)) = 0
                            Else
                                Err.Raise 209, , "Event destroy error"
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Event destroy error"
                        End If
                    End If
                End With
                
                    
             Case this_Game_ObjectCreate
            
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 2 Then
                                With this_Graphic
                                    '柱子

                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                         
                                    ElseIf IsNumeric(T1) = True Then
                                        s = FindNumber(T1)
                                        If t(1) >= 0 And t(1) < 32 And t(2) >= 0 And t(2) < 32 Then
                                            .Map.Objects(s).MapPosition(t(0)).x = t(1)
                                            .Map.Objects(s).MapPosition(t(0)).y = t(2)
                                             Call Event_Get(this_Graphic, 63)
                                        Else
                                            Err.Raise 209, , "Create Event error"
                                        End If
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "Create Event error"
                            End If
                        End If
                        
                    Else
                    
                    End If


             Case this_Game_ObjectDestroy
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then

                                If .Map.Objects(t(0)).MapPosition(t(1)).x > -1 And .Map.Objects(t(0)).MapPosition(t(1)).y > -1 Then
                                    If t(0) = 9 Then
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x + 1, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x + 2, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                    End If
                                End If
                                .Map.Objects(t(0)).MapPosition(t(1)).x = -1
                                .Map.Objects(t(0)).MapPosition(t(1)).y = -1
                                Call Event_Get(this_Graphic, 63)
                            End If
                        Else
                                    Err.Raise 209, , "Event destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If .Map.Objects(t(0)).MapPosition(t(1)).x > -1 And .Map.Objects(t(0)).MapPosition(t(1)).y > -1 Then
                                    If t(0) = 9 Then
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x + 1, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                        .Map.Block(.Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).x + 2, .Map.Objects(CInt(t(0))).MapPosition(CInt(t(1))).y) = 0
                                    End If
                                End If
                                .Map.Objects(t(0)).MapPosition(t(1)).x = -1
                                .Map.Objects(t(0)).MapPosition(t(1)).y = -1
                                Call Event_Get(this_Graphic, 63)
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Event destroy error"
                        End If
                    End If
                End With
                    
                    
             Case this_Game_FaceUp
                this_Graphic.Player(CurrentPlayerID).Info.MoveDirection = 3
             
              Case this_Game_FaceDown
                this_Graphic.Player(CurrentPlayerID).Info.MoveDirection = 0
                
             Case this_Game_FaceLeft
                this_Graphic.Player(CurrentPlayerID).Info.MoveDirection = 1
                
             Case this_Game_FaceRight
                this_Graphic.Player(CurrentPlayerID).Info.MoveDirection = 2
            
                
             Case this_Game_HeroOpen
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                        'Exit Do
                        tftf = FindBorder(tmpExp)
                        'tmpExp = FindNumber(tmpExp)
                        
                        If Err.Number > 0 Then Exit Sub
                        If tftf = 0 Then
 
                            
                            'Select Case tmpExp
                            Select Case .Player(CurrentPlayerID).Info.MoveDirection
                                Case 0
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y + 1 Then
                                                Select Case a
                                                
                                                    Case 9
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 1
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1 Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 2) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y Then
                                                Select Case a
                                                
                                                    Case 9
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 2
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x + 1 Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y Then
                                                Select Case a
                                                
                                                    Case 9
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 3
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y - 1 Then
                                                Select Case a
                                                
                                                    Case 9
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case Else
                                    Err.Raise 209, , "Hero open error"
                            End Select
                        
                        
                        Else
                            Err.Raise 209, , "Hero open error"

                        End If
                    End With
                        
             Case this_Game_HeroCollect
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                        'Exit Do
                        tftf = FindBorder(tmpExp)
                        'tmpExp = FindNumber(tmpExp)
                        
                        If Err.Number > 0 Then Exit Sub
                        If tftf = 0 Then
 
                            
                            'Select Case tmpExp
                            Select Case .Player(CurrentPlayerID).Info.MoveDirection
                                Case 0
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y + 1 Then
                                                Select Case a
                                                
                                                    Case 25
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 1
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1 Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 2) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y Then
                                                Select Case a
                                                
                                                    Case 25
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 2
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x + 1 Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y Then
                                                Select Case a
                                                
                                                    Case 25
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case 3
                                    For a = 0 To 31
                                        For b = 0 To 31
                                            If (.Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(a).MapPosition(b).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(a).MapPosition(b).y = .Player(CurrentPlayerID).Info.C_Position.y - 1 Then
                                                Select Case a
                                                
                                                    Case 25
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 1, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Block(.Map.Objects(a).MapPosition(b).x + 2, .Map.Objects(a).MapPosition(b).y) = 0
                                                        .Map.Objects(a).MapPosition(b).x = -1
                                                        .Map.Objects(a).MapPosition(b).y = -1
                                                End Select
                                            End If
                                        Next b
                                    Next a
                                Case Else
                                    Err.Raise 209, , "Hero open error"
                            End Select
                        
                        
                        Else
                            Err.Raise 209, , "Hero open error"

                        End If
                    End With
                                                
                                                
             Case this_Game_EffectCreate
            
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 3 Then
                                With this_Graphic
                                    '柱子

                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                         
                                    ElseIf IsNumeric(T1) = True Then
                                        s = FindNumber(T1)
                                        If s >= 0 And s <= 7 Then
                                            If t(1) >= 0 And t(1) <= 16 And t(2) >= 0 And t(2) <= 14 And t(0) >= 0 And t(0) <= 1 Then
                                                    .Map.Effect(s).DrawPosition(t(0)).x = t(1)
                                                    .Map.Effect(s).DrawPosition(t(0)).y = t(2)
                                                    .Map.Effect(s).Timer(t(0)) = t(3)
                                                    .EffectTimer(s) = 0
                                                    .Map.Effect(s).FrameRun(t(0)) = True
                                                    .Map.Effect(s).Visible(t(0)) = True
                                            Else
                                                Err.Raise 209, , "Create Effect error"
                                            End If
                                        End If
                                    End If
                                End With
                                
                            ElseIf UBound(t) = 2 Then
                                  With this_Graphic
                                    '柱子

                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                         
                                    ElseIf IsNumeric(T1) = True Then
                                        s = FindNumber(T1)
                                        If s <= 7 And s >= 0 Then
                                            If t(1) >= 0 And t(1) <= 16 And t(2) >= 0 And t(2) <= 14 And t(0) >= 0 And t(0) <= 1 Then
                                                .Map.Effect(s).DrawPosition(t(0)).x = t(1)
                                                .Map.Effect(s).DrawPosition(t(0)).y = t(2)
                                                .EffectTimer(s) = 0
                                                .Map.Effect(s).Timer(t(0)) = DefauleEffectTimer
                                                .Map.Effect(s).FrameRun(t(0)) = True
                                                .Map.Effect(s).Visible(t(0)) = True
                                                
                                            Else
                                                Err.Raise 209, , "Create Effect error"
                                            End If
                                        End If
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "Create Effect error"
                            End If
                        End If
                        
                    Else
                    
                    End If


             Case this_Game_EffectDestroy
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).DrawPosition(t(1)).x = -1
                                    .Map.Effect(t(0)).DrawPosition(t(1)).y = -1
                                    .Map.Effect(t(0)).Visible(t(1)) = False
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Effect destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                 If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).DrawPosition(t(1)).x = -1
                                    .Map.Effect(t(0)).DrawPosition(t(1)).y = -1
                                    .Map.Effect(t(0)).Visible(t(1)) = False
                                End If
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Effect destroy error"
                        End If
                    End If
                End With
                
              Case this_Game_EffectEnable
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).Visible(t(1)) = True
                                    .Map.Effect(t(0)).FrameRun(t(1)) = True
                                    .EffectTimer(t(0)) = 0
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Effect destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).Visible(t(1)) = True
                                    .Map.Effect(t(0)).FrameRun(t(1)) = True
                                    .EffectTimer(t(0)) = 0
                                End If
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Effect enable error"
                        End If
                    End If
                End With
                
              Case this_Game_EffectDisable
            
                With this_Graphic
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Err.Raise 209, , "Event destroy error"
                    ElseIf IsNumeric(tmpExp) = True Then
                        s = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).Visible(t(1)) = False
                                    .Map.Effect(t(0)).FrameRun(t(1)) = False
                                End If
                            End If
                        Else
                                    Err.Raise 209, , "Effect destroy error"
                        End If
                        
                    Else
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) = 1 Then
                                If t(0) <= 7 And t(0) >= 0 And t(1) <= 1 And t(1) >= 0 Then
                                    .Map.Effect(t(0)).Visible(t(1)) = False
                                    .Map.Effect(t(0)).FrameRun(t(1)) = False
                                End If
                            End If
                            If Err.Number > 0 Then Exit Sub
                        Else
                            Err.Raise 209, , "Effect disable error"
                        End If
                    End If
                End With
                    
             Case this_Game_HeroCreate
            
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        s = FindNumber(tmpExp)
                        If InStr(Trim(s), ",") > 0 Then
                            t = Split(LCase(s), ",")
                            If UBound(t) > 0 Then
                                With this_Graphic
                                    .Player(CurrentPlayerID).Info.Alive = True
                                    .Player(CurrentPlayerID).Info.MoveDirection = MoveDown
                                    .Player(CurrentPlayerID).Info.MoveSpeed = 0.2
                                    .Player(CurrentPlayerID).Info.PositionTimer = 0
                                    .Player(CurrentPlayerID).Info.C_Position.x = FindNumber(t(0))
                                    .Player(CurrentPlayerID).Info.C_Position.y = FindNumber(t(1))
                                    .Player(CurrentPlayerID).Info.MagicBall = False
                                    .Player(CurrentPlayerID).Info.MagicBallTimer = 8
                                    If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                                        .Player(CurrentPlayerID).mName = Mid(T1, 2, Len(T1) - 2)
                                    End If
                                End With
                            Else
                                Err.Raise 209, , "The Hero needs a coordinate"
                            End If
                        End If
                        
                    Else
                    
                    End If
                    
            Case this_Game_HeroMoveLeft
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                tftf = FindBorder(tmpExp)
                If Err.Number > 0 Then Exit Sub
                If tmpExp = "" Then
                    Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                ElseIf IsNumeric(Trim(tmpExp)) = True Then
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                Else
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                End If
                    
            Case this_Game_HeroMoveRight
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                tftf = FindBorder(tmpExp)
                If Err.Number > 0 Then Exit Sub
                If tmpExp = "" Then
                    Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                ElseIf IsNumeric(Trim(tmpExp)) = True Then
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                Else
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                End If
                
            Case this_Game_HeroMoveUp
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                tftf = FindBorder(tmpExp)
                If Err.Number > 0 Then Exit Sub
                If tmpExp = "" Then
                    Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                ElseIf IsNumeric(Trim(tmpExp)) = True Then
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                Else
                    tftf = FindNumber(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        For a = 0 To CInt(tftf) - 1
                            Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        Next a
                    Else
                                Err.Raise 209, , "Cannot move the Hero"
                    End If
                    
                End If
        
                Case this_Game_HeroMoveDown
                    tmpExp = Lines(runId).Body
                    If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tmpExp = "" Then
                        Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                    ElseIf IsNumeric(Trim(tmpExp)) = True Then
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 Then
                            For a = 0 To CInt(tftf) - 1
                                Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                            Next a
                        Else
                                    Err.Raise 209, , "Cannot move the Hero"
                        End If
                        
                    Else
                        tftf = FindNumber(tmpExp)
                        If Err.Number > 0 Then Exit Sub
                        If tftf > 0 Then
                            For a = 0 To CInt(tftf) - 1
                                Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                            Next a
                        Else
                                    Err.Raise 209, , "Cannot move the Hero"
                        End If
                        
                    End If
            Case ""
                ce.Exp Lines(runId).strText
            Case this_Code_If
                tmpExp = Lines(runId).Body
                If Lines(runId).if_endif_id <> 0 Then
                    '第一种情况 if logic then
                    If Right(tmpExp, 4) <> this_Code_Then Then
                        Err.Raise 209, , "If needs Then"
                        Exit Sub
                    End If
                    tmpExp = Trim(Left(tmpExp, Len(tmpExp) - 4))
                    ct = ce.logic(tmpExp)
                    If Err.Number > 0 Then
                        Exit Sub
                    End If
                    If ct = True Then
                        'do nothing
                    Else
                        If Lines(runId).if_else_id = 0 Then
                            runId = Lines(runId).if_endif_id
                        Else
                            runId = Lines(runId).if_else_id
                        End If
                    End If
                End If
            Case this_Code_Else
                runId = Lines(Lines(runId).else_if_id).if_endif_id
            Case this_Code_EndIf
                'do nothing
            Case this_Code_For
                '两种情况
                'for var = exp1 to exp2
                'for var = exp1 to exp2 step exp3
                tmpExp = Lines(runId).Body
                var1 = Trim(Left(tmpExp, InstrPlus(tmpExp, "=") - 1))
                Exp(0) = Trim(Mid(tmpExp, InstrPlus(tmpExp, "="), Len(tmpExp)))
                Exp(1) = Trim(Mid(Exp(0), InstrPlus(Exp(0), " " & this_Code_To & " ") + 3, Len(Exp(0))))
                Exp(0) = Trim(Left(Exp(0), InstrPlus(Exp(0), " " & this_Code_To & " ") - 1))
                Exp(0) = Trim(Mid(Exp(0), InStr(Exp(0), "=") + 1, Len(Exp(0))))
                With Lines(runId)
                    If InstrPlus(Exp(1), " " & this_Code_Step & " ") = 0 Then
                        '第1种情况
                        .for_step = 1
                    Else
                        '第2种情况
                        Exp(2) = Trim(Mid(Exp(1), InstrPlus(Exp(1), " " & this_Code_Step & " ") + 5, Len(Exp(1))))
                        Exp(1) = Trim(Left(Exp(1), InstrPlus(Exp(1), " " & this_Code_Step & " ") - 1))
                        .for_step = ce.Exp(Exp(2))
                    End If
                    .for_min = ce.Exp(Exp(0))
                    .for_max = ce.Exp(Exp(1))
                    .for_var_name = var1
                    If .for_step = 0 Then
                        Err.Raise 10, , "The Step cannot be 0"
                        Exit Sub
                    ElseIf .for_step > 0 Then
                        If .for_min > .for_max Then
                            runId = .for_next_id    '不满足for条件
                        Else
                            '满足for条件
                            globalCV.setVar .for_var_name, .for_min
                        End If
                    Else
                        '递减循环
                        If .for_min >= .for_max Then
                            globalCV.setVar .for_var_name, .for_min
                        Else
                            runId = .for_next_id
                        End If
                    End If
                End With
            Case this_Code_Next
                varn = Lines(Lines(runId).next_for_id).for_var_name
                ivar = ce.Exp(varn & "=" & varn & "+(" & Lines(Lines(runId).next_for_id).for_step & ")")
                If (ivar - Lines(Lines(runId).next_for_id).for_max) / Lines(Lines(runId).next_for_id).for_step > 0 Then
                    '同号，说明不满足循环条件了
                    'do nothing
                Else
                    '跳到for
                    runId = Lines(runId).next_for_id
                End If
            Case this_Code_ExitFor
                runId = Lines(Lines(runId).exitfor_for_id).for_next_id
            Case this_Code_Print
                tmpExp = Lines(runId).Body
                Do While True
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        bd = Mid(tmpExp, FindBorder(tmpExp), 1)
                        If bd = "," Then
                            bd = Chr(9)
                        Else
                            bd = ""
                        End If
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                            '输出到控制台
                            s = Mid(T1, 2, Len(T1) - 2) & bd
                        Else
                            s = ce.Exp(T1) & bd
                        End If
                        this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
'                        Debug.Print this_Graphic.SpriteFont(0).G_String
                        this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                        this_Graphic.SpriteFont(0).G_WordWraped = False
                        Call Event_Get(this_Graphic, 1)
   
                    Else
                        Exit Do
                    End If
                Loop
                
                If tmpExp <> "" Then
                    If Left(tmpExp, 1) = """" And Right(tmpExp, 1) = """" Then
                        '输出到控制台
                         s = Mid(tmpExp, 2, Len(tmpExp) - 2) & vbCrLf
                    Else
                         s = ce.Exp(tmpExp) & vbCrLf
                    End If
                    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                    this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                    this_Graphic.SpriteFont(0).G_WordWraped = False
                    Call Event_Get(this_Graphic, 1)

                End If
            Case this_Code_Input
                tmpExp = Lines(runId).Body
                If Right(tmpExp, 1) = "," Or Right(tmpExp, 1) = ";" Then tmpExp = Left(tmpExp, Len(tmpExp) - 1)
                Do While True
                    'Exit Do
                    tftf = FindBorder(tmpExp)
                    If Err.Number > 0 Then Exit Sub
                    If tftf > 0 Then
                        bd = Mid(tmpExp, FindBorder(tmpExp), 1)
                        If bd = "," Then
                            bd = Chr(9)
                        Else
                            bd = ""
                        End If
                        T1 = Trim(Mid(tmpExp, 1, FindBorder(tmpExp) - 1))
                        tmpExp = Trim(Mid(tmpExp, FindBorder(tmpExp) + 1, Len(tmpExp)))
                        If Left(T1, 1) = """" And Right(T1, 1) = """" Then
                            '输出到控制台
                            s = Mid(T1, 2, Len(T1) - 2) & bd
                        Else
                            s = tmpExp
                        End If
                        this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                        this_Graphic.SpriteFont(0).G_MaxLineLen = 300
                        this_Graphic.SpriteFont(0).G_WordWraped = False
                        DoEvents
                        Call Event_Get(this_Graphic, 1)
                        DoEvents
                        n = ""
                        While n = ""
                            DoEvents
                            n = InputBox(this_Graphic.SpriteFont(0).G_String, "Input")
                        Wend
                          
                        globalCV.setVar tmpExp, n ' fMainForm.Txt_Input.Text
                    Else
                        Exit Do
                    End If
                Loop
                
                If tmpExp <> "" Then
                    If Left(tmpExp, 1) = """" And Right(tmpExp, 1) = """" Then
                        '输出到控制台
                        s = Mid(tmpExp, 2, Len(tmpExp) - 2) & vbCrLf
                    Else
                        s = tmpExp
                    End If
                    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                    this_Graphic.SpriteFont(0).G_MaxLineLen = 300
                    this_Graphic.SpriteFont(0).G_WordWraped = False
                    DoEvents
                    Call Event_Get(this_Graphic, 1)
                    DoEvents
                    n = ""
                    While n = ""
                        DoEvents
                        n = InputBox(this_Graphic.SpriteFont(0).G_String, "Input")
                    Wend
                    
                    globalCV.setVar tmpExp, n ' fMainForm.Txt_Input.Text
                End If
            Case this_Code_End
                Exit Do
            Case this_Code_Goto, this_Code_GoSub
                tmpExp = Lines(runId).Body
                If Not IsNumeric(Trim(tmpExp)) Then
                    Err.Raise 50, , "Goto statement can only be followed by numbers "
                    Exit Sub
                End If
                If CInt(tmpExp) = 0 Then
                    Err.Raise 52, , "The line number cannot be found. "
                    Exit Sub
                End If
                cline = 0
                For i = 1 To UBound(Lines)
                    If Lines(i).vLine = CInt(tmpExp) Then
                        cline = i
                        Exit For
                    End If
                Next
                If cline = 0 Then
                    Err.Raise 51, , "The line number cannot be found. "
                    Exit Sub
                End If
                sksub.Push runId
                runId = cline - 1
            Case this_Code_Return
                If sksub.Count > 0 Then
                    runId = sksub.Pop
                End If
            Case this_Code_DoWhile
                tmpExp = Lines(runId).Body
                If ce.logic(tmpExp) = True Then
                    'do nothing
                Else
                    runId = Lines(runId).dowhile_loop
                End If
            Case this_Code_Do
                'do nothing
            Case this_Code_LoopUntil
                tmpExp = Lines(runId).Body
                 If ce.logic(tmpExp) = True Then
                    runId = Lines(runId).loopuntile_do_id
                Else
                   'do nothing
                End If
            Case this_Code_ExitDo
                If Lines(runId).exitdo_do <> "" Then
                    runId = Lines(Lines(runId).exitdo_do).do_loopuntil_id
                Else
                    runId = Lines(Lines(runId).exitdo_dowhile).dowhile_loop
                End If
            Case this_Code_Loop
                runId = Lines(runId).loop_dowhile - 1
            Case this_Code_Stop
                this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & s
                this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                this_Graphic.SpriteFont(0).G_WordWraped = False
                Call Event_Get(this_Graphic, 1)
        End Select
        If Err.Number > 0 Then
            Exit Sub
        End If
        runId = runId + 1
        s = s & vbNewLine & s
    Loop
    
End Sub



'=============================================================
'Describe: Running scripting language
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function Run(this_SourceCode As String)
    '
    On Error Resume Next
    globalCV.Clear              '
    iline = 0                   '
    
    Call Code_a(this_SourceCode)
    Call Code_b(this_SourceCode)
    Call Code_c(this_SourceCode)
    Call Code_d(this_SourceCode)
    Call Code_e(this_SourceCode)

End Function

Private Sub Class_Initialize()
    ReDim Preserve Lines(0)
End Sub



