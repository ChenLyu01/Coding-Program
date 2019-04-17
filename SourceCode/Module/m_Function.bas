Attribute VB_Name = "m_Function"
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about Game Sub Main
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************



'====================================================================Function description====================================================

' Create a game editor that you can use free source code

'====================================================================================================================================


Option Explicit


Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const m_KeyWords As String = "stop,case else,end select,exit for,exit do,elseif ,do while ,end if,select case ,loop until ,print ,input ,if ,for ,next,gosub ,return,while ,wend,end,else,loop,case ,end,do,goto ,gosub ,game.autorun ,message.show ,message.clear ,hero.create ,hero.moveleft ,hero.moveright ,hero.movedown ,hero.moveup ,tile.style ,pathway.style ,pathway.draw ,pathway.clear ,block.show ,block.hide ,hero.show ,hero.hide ,pathway.show ,pathway.hide ,object.show ,object.hide ,event.show ,event.hide ,effect.show ,effect.hide ,event.destroy ,event.create ,object.destroy ,object.create ,object.clear ,block.destroy ,block.create ,effect.clear ,effect.destroy ,effect.create ,effect.disable ,effect.enable ,file.open ,file.ai ,hero.steps ,hero.movexy ,pathway.destroy ,pathway.block ,hero.say ,event.clear ,block.clear ,hero.open ,hero.collect ,game.task ,message.close ,hero.faceup ,hero.facedown ,hero.faceleft ,hero.faceright "
Public Const m_Keyword As String = "mod,xor,abs,sqr,stop,if,else,elseif,for,next,to,end,and,or,not,exit,print,step,while,wend,select,case,input,do,loop,until,goto,gosub,return,and,or,not,then,game.autorun,hero.create,message.show,message.clear,hero.moveleft,hero.moveright,hero.movedown,hero.moveup,tile.style,pathway.style,pathway.draw,pathway.clear,block.show,block.hide,hero.show,hero.hide,pathway.show,pathway.hide,object.show,object.hide,event.show,event.hide,effect.show,effect.hide,event.destroy,event.create,object.destroy,object.create,object.clear,block.destroy,block.create,effect.clear,effect.destroy,effect.create,effect.disable,effect.enable,file.open,file.ai,hero.steps,hero.movexy,pathway.destroy,pathway.block,hero.say,event.clear,block.clear,hero.open,hero.collect,game.task,message.close,hero.faceup,hero.facedown,hero.faceleft,hero.faceright"

Public ErrLine As Long      '³ö´íÎ»ÖÃ
Public globalVars As Collection
Public globalPause As New Collection
 


Public flagChange As Boolean 'Has the text been edited?
Public globalln As Long
'Recording window coordinates of frmVars
Public frmVars_top As Long
Public frmVars_left As Long

Public globalCV As New clsVar

Public fMainForm As frmMain

Public Const EM_SCROLLCARET = &HB7



'=============================================================
'Describe:Program startup function
'Author:  Chenlyu
'Parameter:
'=============================================================
Sub Main()
    
    this_FilePath.Graphics = App.Path & "\Graphics\Skins\"
    this_FilePath.CourseMap = App.Path & "\Map\"
    this_FilePath.Code = App.Path & "\Data\"
    this_FilePath.Story = App.Path & "\Story\"
    
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Filelist.Path = this_FilePath.Code
    Call Game_Initialize(fMainForm.PicMain.hDC, this_FilePath)

    this_Graphic.AILoaded = False
     
    fMainForm.Show
    '
    Call AI_Main
    If Command = "" Then
        '        this_FilePath.CodeName = "Plarn"
        '        Call Files_Data_Map_Save(this_Graphic, this_FilePath, "Map")
        '        Call Files_Data_Map_Save(this_Graphic, this_FilePath, "Player")
        '         Call Files_Data_Map_Save(this_Graphic, this_FilePath, "Story")
        '        Call Files_Data_Map_Read(this_Graphic, this_FilePath, "Map")
        '        Call Files_Data_Map_Read(this_Graphic, this_FilePath, "Player")
        '        Call Files_Data_Map_Read(this_Graphic, this_FilePath, "Story")
         
        fMainForm.Timer3.Enabled = True
        Call CodeFileLoad(this_Graphic, "plarn.pl")
    Else
        fMainForm.txtSource.Text = FileRead(Trim(Command))
        this_FilePath.CodeName = Command
        flagChange = False
        fMainForm.Timer3.Enabled = True
    End If
    
'    MySoftwarePassWord = Replace(GetSetting("Plarn", "Reg", "DiskNum", 1), "-", "")
'
'    If MySoftwarePassWord = "" Or MySoftwarePassWord = "1" Then MySoftwarePassWord = "Plarn"
'    User_Code = Replace(GetSetting("Plarn", "Reg", "RegPSD", 1), "-", "")
'    SoftwareRegPSD = Replace(GetSetting("Plarn", "Reg", "PC", 1), "-", "")
    
End Sub


'=============================================================
'Describe:Create a batch of new markers
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub BanAdd(ByVal Item As Integer)
    Dim i As Integer
    For i = 1 To globalPause.Count
        If globalPause(i) = Item Then Exit Sub
    Next
    globalPause.Add Item
End Sub

'=============================================================
'Describe:Delete a marker
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub BanDel(ByVal Item As Integer)
    Dim i As Integer
    For i = 1 To globalPause.Count
        If globalPause(i) = Item Then
            globalPause.Remove i
            Exit For
        End If
    Next
End Sub

'=============================================================
'Describe:Clear all markers
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub BanClear()
    While globalPause.Count > 0
        globalPause.Remove 1
    Wend
End Sub

'=============================================================
'Describe:Check for presence of markers
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function BanHave(ByVal Item As Integer) As Boolean
    Dim i As Integer
    For i = 1 To globalPause.Count
        If globalPause(i) = Item Then
            BanHave = True
            Exit Function
        End If
    Next
    BanHave = False
End Function

Public Function inArray(ss() As String, ByVal s As String)
    Dim i
    For Each i In ss
        If i = s Then
            inArray = True
            Exit Function
        End If
    Next
    inArray = False
End Function

'=============================================================
'Describe:Functions that read strings as needed
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function InstrPlus(ByVal s As String, ByVal strin As String)
    '
    On Error Resume Next
    Dim ts As String
    Dim i As Integer
    Dim c As String
    Dim Count As Integer
    If s = "" Or strin = "" Then
        InstrPlus = 0
        Exit Function
    End If
    strin = LCase(strin)
    Count = 1
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = """" Then
            If InStr(ts, """") = 0 Then
                ts = ts & """"
            Else
                ts = Mid(ts, 1, InStrRev(ts, """") - 1)
            End If
        Else
            If ts = "" Then
                If c = Mid(strin, Count, 1) Then
                    Count = Count + 1
                    If Count = Len(strin) + 1 Then
                        InstrPlus = i - Len(strin) + 1
                        Exit Function
                    End If
                Else
                    Count = 1
                End If
            End If
        End If
    Next
    If ts <> "" Then
        Err.Raise 202, , "Expression " & s & "syntax error"
        Exit Function
    End If
    InstrPlus = 0
End Function

'=============================================================
'Describe:Functions that read strings as needed
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function ReplacePlus(ByVal s As String, ByVal ss As String, ByVal sd As String)
    If InStr(sd, ss) > 0 Then
        Err.Raise 333, , "Cannot be replaced"
        Exit Function
    End If
    Do While InstrPlus(s, ss) > 0
        s = Left(s, InstrPlus(s, ss) - 1) & sd & Mid(s, InstrPlus(s, ss) + Len(ss), Len(s))
    Loop
    ReplacePlus = s
End Function

'=============================================================
'Describe:Functions that read strings as needed
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function FindBorder(ByVal s As String)
    '
    Dim ts As String
    Dim i As Integer
    Dim c As String
    If s = "" Then
        Exit Function
    End If
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = "(" Then
            ts = ts + c
        ElseIf c = ")" Then
            If InStr(ts, "(") = 0 Then
                Err.Raise 201, , "Expression error, missing £¨"
                Exit Function
            End If
            ts = Mid(ts, 1, InStrRev(ts, "(") - 1)
        ElseIf c = """" Then
            If InStr(ts, """") = 0 Then
                ts = ts & """"
            Else
                ts = Mid(ts, 1, InStrRev(ts, """") - 1)
            End If
        Else
            If ts = "" And (c = "," Or c = ";") Then
                FindBorder = i
                Exit Function
            End If
        End If
    Next
    If ts <> "" Then
        Err.Raise 202, , "Expression " & s & " error "
        Exit Function
    End If
    FindBorder = 0
End Function

'=============================================================
'Describe:Functions that read strings as needed
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function FindNumber(ByVal s As String) As String

    '
    Dim ts As String
    Dim i, a As Integer
    Dim c As String
    Dim t() As String
    If s = "" Then
        Exit Function
    End If
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = "(" Then
            ts = ts + c
        ElseIf c = ")" Then
            If InStr(ts, "(") = 0 Then
                Err.Raise 201, , "Expression error, missing £¨"
                Exit Function
            End If
            ts = Mid(ts, InStrRev(ts, "(") + 1, Len(ts))
        ElseIf c = """" Then
            FindNumber = 0
            Err.Raise 201, , "Expression error, needs number"
            Exit Function
        Else
            ts = ts & c
        End If
    Next
    

    
        If IsNumeric(Trim(ts)) = True Then
            
            If InStr(Trim(ts), ",") > 0 Then
                FindNumber = CStr(ts)
                Exit Function
            Else
                FindNumber = CLng(Trim(ts))
                Exit Function
            End If
                
        Else
            If InStr(Trim(ts), ",") > 0 Then
                t = Split(ts, ",")
                If UBound(t) > 0 Then
                    ts = ""
                    For a = 0 To UBound(t)
                        If IsNumeric(Trim(t(a))) = True Then
                        
                        Else
                            For i = 1 To globalCV.VarNames.Count
                                'Debug.Print globalCV.VarNames(i) & "====" & globalCV.getVar(globalCV.VarNames(i))
                                If Trim(t(a)) = globalCV.VarNames(i) Then
                                    t(a) = globalCV.getVar(globalCV.VarNames(i))
                                    'Debug.Print t(a)
                                    If IsNumeric(Trim(t(a))) = True Then
                                        Exit For
                                    Else
                                        Err.Raise 201, , "Expression error, needs number"
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                        ts = ts & ", " & t(a)
                    Next a
                
                End If
                FindNumber = Right(ts, Len(ts) - 1)
                Exit Function
            Else
                For i = 1 To globalCV.VarNames.Count
                    If Trim(ts) = globalCV.VarNames(i) Then
                        ts = globalCV.getVar(globalCV.VarNames(i))
                        Exit For
                    End If
                Next
                
                If IsNumeric(Trim(ts)) = True Then
                    If InStr(Trim(ts), ",") > 0 Then
                        FindNumber = CStr(ts)
                        Exit Function
                    Else
                        FindNumber = CLng(Trim(ts))
                        Exit Function
                    End If
                Else
                    If InStr(Trim(ts), ",") > 0 Then
                        FindNumber = CStr(ts)
                        Exit Function
                    Else
                        FindNumber = 0
                        Err.Raise 201, , "Expression error, needs number"
                        Exit Function
                    End If
                End If
            End If
        End If

End Function

