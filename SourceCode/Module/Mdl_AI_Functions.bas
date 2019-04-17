Attribute VB_Name = "Mdl_AI_Functions"
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about Chat Robot
'Author:  Chenlyu and Greg Cohen
'E-mail: plarn@foxmail.com
'**************************************************************************

'====================================================================Function description====================================================
'Project Bob - An Artificially Intelligent Computer Program
'Lead Programmer Greg Cohen
'Assistant Programmer Brandon Resnick
'Artwork David McLennon
'Interface assistance Tim Berwyn - Taylor
'Tea and Refreshments Steven Cohen
'Project Board Design Brandon Resnick
'====================================================================================================================================


Option Explicit

Public Type Keyword
    KeywordText As String
    KeywordNo As String
    KeywordFile As String
    KeywordOrigin As String
End Type

Public Type Response
    ResponseText As String      'The response being used
    ResponseSubject As String   'The subject of the current response
    ResponseAction As String    'Misc information about the sentence (i.e  sarcastic, rhetorical, threat)
    Question  As Boolean        'Is it a question for the user
    QuestionType As String      'Stores the type of question and any extra info about the question (i.e yes/no, opinion, time)
    SubjectChange As Boolean    'Does this response cause a subject change
End Type


Public Enum ResponseType
    Statement = 1
    Question = 2
    Either = 3
    Other = 4
End Enum


'=============================================================
'Describe:Replace the Subject in the Sentenc
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function ReversePronoun(sSentence As String) As String
    Dim Working As String
    
    Working = " " & LCase(sSentence)
    'I
    Working = Replace(Working, " i ", " you ")
    Working = Replace(Working, " i've ", " you've ")
    Working = Replace(Working, " i'm ", " you're ")
    'You
    Working = Replace(Working, " you ", " i ")
    Working = Replace(Working, " you've ", " i've ")
    Working = Replace(Working, " your ", " my ")
    'Me
    Working = Replace(Working, " me ", " you ")
    'Us
    Working = Replace(Working, " us ", "you")
        
    ReversePronoun = Trim(Working)
End Function


'=============================================================
'Describe:Replace the Subject in the Sentenc
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function ReversePronoun2(sSentence As String) As String
    Dim Working As String
    
    Working = " " & LCase(sSentence)
    'I
    Working = Replace(Working, " i ", " I ")
    Working = Replace(Working, " i've ", " I've ")
    Working = Replace(Working, " i'm ", " I'm ")
    Working = Replace(Working, " i'll ", " I'll ")
        
    Working = Replace(Working, "i've ", "I've ")
    Working = Replace(Working, "i'm ", "I'm ")
    Working = Replace(Working, "i'll ", "I'll ")
    
    Working = Replace(Working, " i've", " I've")
    Working = Replace(Working, " i'm", " I'm")
    Working = Replace(Working, " i'll", " I'll")
    
    Working = Replace(Working, " not not ", " ")
    
    Working = Trim(Working)
    If Len(Working) > 1 Then Working = UCase(Left(Working, 1)) & Right(Working, Len(Working) - 1)
    
    Username = Trim(Username)
    If Len(Username) > 1 Then Username = UCase(Left(Username, 1)) & Right(Username, Len(Username) - 1)
    
    Working = Replace(Working, "|" & LCase(Username) & "|", Username)
    
    ReversePronoun2 = " " & Trim(Working)
End Function


'=============================================================
'Describe:Keyword Search
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function MUD_KeywordSearch(sKeywordFile As String, sSentence As String, sFiles() As tMUD_Files, Optional bSearchExclude As Boolean = False) As tMUD_Keyword

Dim i, j As Integer

    MUD_KeywordSearch.KeywordOrigin = sSentence
    For i = 0 To UBound(sFiles) - 1
        If LCase(Trim(sKeywordFile)) = LCase(Trim(sFiles(i).Name)) Then
            With sFiles(i)
                For j = 1 To UBound(.Keyword)
                    If InStr(1, LCase(sSentence), LCase(.Keyword(j).KeywordText)) <> 0 Then
                        If bSearchExclude = False Then
                            
                            MUD_KeywordSearch.KeywordText = LCase(.Keyword(j).KeywordText)
                            MUD_KeywordSearch.KeywordNo = LCase(.Keyword(j).KeywordNo)
                            MUD_KeywordSearch.KeywordFile = LCase(.Keyword(j).KeywordFile)
                            Exit For
                        Else
                            If LCase(.Keyword(j).KeywordText) = LastKeyword.KeywordText Then
                                'Ignore keyword
                            Else
                                MUD_KeywordSearch.KeywordText = LCase(.Keyword(j).KeywordText)
                                MUD_KeywordSearch.KeywordNo = LCase(.Keyword(j).KeywordNo)
                                MUD_KeywordSearch.KeywordFile = LCase(.Keyword(j).KeywordFile)
                                Exit For
                            End If
                        End If
                        
                    End If
                Next j
            End With
            Exit For
        Else
            Debug.Print ""
        End If
    Next i
    
End Function

'=============================================================
'Describe:Get the corresponding random sentences according to the keywords
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function MUD_GetRandomReply(sKeyword As tMUD_Keyword, sFiles() As tMUD_Files, Optional ReplyTypes As tMUD_ResponseType = 3) As tMUD_Response

    Dim AnswerCollection As New Collection
    Dim a, b As Integer
    Dim CurrentString As String
    Dim CurrentSubject As String
    
    Dim i, j As Integer
    For i = 0 To UBound(sFiles) - 1
        If LCase(Trim(sKeyword.KeywordFile)) = LCase(Trim(sFiles(i).Name)) Then
            With sFiles(i)
                For j = 1 To UBound(.Response)
                    If sKeyword.KeywordNo = .Response(j).ResponseNo Then

                        If .Response(j).ResponseType = Statement Then
                            If ReplyTypes = Statement Or ReplyTypes = Either Then AnswerCollection.Add "# " & .Response(j).ResponseText
                        End If
                        
                        If .Response(j).ResponseType = Question Then
                            If ReplyTypes = Question Or ReplyTypes = Either Then AnswerCollection.Add "$ " & .Response(j).ResponseText
                        End If
                        
                        If .Response(j).ResponseType = ExtraSearch Then
                            ExtraSearchFile = Trim(LCase(.Response(j).ResponseText))
                        End If
                    
                        CurrentSubject = Trim(LCase(.Response(j).ResponseSubject))
                    End If
                Next j
                MUD_GetRandomReply.ResponseSubject = CurrentSubject
                If AnswerCollection.Count > 0 Then
                    b = AnswerCollection.Count - 1
                    Randomize Timer
                    a = Int((b - 0 + 1) * Rnd + 1)
                    CurrentString = AnswerCollection.Item(a)
                
                    CurrentString = LCase(CurrentString)
                    If Left(CurrentString, 1) = "$" Then MUD_GetRandomReply.Question = True
                    CurrentString = Right(CurrentString, Len(CurrentString) - 1)
                    CurrentString = MUD_ReplaceCodeWords(CurrentString, sKeyword)
                
                    If InStr(1, CurrentString, "{") <> 0 Then
                        MUD_GetRandomReply.ResponseAction = Right(CurrentString, Len(CurrentString) - InStr(1, CurrentString, "{") + 1)
                        CurrentString = Trim(Left(CurrentString, InStr(1, CurrentString, "{") - 1))
                    End If
            
                    MUD_GetRandomReply.ResponseText = ReversePronoun2(CurrentString)
                End If
            End With
            Exit For
        Else
            Debug.Print ""
        End If
    Next i
    

    
 

End Function


'=============================================================
'Describe:Replace the corresponding keywords
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function MUD_ReplaceCodeWords(sSentence As String, sKeywordInfo As tMUD_Keyword) As String
On Error Resume Next
    Dim z As String
    Dim y() As String
    Dim x As Integer
    Dim W As String
    Username = Trim(Username)
    If Len(Username) > 1 Then Username = UCase(Left(Username, 1)) & Right(Username, Len(Username) - 1)
    z = Replace(sSentence, "[name]", "|" & Username & "|")
    z = Replace(z, "[keyword]", ReversePronoun(LCase(sKeywordInfo.KeywordText)))
    If InStr(1, z, "[next]") Then
        x = InStr(1, LCase(sKeywordInfo.KeywordOrigin), LCase(sKeywordInfo.KeywordText))
        W = Trim(Right(LCase(sKeywordInfo.KeywordOrigin), Len(sKeywordInfo.KeywordOrigin) - (x + Len(sKeywordInfo.KeywordText))))
        W = W & " "
        W = Replace(W, "?", "")
        W = Replace(W, "!", "")
        x = InStr(1, W, " ")
        W = Left(W, x)
        W = ReversePronoun(W)
        z = Replace(z, "[next]", W)
    End If
    If InStr(1, z, "[following]") Then
        x = InStr(1, LCase(sKeywordInfo.KeywordOrigin), LCase(sKeywordInfo.KeywordText))
        W = Trim(Right(LCase(sKeywordInfo.KeywordOrigin), Len(sKeywordInfo.KeywordOrigin) - (x + Len(sKeywordInfo.KeywordText))))
        '-------------------------------------------------------------------------------------------------------------------------
        W = Replace(W, "?", "")
        W = Replace(W, "!", "")
        W = ReversePronoun(W)
        z = Replace(z, "[following]", W)
    End If
    MUD_ReplaceCodeWords = z
End Function


'=============================================================
'Describe:Show related sentences according to key words
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function MUD_GetKeywordReply(sExtraSearchFile As String, sSentence As String, sFiles() As tMUD_Files) As tMUD_SimpleResponse
    Dim sResponse As tMUD_Response
    Dim sKeyword As tMUD_Keyword
    Dim Cont As Boolean
    
    Cont = True
    If sExtraSearchFile <> "" Then
        sKeyword = MUD_KeywordSearch(sExtraSearchFile, sSentence, sFiles())
        If sKeyword.KeywordNo <> "" Then
            Cont = False
        End If

        sExtraSearchFile = ""

    End If
    
    If Cont = True Then
        sKeyword = MUD_KeywordSearch("Keywords", sSentence & " ", sFiles())
        If sKeyword.KeywordText = LastKeyword.KeywordText Then
            sKeyword = MUD_KeywordSearch("Keywords", sSentence & " ", sFiles(), True)
        End If
    End If
    
    
    If sKeyword.KeywordText <> "" Then
        sResponse = MUD_GetRandomReply(sKeyword, sFiles())
    Else
        sKeyword.KeywordFile = "Preset3"
        sKeyword.KeywordNo = "0001"
        sResponse = MUD_GetRandomReply(sKeyword, sFiles())
    End If
    
    'Remove Not Not error
    
    
    
    MUD_GetKeywordReply.sReply = sResponse.ResponseText
    MUD_GetKeywordReply.sAction = sResponse.ResponseAction
    MUD_GetKeywordReply.sQuestion = sResponse.Question
    

    
    LastKeyword = sKeyword
    LastResponse = sResponse
End Function

'=============================================================
'Describe:AI asks the appropriate questions.
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function AIRobotAsk(myFunction As String, sFiles() As tMUD_Files) As String
    Dim sKey As tMUD_Keyword

    ReturnFunction = myFunction
    Select Case myFunction
        Case "GetUsername"
            sKey.KeywordFile = "Preset2"
            sKey.KeywordNo = "0002"
        Case "but"
            sKey.KeywordFile = "Preset1"
            sKey.KeywordNo = "0010"
            sKey.KeywordText = "But"
            
        Case Else
 
    End Select
        
'    AIRobotAsk = GetRandomReply(sKey).ResponseText
    AIRobotAsk = MUD_GetRandomReply(sKey, sFiles).ResponseText
    
End Function

Public Function AIRobotSpeak(sKey As tMUD_Keyword, sFiles() As tMUD_Files) As String

    AIRobotSpeak = MUD_GetRandomReply(sKey, sFiles).ResponseText
    
End Function


'=============================================================
'Describe:Robots answer questions
'Author:  Greg Cohen & Chen Lyu
'Parameter:
'=============================================================
Public Function AIRobotResponse(myTest As String, sFiles() As tMUD_Files) As String
    Dim sKey As tMUD_Keyword
    Dim sResponse As tMUD_Response
    Dim myFunction As String
    Dim sReply As tMUD_SimpleResponse
    myFunction = ReturnFunction
    Select Case myFunction
        Case "GetUsername"
            If myTest = "" Or myTest = " " Then myTest = "{SILENCE}"
'            sKey = KeywordSearchFile(UCase(myTest), App.Path & "\data\preset1.txt")
            sKey = MUD_KeywordSearch("presets", UCase(myTest), sFiles)
            

            If sKey.KeywordText = "" Then
                If InStr(1, Trim(myTest), " ") > 0 Then
                    sKey.KeywordFile = "Preset1"
                    sKey.KeywordNo = "0013"
                    sKey.KeywordText = myTest
                    sKey.KeywordOrigin = myTest
                Else
                    Username = Trim(myTest)
                    sKey.KeywordFile = "Preset2"
                    sKey.KeywordNo = "0001"
                    sKey.KeywordText = myTest
                    sKey.KeywordOrigin = myTest
                    ReturnFunction = ""
                End If
            End If
            
'            sResponse = GetRandomReply(sKey)
            sResponse = MUD_GetRandomReply(sKey, sFiles)
            AIRobotResponse = sResponse.ResponseText
        Case "help"
            sKey.KeywordFile = "Preset1"
            sKey.KeywordNo = "0015"
            sKey.KeywordText = "help"
            sKey.KeywordOrigin = myTest
            ReturnFunction = ""
            sResponse = MUD_GetRandomReply(sKey, sFiles)
            AIRobotResponse = sResponse.ResponseText
        Case Else
'             sReply = GetKeywordReply(myTest)
             sReply = MUD_GetKeywordReply(ExtraSearchFile, myTest, sFiles)
             AIRobotResponse = sReply.sReply
    End Select
        
    
End Function
