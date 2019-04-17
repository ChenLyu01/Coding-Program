Attribute VB_Name = "Mdl_Functions"
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and James L. Dean
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Adventures in 4 Dimensions Copyright James L. Dean csvcjld@nomvs.lsumc.edu This application may be distributed or used without payment to James L. Dean.

'====================================================================================================================================


Option Explicit
Dim Mud_Answer(2) As String
Dim Mud_KeyWordsSentence(2) As String
Dim Mud_KeyWordsSentencees As String



 

'=============================================================
'Describe:On the Functions of English Learning Lexicon
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub WordsControl()
    Dim i As Long
    Dim s As String
    Dim j As Integer
    For i = 0 To UBound(MUD_Word_Words) - 1
        oDict.Add MUD_Word_Words(i).Words, i
    Next i
    
    
    For i = 0 To UBound(MUD_IrregularVerb) - 1
        With MUD_IrregularVerb(i)
            If .WordsUID > 0 Then
                For j = 0 To UBound(.PastParticipleWord) - 1
                    s = Replace(.PastParticipleWord(j), Chr(0), "")
                    If s <> "" Then
                        If oDict.Exists(s) Then
                            oDict.Item(s) = .WordsUID
                        Else
                            oDict.Add (s), .WordsUID
                        End If
                    End If
                Next j
                For j = 0 To UBound(.PastTenseWord) - 1
                    s = Replace(.PastParticipleWord(j), Chr(0), "")
                    If s <> "" Then
                        If oDict.Exists(s) Then
                            oDict.Item(s) = .WordsUID
                        Else
                            oDict.Add (s), .WordsUID
                        End If
                    End If
                Next j
            End If
'            Debug.Print MUD_IrregularVerb(i).WordsUID
'            Debug.Print MUD_Word_Words(MUD_IrregularVerb(i).WordsUID).Words
'            Debug.Print MUD_IrregularVerb(i).InfinitiveWord
        End With
    Next i
    
    For i = 0 To UBound(MUD_Word_Addition) - 1
        With MUD_Word_Addition(i)
            s = Trim(Replace(.Addition, Chr(0), ""))
                If s <> "" Then
                    If oDict.Exists(s) Then
                        oDict.Item(s) = .ID
                    Else
                        oDict.Add (s), .ID
                    End If
                End If
        
'            Debug.Print MUD_Word_Addition(i).Addition
        End With
    Next i
End Sub

'=============================================================
'Describe:Read English Corpus
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Function Files_Files_Read(myUserFiles As String) As String
    Dim Filenr As Integer
    Dim s As String
    Dim LoadBytes() As Byte
    Dim i As Long
 
    s = App.Path & "\" & myUserFiles & ".txt"
    If PathFileExists(s) = 0 Then
        Files_Files_Read = ""
        Exit Function
    End If
        
    Filenr = FreeFile
    Open s For Binary As #Filenr
        ReDim LoadBytes(1 To LOF(Filenr)) As Byte
        Get #Filenr, , LoadBytes
        Files_Files_Read = StrConv(LoadBytes, vbUnicode)
    Close #Filenr
End Function

'=============================================================
'Describe:Save user data
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_MUDUser_Save(myUserName As String, myUser As tMUD_User)

    Dim Filenr As Integer
    Dim s As String
    s = App.Path & "\Data\AI\" & UCase(myUserName) & ".Usr"
    Filenr = FreeFile
    
    Dim i As Integer
    MUD_User.HisScore(9) = Trim(Left(MUD_ValROM(2), InStr(1, MUD_ValROM(2), "/") - 1))
    For i = 0 To 9
        MUD_User.level(0 + i) = MUD_ValROM(8 + i)
    Next i
    
    Open s For Binary Access Write As #Filenr
        'Put #Filenr, , UBound(myUser)
        Put #Filenr, , myUser
    Close #Filenr
End Sub

'=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Practice_Read(myPractice() As tMUD_Practice)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Practice" & ".Data"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myPractice(i)
        Get #Filenr, , myPractice
    Close #Filenr
End Sub

'=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_MUDUser_Read(myUserName As String, myUser As tMUD_User)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\" & UCase(myUserName) & ".Usr"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        'Get #Filenr, , i
        'ReDim myUser(i)
        Get #Filenr, , myUser
    Close #Filenr
End Sub

'=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Factor_Read(myWord_Factor() As tMUD_Factor)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Word_Factor" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myWord_Factor(i)
        Get #Filenr, , myWord_Factor
    Close #Filenr
End Sub
'=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_FactorIndex_Read(myWord_FactorIndex() As tMUD_Factor_Index)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Word_FactorIndex" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myWord_FactorIndex(i)
        Get #Filenr, , myWord_FactorIndex
    Close #Filenr
End Sub
 
 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Correlation_Read(myWordsCorrelation() As tMUD_Correlation)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "WordsCorrelation" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myWordsCorrelation(i)
        Get #Filenr, , myWordsCorrelation
    Close #Filenr
End Sub

 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_CorrelationIndex_Read(myCorrelation_Index() As tMUD_Correlation_Index)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "WordsCorrelation_Index" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myCorrelation_Index(i)
        Get #Filenr, , myCorrelation_Index
    Close #Filenr
End Sub
 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_DataFrequency_Read(my_Word_Frequency() As tMUD_Word_Frequency)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Frequency" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_Word_Frequency(i)
        Get #Filenr, , my_Word_Frequency
    Close #Filenr
End Sub
 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_DataMember_Read(my_Word_Member() As String)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\MemonyHelp.Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_Word_Member(i)
        Get #Filenr, , my_Word_Member
    Close #Filenr
End Sub
 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_DataIrregularVerb_Read(my_IrregularVerb() As tMUD_IrregularVerb)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "IrregularVerb" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_IrregularVerb(i)
        Get #Filenr, , my_IrregularVerb
    Close #Filenr
End Sub
 '=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_DataAddition_Read(my_Word_Addition() As tMUD_Word_Addition)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Addition" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_Word_Addition(i)
        Get #Filenr, , my_Word_Addition
    Close #Filenr
End Sub
'=============================================================
'Describe:The Read-Write Operation of the Corpus on TOEFL Learning
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_DataWordsWords_Read(my_Words() As tMUD_Word_Words)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\User\" & "Words" & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_Words(i)
        Get #Filenr, , my_Words
    Close #Filenr
End Sub

'=============================================================
'Describe:Reading and Writing of Data in Four-Dimensional Space
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Treasure_Save(myTreasure() As tMUD_Treasure)

    Dim Filenr As Integer
    Dim s As String
    s = App.Path & "\Data\AI\" & "Treasure" & ".Data"
    Filenr = FreeFile
    
    Open s For Binary Access Write As #Filenr
        Put #Filenr, , UBound(myTreasure)
        Put #Filenr, , myTreasure
    Close #Filenr
End Sub

'=============================================================
'Describe:Reading and Writing of Data in Four-Dimensional Space
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Treasure_Read(myTreasure() As tMUD_Treasure)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\" & "Treasure" & ".Data"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myTreasure(i)
        Get #Filenr, , myTreasure
    Close #Filenr
End Sub

'=============================================================
'Describe:Reading and Writing of Data in Four-Dimensional Space
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Rooms_Save(myRooms() As tMUD_Rooms)

    Dim Filenr As Integer
    Dim s As String
    s = App.Path & "\Data\AI\" & "Descript" & ".Data"
    Filenr = FreeFile
    
    Open s For Binary Access Write As #Filenr
        Put #Filenr, , UBound(myRooms)
        Put #Filenr, , myRooms
    Close #Filenr
End Sub

'=============================================================
'Describe:Reading and Writing of Data in Four-Dimensional Space
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Rooms_Read(myRooms() As tMUD_Rooms)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\" & "Descript" & ".Data"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myRooms(i)
        Get #Filenr, , myRooms
    Close #Filenr
End Sub


'=============================================================
'Describe:Reading and Writing of Data in Four-Dimensional Space
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_my_Files_Read(MUD_Files() As tMUD_Files, FileName As String)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
    Dim j As Long
    Dim my_Files() As my_Files
    Dim myKeyword() As tMUD_Keyword
    Dim myResponse() As tMUD_Response
    
    
    s = App.Path & "\Data\AI\User\" & FileName & ".Dat"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim my_Files(i)
        Get #Filenr, , my_Files
    Close #Filenr
    
    ReDim MUD_Files(UBound(my_Files()))
    For i = 0 To UBound(my_Files()) - 1
        MUD_Files(i).Name = my_Files(i).Name
        MUD_Files(i).Type = my_Files(i).Type
        MUD_Files(i).Count = my_Files(i).Count
    Next i
    
    For i = 0 To UBound(my_Files) - 1
        With MUD_Files(i)
            s = App.Path & "\Data\AI\User\" & .Name & ".Dat"
            Filenr = FreeFile
            Open s For Binary Access Read As #Filenr
                If .Type = "K" Then
                    Get #Filenr, , j
                    ReDim myKeyword(j)
                    Get #Filenr, , myKeyword
                    Let .Keyword = myKeyword()
                ElseIf .Type = "T" Then
                    Get #Filenr, , j
                    ReDim myResponse(j)
                    Get #Filenr, , myResponse
                    Let .Response = myResponse()
                End If
                
            Close #Filenr
        End With
    Next i
End Sub



'=============================================================
'Describe:Relevant Functions for Loading and Reading Four-Dimensional Games
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub MUD_Info(myDialogueROM() As tDialogue)
'    ReDim MUD_DialogueROM(10)
'    MUD_DialogueROM(0).ID = 1
'    MUD_DialogueROM(0).Text(0) = "ZORK VII: The Husky's Matrix \n Copyright (c) 1981,1982... 2014-2019 by Chen Lyu. \n <lyu.c@husky.neu.edu> All rights NOT reserved.  \n ZORK VII, this copy ,you can send to anywhere to edit and play around with it. "
'    MUD_DialogueROM(0).Type = 1
'
'    MUD_DialogueROM(1).ID = 3
'    MUD_DialogueROM(1).Text(0) = "Now you are in the "
'    MUD_DialogueROM(1).Type = 2
'
'
'    MUD_DialogueROM(2).ID = 5
'    MUD_DialogueROM(2).Text(0) = " Dimensional World! Your talking guy is a pretty puppy AI. "
'    MUD_DialogueROM(2).Type = 1
'
'    Call Files_Data_Dialogue_Save(myDialogueROM)


    Call Files_Data_Dialogue_Read(myDialogueROM)
End Sub

'=============================================================
'Describe:Data Reading Operation on Chat Robot
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Dialogue_Save(myDialogueROM() As tDialogue)

    Dim Filenr As Integer
    Dim s As String
    s = App.Path & "\Data\AI\" & "Dialogue" & ".Data"
    Filenr = FreeFile
    
    Open s For Binary Access Write As #Filenr
        Put #Filenr, , UBound(myDialogueROM)
        Put #Filenr, , myDialogueROM
    Close #Filenr
End Sub

'=============================================================
'Describe:Data Reading Operation on Chat Robot
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Files_Data_Dialogue_Read(myDialogueROM() As tDialogue)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
 
    s = App.Path & "\Data\AI\" & "Dialogue" & ".Data"
    Filenr = FreeFile
    Open s For Binary Access Read As #Filenr
        Get #Filenr, , i
        ReDim myDialogueROM(i)
        Get #Filenr, , myDialogueROM
    Close #Filenr
    
End Sub

'=============================================================
'Describe:Data Reading Operations for Four-Dimensional Space Games
'Author:  James L. Dean &  Chen Lyu
'Parameter:
'=============================================================
Public Sub GameUpdate(myValROM() As String)
    Dim sUserType As String
    Dim t() As String
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim sss As String
    Dim strKey
    myValROM(1) = nMoves

    sss = ""
    If (Not bEuclidean) Then
        If nXCoordinate < 0 Then
            nYCoordinate = nWidth(1) - 1 - nYCoordinate
            nZCoordinate = nWidth(2) - 1 - nZCoordinate
            nTCoordinate = nWidth(3) - 1 - nTCoordinate
            nXCoordinate = nWidth(0) - 1
        Else

            If nXCoordinate >= nWidth(0) Then
                nYCoordinate = nWidth(1) - 1 - nYCoordinate
                nZCoordinate = nWidth(2) - 1 - nZCoordinate
                nTCoordinate = nWidth(3) - 1 - nTCoordinate
                nXCoordinate = 0
            End If
        End If

        If nYCoordinate < 0 Then
            nXCoordinate = nWidth(0) - 1 - nXCoordinate
            nZCoordinate = nWidth(2) - 1 - nZCoordinate
            nTCoordinate = nWidth(3) - 1 - nTCoordinate
            nYCoordinate = nWidth(1) - 1
        Else

            If nYCoordinate >= nWidth(1) Then
                nXCoordinate = nWidth(0) - 1 - nXCoordinate
                nZCoordinate = nWidth(2) - 1 - nZCoordinate
                nTCoordinate = nWidth(3) - 1 - nTCoordinate
                nYCoordinate = 0
            End If
        End If

        If nZCoordinate < 0 Then
            nXCoordinate = nWidth(0) - 1 - nXCoordinate
            nYCoordinate = nWidth(1) - 1 - nYCoordinate
            nTCoordinate = nWidth(3) - 1 - nTCoordinate
            nZCoordinate = nWidth(2) - 1
        Else

            If nZCoordinate >= nWidth(2) Then
                nXCoordinate = nWidth(0) - 1 - nXCoordinate
                nYCoordinate = nWidth(1) - 1 - nYCoordinate
                nTCoordinate = nWidth(3) - 1 - nTCoordinate
                nZCoordinate = 0
            End If
        End If

        If nTCoordinate < 0 Then
            nXCoordinate = nWidth(0) - 1 - nXCoordinate
            nYCoordinate = nWidth(1) - 1 - nYCoordinate
            nZCoordinate = nWidth(2) - 1 - nZCoordinate
            nTCoordinate = nWidth(3) - 1
        Else

            If nTCoordinate >= nWidth(3) Then
                nXCoordinate = nWidth(0) - 1 - nXCoordinate
                nYCoordinate = nWidth(1) - 1 - nYCoordinate
                nZCoordinate = nWidth(2) - 1 - nZCoordinate
                nTCoordinate = 0
            End If
        End If
    End If

    nRoom1 = nCell(nXCoordinate, nYCoordinate, nZCoordinate, nTCoordinate)

    Randomize
    If ((nRoom1 <> 0) And (strWayOut = "") And (Int(100# * Rnd) = 0)) Then
        nRoom2 = 0

        Do While nRoom2 <= 0
            Randomize
            nXCoordinate = Int(CDbl(nWidth(0)) * Rnd)
            Randomize
            nYCoordinate = Int(CDbl(nWidth(1)) * Rnd)
            Randomize
            nZCoordinate = Int(CDbl(nWidth(2)) * Rnd)
            Randomize
            nTCoordinate = Int(CDbl(nWidth(3)) * Rnd)
            nRoom2 = nCell(nXCoordinate, nYCoordinate, nZCoordinate, nTCoordinate)
        Loop

        If nRoom2 <> nRoom1 Then
            nRoom1 = nRoom2
            sss = sss & Trim("Yeowwww! A flock of bats grabs you,  flies you through the caverns,  and drops you.") & vbNewLine
        End If
    End If

    strWayOut = ""
    nTreasuresRecovered = 0
    nTreasure1 = 0
    bTreasureCarried = False

    Do While (nTreasure1 < nTreasures) And (Not bTreasureCarried)

        If MUD_Treasure(nTreasure1).nTreasureRoom < 0 Then
            bTreasureCarried = True
        Else
            nTreasure1 = nTreasure1 + 1
        End If

    Loop

    If bTreasureCarried Then
        Randomize
        If Int(CDbl(2 * nRooms) * Rnd) = 0 Then
            nRoom2 = 0

            Do While nRoom2 <= 0
                nDimension1 = 0

                Do While nDimension1 < nDimensions
                    Randomize
                    nCoordinate(nDimension1) = Int(CDbl(nWidth(nDimension1)) * Rnd)
                    nDimension1 = nDimension1 + 1
                Loop

                nRoom2 = nCell(nCoordinate(0), nCoordinate(1), nCoordinate(2), nCoordinate(3))

                If nRoom1 = nRoom2 Then
                    nRoom2 = -1
                End If

            Loop

            nTreasure1 = 0

            Do While nTreasure1 < nTreasures

                If MUD_Treasure(nTreasure1).nTreasureRoom < 0 Then
                    MUD_Treasure(nTreasure1).nTreasureRoom = nRoom2
                End If

                nTreasure1 = nTreasure1 + 1
            Loop

            bTreasureCarried = False
            sss = sss & Trim("A pirate jumps out of the shadows and takes your treasure.") & vbNewLine
            sss = sss & Trim("As he leaves,  he says,  'Arggh!  I'll hide me booty better this time.'") & vbNewLine
        End If
    End If

    nTreasure1 = 0
    nTreasure2 = 0
    strTreasures = ""
    myValROM(19) = " "
    nTreasuresCarried = 0

    Do While nTreasure1 < nTreasures

        If MUD_Treasure(nTreasure1).nTreasureRoom = 0 Then
            nTreasuresRecovered = nTreasuresRecovered + 1

            If nRoom1 = 0 Then
                strTreasures = strTreasures & "  There's " & MUD_Treasure(nTreasure1).strTreasure & " here. "
            End If

        Else

            If MUD_Treasure(nTreasure1).nTreasureRoom = nRoom1 Then
                strTreasures = strTreasures & "  There's " & MUD_Treasure(nTreasure1).strTreasure & " here. "

                If MUD_Treasure(nTreasure1).nGuardRoom = nRoom1 Then
                    strLine = Left(MUD_Treasure(nTreasure1).strGuard, 1)

                    If ((strLine = "a") Or (strLine = "e") Or (strLine = "i") Or (strLine = "o") Or (strLine = "u")) Then
                        strTreasures = strTreasures & "  It's guarded by an " & MUD_Treasure(nTreasure1).strGuard & "."
                    Else
                        strTreasures = strTreasures & "  It's guarded by a " & MUD_Treasure(nTreasure1).strGuard & "."
                    End If
                End If

            Else

                If MUD_Treasure(nTreasure1).nTreasureRoom = -1 Then
                    bTreasureCarried = True
                    nTreasuresCarried = nTreasuresCarried + 1
                    nTreasure2 = nTreasure2 + 1
                    myValROM(19) = myValROM(19) & MUD_Treasure(nTreasure1).strTreasure & ", "
                End If
            End If
        End If

        If MUD_Treasure(nTreasure1).nWeaponRoom = nRoom1 Then
            strLine = Left(MUD_Treasure(nTreasure1).strWeapon, 1)

            If ((strLine = "a") Or (strLine = "e") Or (strLine = "i") Or (strLine = "o") Or (strLine = "u")) Then
                strTreasures = strTreasures & " There's an " & MUD_Treasure(nTreasure1).strWeapon & " here."
            Else
                strTreasures = strTreasures & " There's a " & MUD_Treasure(nTreasure1).strWeapon & " here."
            End If

        Else

            If MUD_Treasure(nTreasure1).nWeaponRoom = -1 Then
                nTreasure2 = nTreasure2 + 1
                myValROM(19) = myValROM(19) & MUD_Treasure(nTreasure1).strWeapon & ", "
            End If
        End If

        nTreasure1 = nTreasure1 + 1
    Loop

    myValROM(4) = nTreasuresRecovered & "/" & nTreasures
 

    If (Not MUD_Rooms(nRoom1).bVisited) Then
        nVisited = nVisited + 1
        MUD_Rooms(nRoom1).bVisited = True
    End If

    myValROM(3) = nVisited & "/" & nRooms
 
    dblScore = 25# * CDbl(nVisited) / CDbl(nRooms) + 75# * CDbl(nTreasuresRecovered) / CDbl(nTreasures) + 45# * CDbl(nTreasuresCarried) / CDbl(nTreasures)

    If nVisited > 5 * nRooms Then
        dblScore = dblScore - CDbl(nVisited) / (5# * CDbl(nRooms))

        If dblScore < 0# Then
            dblScore = 0#
        End If
    End If

    nScore = Int(dblScore)
    myValROM(2) = nScore & "/100"
 

    
    myValROM(0) = " "
    If strTreasures = "" Then
            Select Case this_Graphic.Player(CurrentPlayerID).Info.MoveDirection
                Case MoveLeft
                    If this_Graphic.Player(CurrentPlayerID).Info.C_Position.x > 0 Then
                        this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x - 1, this_Graphic.Player(i).Info.C_Position.y) = 0
                    End If
                Case MoveRight
                    this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x + 1, this_Graphic.Player(i).Info.C_Position.y) = 0
                Case MoveUp
                    If this_Graphic.Player(CurrentPlayerID).Info.C_Position.y > 0 Then
                        this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x, this_Graphic.Player(i).Info.C_Position.y - 1) = 0
                    End If
                Case MoveDown
                    this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x, this_Graphic.Player(i).Info.C_Position.y + 1) = 0
            End Select
    Else
        myValROM(0) = myValROM(0) & "Carry, "
            Select Case this_Graphic.Player(CurrentPlayerID).Info.MoveDirection
                Case MoveLeft
                    If this_Graphic.Player(CurrentPlayerID).Info.C_Position.x > 0 Then
                        this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x - 1, this_Graphic.Player(i).Info.C_Position.y) = 10
                    End If
                Case MoveRight
                    this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x + 1, this_Graphic.Player(i).Info.C_Position.y) = 10
                Case MoveUp
                    If this_Graphic.Player(CurrentPlayerID).Info.C_Position.y > 0 Then
                        this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x, this_Graphic.Player(i).Info.C_Position.y - 1) = 10
                    End If
                Case MoveDown
                    this_Graphic.Map.Events(this_Graphic.Player(CurrentPlayerID).Info.C_Position.x, this_Graphic.Player(i).Info.C_Position.y + 1) = 10
            End Select
    End If

    If ((nRoom1 = 0) And (bTreasureCarried)) Then
        myValROM(0) = myValROM(0) & "Drop, "
    Else
    End If

    
    If MUD_Rooms(nRoom1).bConnected(0, 0) Then
        If nDimensions = 2 Then
            myValROM(0) = myValROM(0) & "Up, "
        Else
            myValROM(0) = myValROM(0) & "North, "
        End If
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(0, 1) Then
        If nDimensions = 2 Then
            myValROM(0) = myValROM(0) & "Down, "
        Else
            myValROM(0) = myValROM(0) & "South, "
        End If
       ' myValROM(0) = myValROM(0) & "South, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(1, 0) Then
        If nDimensions = 2 Then
            myValROM(0) = myValROM(0) & "Right, "
        Else
            myValROM(0) = myValROM(0) & "East, "
        End If
        'myValROM(0) = myValROM(0) & "East, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(1, 1) Then
        If nDimensions = 2 Then
            myValROM(0) = myValROM(0) & "Left, "
        Else
            myValROM(0) = myValROM(0) & "West, "
        End If
'        myValROM(0) = myValROM(0) & "West, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(2, 0) Then
        myValROM(0) = myValROM(0) & "MoveUp, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(2, 1) Then
        myValROM(0) = myValROM(0) & "MoveDown, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(3, 0) Then
        myValROM(0) = myValROM(0) & "Forward, "
    Else
    End If

    If MUD_Rooms(nRoom1).bConnected(3, 1) Then
        myValROM(0) = myValROM(0) & "Backward, "
    Else
    End If

    'If Trim(MUD_Rooms(nRoom1).strDescription & strTreasures) <> "" Then
        If Trim(MUD_Rooms(nRoom1).strDescription) <> "" Then
            sss = sss & "You're in " & Trim(MUD_Rooms(nRoom1).strDescription) & vbNewLine
        Else
            If Trim(strTreasures) = "" Then sss = sss & "You're in " & Trim(MUD_Rooms(nRoom1).strDescription) & vbNewLine
        End If
        If Trim(strTreasures) <> "" Then sss = sss & Trim(strTreasures) & vbNewLine
    'End If
 
     If InStr(this_Graphic.SpriteFont(0).G_String, vbNewLine) > 0 Then
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, vbNewLine, "-")
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "--", "-")
    End If
    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & vbNewLine & Trim(sss)
    this_Graphic.SpriteFont(0).G_MaxLineLen = 500
    this_Graphic.SpriteFont(0).G_WordWraped = False
    this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "-", vbNewLine & "")
    Call Event_Get(this_Graphic, 1)
    

    
    this_AI_text.x = 0
    this_AI_text.y = 0
    Call Draw_TilesBoardNote(this_Graphic, "R:" & myValROM(3) & " T:" & myValROM(4), 0, 0, this_AI_text, 1)
    'this_AI_text.x = 0
    'this_AI_text.Y = 1
    'Call Draw_TilesBoardNote(this_Graphic, "Treasures" & myValROM(4), 0, 0, this_AI_text, 1)
    Call Type_Check("", myValROM)
End Sub

'=============================================================
'Describe:Data Reading Operations for Four-Dimensional Space Games£¨The main thing is to set up the action of playing in each room.£©
'Author:  James L. Dean &  Chen Lyu
'Parameter:
'=============================================================
Public Function Type_Check(sUserType As String, myValROM() As String)
    Dim sss As String
    Dim t() As String
    If MUD_Practice_Start = True Then
        If LCase(Replace(sUserType, " ", "")) = LCase(Replace(Replace(Mud_Answer(0) & Mud_Answer(1), Chr(10), ""), Chr(13), "")) Then
            sss = sss & UCase(Replace(sUserType, " ", "")) & " is Correct!" & vbNewLine
        Else
            sss = sss & UCase(Replace(sUserType, " ", "")) & " is Wrong..." & vbNewLine
        End If
         MUD_Practice_Start = False
         GoTo RetypeWhere101:
    ElseIf MUD_Practice_Start = False Then
        If LCase(sUserType) = "file.ai.unload" Then
            this_Graphic.AILoaded = False
        Else
            
            Select Case sUserType
                Case "west"
                    If MUD_Rooms(nRoom1).bConnected(1, 1) Then
                        Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                Case "left"
                    If MUD_Rooms(nRoom1).bConnected(1, 1) Then
                        Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                Case "l"
                    If MUD_Rooms(nRoom1).bConnected(1, 1) Then
                        Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                Case "w"
                    If MUD_Rooms(nRoom1).bConnected(1, 1) Then
                        Call Player_Event_ClicktoMove(MoveLeft, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "east"
                    If MUD_Rooms(nRoom1).bConnected(1, 0) Then
                        Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "right"
                    If MUD_Rooms(nRoom1).bConnected(1, 0) Then
                        Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                  Case "r"
                    If MUD_Rooms(nRoom1).bConnected(1, 0) Then
                        Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "e"
                    If MUD_Rooms(nRoom1).bConnected(1, 0) Then
                        Call Player_Event_ClicktoMove(MoveRight, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nYCoordinate = nYCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "north"
                    If MUD_Rooms(nRoom1).bConnected(0, 0) Then
                        Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "up"
                    If MUD_Rooms(nRoom1).bConnected(0, 0) Then
                        Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "u"
                    If MUD_Rooms(nRoom1).bConnected(0, 0) Then
                        Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "n"
                    If MUD_Rooms(nRoom1).bConnected(0, 0) Then
                        Call Player_Event_ClicktoMove(MoveUp, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "south"
                    If MUD_Rooms(nRoom1).bConnected(0, 1) Then
                        Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "down"
                    If MUD_Rooms(nRoom1).bConnected(0, 1) Then
                        Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                Case "d"
                    If MUD_Rooms(nRoom1).bConnected(0, 1) Then
                        Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "s"
                    If MUD_Rooms(nRoom1).bConnected(0, 1) Then
                        Call Player_Event_ClicktoMove(MoveDown, CurrentPlayerID)
                        nMoves = nMoves + 1
                        nXCoordinate = nXCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "forward"
                    If MUD_Rooms(nRoom1).bConnected(3, 0) Then
                        nMoves = nMoves + 1
                        nTCoordinate = nTCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "f"
                    If MUD_Rooms(nRoom1).bConnected(3, 0) Then
                        nMoves = nMoves + 1
                        nTCoordinate = nTCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "backward"
                    If MUD_Rooms(nRoom1).bConnected(3, 1) Then
                        nMoves = nMoves + 1
                        nTCoordinate = nTCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "b"
                    If MUD_Rooms(nRoom1).bConnected(3, 1) Then
                        nMoves = nMoves + 1
                        nTCoordinate = nTCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "movedown"
                    If MUD_Rooms(nRoom1).bConnected(2, 1) Then
                        nMoves = nMoves + 1
                        nZCoordinate = nZCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "md"
                    If MUD_Rooms(nRoom1).bConnected(2, 1) Then
                        nMoves = nMoves + 1
                        nZCoordinate = nZCoordinate + 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "moveup"
                    If MUD_Rooms(nRoom1).bConnected(2, 0) Then
                        nMoves = nMoves + 1
                        nZCoordinate = nZCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                 Case "mu"
                    If MUD_Rooms(nRoom1).bConnected(2, 0) Then
                        nMoves = nMoves + 1
                        nZCoordinate = nZCoordinate - 1
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                  Case "carry"
                    If strTreasures <> "" Then
                        Call Carry(myValROM)
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                  Case "c"
                    If strTreasures <> "" Then
                        Call Carry(myValROM)
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                  Case "drop"
                    If ((nRoom1 = 0) And (bTreasureCarried)) Then
                        Call Drop
                        Call GameUpdate(myValROM)
                    Else
                        GoTo RetypeWhere:
                    End If
                  Case "wayout"
                    Call WayOut(myValROM)
                 Case Else
                    If InStr(1, LCase(Trim(sUserType)), " ") > 0 Then
                        t = Split(LCase(Trim(sUserType)), " ")
                        Select Case UBound(t)
                            Case 1
                                Select Case t(0)
                                    Case "read"
                                        sss = sss & WordsCount(myValROM, t, 0) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "r"
                                        sss = sss & WordsCount(myValROM, t, 0) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "print"
                                        MUD_UserFiles = Files_Files_Read(t(1))
                                        sss = sss & MUD_UserFiles & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "word"
                                        sss = sss & WordsPrint(myValROM, t(1)) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "w"
                                        sss = sss & WordsPrint(myValROM, t(1)) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "practice"
                                        MUD_Practice_Start = True
                                        sss = sss & PracticePrint(myValROM, MUD_Practice, t(), -1) & vbNewLine
                                        sss = sss & Mud_KeyWordsSentencees
                                        GoTo RetypeWhere102:
                                    Case "p"
                                        MUD_Practice_Start = True
                                        sss = sss & PracticePrint(myValROM, MUD_Practice, t(), -1) & vbNewLine
                                        sss = sss & Mud_KeyWordsSentencees
                                        GoTo RetypeWhere102:
                                    Case Else
                                        GoTo RetypeWhere100:
                                End Select
                            Case 2
                                Select Case t(0)
                                    Case "read"
                                        sss = sss & WordsCount(myValROM, t, 8) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "r"
                                        sss = sss & WordsCount(myValROM, t, 8) & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "print"
                                        MUD_UserFiles = Files_Files_Read(t(1))
                                        sss = sss & MUD_UserFiles & vbNewLine
                                        sUserType = t(1)
                                        GoTo RetypeWhere101:
                                    Case "practice"
                                        MUD_Practice_Start = True
                                        sss = sss & PracticePrint(myValROM, MUD_Practice, t(), 8) & vbNewLine
                                        sss = sss & Mud_KeyWordsSentencees
                                        GoTo RetypeWhere102:
                                    Case "p"
                                        MUD_Practice_Start = True
                                        sss = sss & PracticePrint(myValROM, MUD_Practice, t(), 8) & vbNewLine
                                        sss = sss & Mud_KeyWordsSentencees
                                        GoTo RetypeWhere102:
                                    Case Else
                                        GoTo RetypeWhere100:
                                End Select
                            Case Else
                                GoTo RetypeWhere100:
                        End Select
                    Else
                        GoTo RetypeWhere100:
                    End If
RetypeWhere100:
                    ReDim t(1)
                    t(0) = Trim(sUserType)
                    'sss = sss & WordsCount(myValROM, t, -1) & vbNewLine
                    'Call WordsCount(myValROM, t, -1)
RetypeWhere101:
                    sss = sss & AIRobotResponse(sUserType, MUD_Files) & vbNewLine
RetypeWhere102:
                    
                    If InStr(this_Graphic.SpriteFont(0).G_String, vbNewLine) > 0 Then
                        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, vbNewLine, "-")
                        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "--", "-")
                    End If
'                    this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "-", vbNewLine)
                    
                    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & vbNewLine & Trim(sss)
                    this_Graphic.SpriteFont(0).G_MaxLineLen = 500
                    this_Graphic.SpriteFont(0).G_WordWraped = False
                    
                    Call Event_Get(this_Graphic, 1)
                    this_AI_text.x = 0
                    this_AI_text.y = 0
                    Call Draw_TilesBoardNote(this_Graphic, "R:" & myValROM(3) & " T:" & myValROM(4), 2, 0, this_AI_text, 1)
                    this_AI_text.x = 260
                    this_AI_text.y = 34
                    
                    Call Draw_TilesBoardNote(this_Graphic, myValROM(0), 97, 400, this_AI_text, 5)
                    DoEvents
    
            End Select
        End If
    End If
RetypeWhere:
End Function

'=============================================================
'Describe:Functions of TOEFL Vocabulary Learning
'Author:  Chen Lyu
'Parameter:
'=============================================================
Public Function PracticePrint(myValROM() As String, myPractice() As tMUD_Practice, myt() As String, Optional myLV As Integer = 0) As String
 
    Dim i As Integer
    Dim j As Integer
    Dim s(4) As String
    Dim myPractice_ID As Integer
    s(0) = "A."
    s(1) = "B."
    s(2) = "C."
    s(3) = "D."
    
    myPractice_ID = Val(myt(1)) - 1
    If myPractice_ID > UBound(myPractice) Or myPractice_ID < 0 Then myPractice_ID = 0
    Mud_KeyWordsSentence(0) = ""
    Mud_KeyWordsSentence(1) = ""
    Mud_Answer(0) = ""
    Mud_Answer(1) = ""
    With myPractice(myPractice_ID)
        myt(0) = Cutoff(Replace(Replace(.Material, Chr(10), ""), Chr(13), ""))
        PracticePrint = myt(0)
        PracticePrint = PracticePrint & vbNewLine & WordsCount(myValROM, myt, myLV)
        For i = 0 To 1
            'PracticePrint = PracticePrint & vbNewLine & .KeyWordsSentence(i)
            Mud_KeyWordsSentence(i) = Replace(Replace(.KeyWordsSentence(i), Chr(10), ""), Chr(13), "")
            For j = 1 To 4
                PracticePrint = PracticePrint & vbNewLine & s(j - 1) & Replace(Replace(.Options(i, j), Chr(10), ""), Chr(13), "")
            Next j
            Mud_Answer(i) = .Answer(i)
'            PracticePrint = PracticePrint & vbNewLine & .Answer(i)
            If .KeyWordsSentence(1) = "" Then Exit For
        Next i
    End With
    Mud_KeyWordsSentencees = Mud_KeyWordsSentence(0) & vbNewLine & Mud_KeyWordsSentence(1)
    
End Function

Public Function Cutoff(myString As String) As String
    
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    s = Replace(myString, "", "")
    j = 0
    Cutoff = ""
    Do
        i = InStr(100, s, " ")
        If i > 0 Then
            Cutoff = Cutoff & Left(s, i) & vbNewLine
            j = Len(s) - i
            If j < 0 Then j = 0
            s = Right(Replace(myString, "", ""), j)
        Else
            Cutoff = Cutoff & s
        End If
    Loop Until i = 0

End Function


Public Function WordsPrint(myValROM() As String, myt As String) As String
    Dim j As Long
    Dim i As Long
    Dim V As Integer
    Dim strKey
    Dim s As String
    ReDim MUD_UserWordsLevel(12)
        For Each strKey In oDict.keys
            If LCase(Trim(myt)) = strKey Then
                j = Val(oDict.Item(strKey))
'                WordsPrint = MUD_Word_Words(j).Words & "  LV= " & MUD_Word_Words(j).Frequency & vbNewLine & MUD_Word_Words(j).Explain & MUD_Word_Words(j).EEDef          ' vbNewLine
               WordsPrint = MUD_Word_Words(j).Words & "[" & MUD_Word_Words(j).Frequency & "] means:" & vbNewLine & MUD_Word_Words(j).EEDef          ' vbNewLine
                MUD_User.Words(j).TypeCount = MUD_User.Words(j).TypeCount + 1
                
                Exit For
            End If
        Next strKey
        For V = 0 To 11
            oDictLVs(V).RemoveAll
        Next V
        V = MUD_Word_Words(j).Frequency
        If oDictLVs(V).Exists(s) Then
            oDictLVs(V).Item(s) = oDict.Item(s)
        Else
            oDictLVs(V).Add (s), j
        End If
                        
    'MUD_UserWordsLevel(MUD_Word_Words(j).Frequency) = MUD_UserWordsLevel(MUD_Word_Words(j).Frequency) + 1
    myValROM(MUD_Word_Words(j).Frequency + 7) = myValROM(MUD_Word_Words(j).Frequency + 7) + 1
    s = ""
    For i = 0 To 11
        s = s & "  LV" & i + 1 & "=" & oDictLVs(i).Count
    Next i
    myValROM(19) = s
    Call Files_MUDUser_Save(MUD_UserName, MUD_User)
End Function
Public Function WordsCount(myValROM() As String, myt() As String, Optional myLV As Integer = 0) As String
    Dim t() As String
    Dim i As Long
    Dim j As Long
    Dim V As Integer
    Dim s As String
    Dim strKey
    
    If myLV >= 0 And myLV < 9 Then
        MUD_UserFiles = Files_Files_Read(myt(1))
    Else
        MUD_UserFiles = myt(0)
    End If
    
    MUD_UserWords = Split(LCase(Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(MUD_UserFiles, "?", " "), "'", " "), """", " "), ";", " "), ".", " "), ",", " "), ":", " "), "!", " ")) & " "), " ")
    
    ReDim MUD_UserWordsLevel(12)
    'ReDim MUD_UserLevelWordsID(1)
    oDictLV.RemoveAll
    For V = 0 To 11
        oDictLVs(V).RemoveAll
    Next V
    For i = 0 To UBound(MUD_UserWords)
        s = Trim(Replace(MUD_UserWords(i), Chr(0), ""))
           If s <> "" Then
               If oDict.Exists(s) Then
                    j = Val(oDict.Item(s))
                   'MUD_UserWordsLevel(MUD_Word_Words(j).Frequency) = MUD_UserWordsLevel(MUD_Word_Words(j).Frequency) + 1
                   myValROM(MUD_Word_Words(j).Frequency + 7) = myValROM(MUD_Word_Words(j).Frequency + 7) + 1
                   MUD_User.Words(j).TypeCount = MUD_User.Words(j).TypeCount + 1
                   
                        V = MUD_Word_Words(j).Frequency
                        If oDictLVs(V).Exists(s) Then
                            oDictLVs(V).Item(s) = oDict.Item(s)
                        Else
                            oDictLVs(V).Add (s), j
                        End If
                        If myLV > 0 Then
                            If IsNumeric(myt(2)) = True Then
                                
                                    If MUD_Word_Words(j).Frequency = myt(2) Then
                                        If oDictLV.Exists(s) Then
                                            oDictLV.Item(s) = oDict.Item(s)
                                        Else
                                            oDictLV.Add (s), j
                                        End If
                                    End If
                                
                            End If
                        End If
               Else
                    Debug.Print s
                   'oDict.Add (s), s
                   MUD_UserWordsLevel(11) = MUD_UserWordsLevel(11) + 1
               End If
           End If
    Next i
        s = ""
        For i = 0 To 11
            s = s & "  LV" & i + 1 & "=" & oDictLVs(i).Count
        Next i
        myValROM(19) = s
        s = ""
    If myLV > 1 Then
        s = ""
        s = "Count=" & UBound(MUD_UserWords) & vbNewLine
        i = 1
        For Each strKey In oDictLV.keys
            s = s & Left(strKey & Space(50), 20)
            If i Mod 7 = 0 Then s = s & vbNewLine
            i = i + 1
        Next strKey
        's = s & Cutoff(MUD_UserFiles)
    Else
        s = "Count=" & UBound(MUD_UserWords) '& " " & s
    End If
    WordsCount = s
    Call Files_MUDUser_Save(MUD_UserName, MUD_User)
End Function


Public Sub Carry(myValROM() As String)
    nTreasure1 = 0
    Dim sss As String

    Do While nTreasure1 < nTreasures

        If MUD_Treasure(nTreasure1).nWeaponRoom = nRoom1 Then
            MUD_Treasure(nTreasure1).nWeaponRoom = -1
        End If

        nTreasure1 = nTreasure1 + 1
    Loop

    nTreasure1 = 0

    Do While nTreasure1 < nTreasures

        If MUD_Treasure(nTreasure1).nTreasureRoom = nRoom1 Then
            If MUD_Treasure(nTreasure1).nWeaponRoom < 0 Then
                MUD_Treasure(nTreasure1).nTreasureRoom = -1
                nTreasuresRecovered = nTreasuresRecovered + 1

                If MUD_Treasure(nTreasure1).nGuardRoom = nRoom1 Then
                    MUD_Treasure(nTreasure1).nGuardRoom = -1
                    MUD_Treasure(nTreasure1).nWeaponRoom = -2
                    sss = sss & Trim("Way to go! You're " & MUD_Treasure(nTreasure1).strWeapon & " overcomes the " & MUD_Treasure(nTreasure1).strGuard & ".")
                End If

            Else
                sss = sss & Trim("Whoops! You carry nothing to overcome the " & MUD_Treasure(nTreasure1).strGuard & ".")
            End If
        End If

        If MUD_Treasure(nTreasure1).nWeaponRoom = nRoom1 Then
            MUD_Treasure(nTreasure1).nWeaponRoom = -1
        End If

        nTreasure1 = nTreasure1 + 1
    Loop

    
    If InStr(this_Graphic.SpriteFont(0).G_String, vbNewLine) > 0 Then
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, vbNewLine, "-")
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "--", "-")
    End If
'                    this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "-", vbNewLine)
    
    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & vbNewLine & Trim(sss)
    this_Graphic.SpriteFont(0).G_MaxLineLen = 500
    this_Graphic.SpriteFont(0).G_WordWraped = False
    
    Call Event_Get(this_Graphic, 1)
 
 End Sub


Public Sub Drop()
    nTreasure1 = 0

    Do While nTreasure1 < nTreasures

        If MUD_Treasure(nTreasure1).nTreasureRoom = -1 Then
            MUD_Treasure(nTreasure1).nTreasureRoom = 0
        End If

        nTreasure1 = nTreasure1 + 1
    Loop

End Sub

Public Sub WayOut(myValROM() As String)
    Dim Response As Long
    bPathFound = False
    Dim sss As String
    
    If ((bTreasureCarried) And (nRoom1 <> 0)) Then
        nCoordinate(0) = nXCoordinate
        nCoordinate(1) = nYCoordinate
        nCoordinate(2) = nZCoordinate
        nCoordinate(3) = nTCoordinate
        nWayOutHead = 0
        nRoom2 = 0

        Do While nRoom2 < nRooms
            MUD_Rooms(nRoom2).bRoomUsed = False
            nRoom2 = nRoom2 + 1
        Loop

        MUD_Rooms(nRoom1).bRoomUsed = True
        nDirectionsUsed(nWayOutHead) = 0
        nDirectionsPossible = 2 * nDimensions
        nDimension1 = 0

        Do While nDimension1 < nDimensions
            nDirection1 = 0

            Do While nDirection1 < 2
                bDirectionUsed(nWayOutHead, nDimension1, nDirection1) = False
                nDirection1 = nDirection1 + 1
            Loop

            nDimension1 = nDimension1 + 1
        Loop

        strWayOut = ""
        nRoom2 = nRoom1
        nTrial = 0

        Do While (nTrial < 500) And (nRoom2 <> 0) And (nWayOutHead < 255)
            nTrial = nTrial + 1
            bDirectionFound = False

            Do While (Not bDirectionFound) And (nDirectionsUsed(nWayOutHead) < nDirectionsPossible)
                Randomize
                nDirection1 = Int(2# * Rnd)
                Randomize
                nDimension1 = Int(CDbl(nDimensions) * Rnd)

                If (Not bDirectionUsed(nWayOutHead, nDimension1, nDirection1)) Then
                    bDirectionUsed(nWayOutHead, nDimension1, nDirection1) = True
                    nDirectionsUsed(nWayOutHead) = nDirectionsUsed(nWayOutHead) + 1

                     
                    If MUD_Rooms(nRoom2).bConnected(nDimension1, nDirection1) Then
                        nCoordinateNext(0) = nCoordinate(0)
                        nCoordinateNext(1) = nCoordinate(1)
                        nCoordinateNext(2) = nCoordinate(2)
                        nCoordinateNext(3) = nCoordinate(3)
                        nCoordinateNext(nDimension1) = nCoordinate(nDimension1) + 2 * nDirection1 - 1

                        If (Not bEuclidean) Then
                            If nCoordinateNext(nDimension1) < 0 Then
                                nDimension2 = 0

                                Do While nDimension2 < nDimensions
                                    nCoordinateNext(nDimension2) = nWidth(nDimension2) - nCoordinateNext(nDimension2) - 1
                                    nDimension2 = nDimension2 + 1
                                Loop

                                nCoordinateNext(nDimension1) = nWidth(nDimension1) - 1
                            Else

                                If nCoordinateNext(nDimension1) >= nWidth(nDimension1) Then
                                    nDimension2 = 0

                                    Do While nDimension2 < nDimensions
                                        nCoordinateNext(nDimension2) = nWidth(nDimension2) - nCoordinateNext(nDimension2) - 1
                                        nDimension2 = nDimension2 + 1
                                    Loop

                                    nCoordinateNext(nDimension1) = 0
                                End If
                            End If
                        End If

                        If (Not MUD_Rooms(nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3))).bRoomUsed) Then
                            bDirectionFound = True
                        End If
                    End If
                End If

            Loop

            If bDirectionFound Then
                nRoom2 = nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3))
                nWayOutHead = nWayOutHead + 1
                MUD_Rooms(nRoom2).bRoomUsed = True
                nDirectionsUsed(nWayOutHead) = 0
                nDimension2 = 0

                Do While nDimension2 < nDimensions
                    nDirection2 = 0

                    Do While nDirection2 < 2
                        bDirectionUsed(nWayOutHead, nDimension2, nDirection2) = False
                        nDirection2 = nDirection2 + 1
                    Loop

                    nDimension2 = nDimension2 + 1
                Loop

                nWayOutDimension(nWayOutHead) = nDimension1
                nWayOutDirection(nWayOutHead) = 1 - nDirection1

                Select Case nDimension1

                    Case 0

                        If nDirection1 = 0 Then
                            strWayOut = strWayOut & "N"
                        Else
                            strWayOut = strWayOut & "S"
                        End If

                    Case 1

                        If nDirection1 = 0 Then
                            strWayOut = strWayOut & "E"
                        Else
                            strWayOut = strWayOut & "W"
                        End If

                    Case 2

                        If nDirection1 = 0 Then
                            strWayOut = strWayOut & "U"
                        Else
                            strWayOut = strWayOut & "D"
                        End If

                    Case Else

                        If nDirection1 = 0 Then
                            strWayOut = strWayOut & "F"
                        Else
                            strWayOut = strWayOut & "B"
                        End If

                End Select

            Else
                nDirection1 = nWayOutDirection(nWayOutHead)
                nDimension1 = nWayOutDimension(nWayOutHead)
                nCoordinateNext(0) = nCoordinate(0)
                nCoordinateNext(1) = nCoordinate(1)
                nCoordinateNext(2) = nCoordinate(2)
                nCoordinateNext(3) = nCoordinate(3)
                nCoordinateNext(nDimension1) = nCoordinateNext(nDimension1) + 2 * nDirection1 - 1

                If (Not bEuclidean) Then
                    If nCoordinateNext(nDimension1) < 0 Then
                        nDimension2 = 0

                        Do While nDimension2 < nDimensions
                            nCoordinateNext(nDimension2) = nWidth(nDimension2) - nCoordinateNext(nDimension2) - 1
                            nDimension2 = nDimension2 + 1
                        Loop

                        nCoordinateNext(nDimension1) = nWidth(nDimension1) - 1
                    Else

                        If nCoordinateNext(nDimension1) >= nWidth(nDimension1) Then
                            nDimension2 = 0

                            Do While nDimension2 < nDimensions
                                nCoordinateNext(nDimension2) = nWidth(nDimension2) - nCoordinateNext(nDimension2) - 1
                                nDimension2 = nDimension2 + 1
                            Loop

                            nCoordinateNext(nDimension1) = 0
                        End If
                    End If
                End If

                nRoom2 = nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3))
                nWayOutHead = nWayOutHead - 1

                If Len(strWayOut) > 1 Then
                    strWayOut = Left(strWayOut, Len(strWayOut) - 1)
                Else
                    strWayOut = ""
                End If
            End If

            nCoordinate(0) = nCoordinateNext(0)
            nCoordinate(1) = nCoordinateNext(1)
            nCoordinate(2) = nCoordinateNext(2)
            nCoordinate(3) = nCoordinateNext(3)
        Loop


        If nRoom2 = 0 Then
            bPathFound = True
        End If
    End If

    If bPathFound Then
        nTreasure1 = 0
        nRoom2 = 0

        Do While (nTreasure1 < nTreasures) And (nRoom2 >= 0)
            nRoom2 = MUD_Treasure(nTreasure1).nTreasureRoom

            If nRoom2 >= 0 Then
                nTreasure1 = nTreasure1 + 1
            End If

        Loop

        nRoom2 = nRoom1

        Do While nRoom1 = nRoom2
            Randomize
            nRoom2 = 1 + Int(CDbl(nRooms - 1) * Rnd)
        Loop

        MUD_Treasure(nTreasure1).nTreasureRoom = nRoom2
        sss = sss & Trim("The pirate takes one of your treasures. As he leaves,  he shouts the letters, '" & strWayOut & "'.")
        Call GameUpdate(myValROM)
    Else
        sss = sss & Trim("Nothing happens. Try again later.")
    End If

    If InStr(this_Graphic.SpriteFont(0).G_String, vbNewLine) > 0 Then
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, vbNewLine, "-")
        this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "--", "-")
    End If
'                    this_Graphic.SpriteFont(0).G_String = ReplacePlus(this_Graphic.SpriteFont(0).G_String, "-", vbNewLine)
    
    this_Graphic.SpriteFont(0).G_String = this_Graphic.SpriteFont(0).G_String & vbNewLine & Trim(sss)
    this_Graphic.SpriteFont(0).G_MaxLineLen = 500
    this_Graphic.SpriteFont(0).G_WordWraped = False
    
    Call Event_Get(this_Graphic, 1)
 
                    
End Sub

