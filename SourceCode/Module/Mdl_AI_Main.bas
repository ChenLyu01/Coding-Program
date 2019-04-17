Attribute VB_Name = "Mdl_AI_Main"
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


'=============================================================
'Describe:Starting Function of AI Main Program
'Author:  Chen Lyu
'Parameter:
'=============================================================
Public Sub AI_Main(Optional m_index As Byte)
    Dim sUserTypes As String
    Dim i As Integer
    Dim sPath As String
    
    Dim s As String
    
    If this_Graphic.AILoaded = False Then
    
        Set oDict = CreateObject("Scripting.Dictionary")
        Set oDictLV = CreateObject("Scripting.Dictionary")
        For i = 0 To 11
            Set oDictLVs(i) = CreateObject("Scripting.Dictionary")
        Next i
        
        Call Files_my_Files_Read(MUD_Files, "Plarn_Index")
    
         '=======================AI==========================
             Call MUD_Info(MUD_DialogueROM)
        Call Files_Data_Correlation_Read(MUD_CorrelationWords)
        Call Files_Data_CorrelationIndex_Read(MUD_Correlation_Index)
        Call Files_Data_DataAddition_Read(MUD_Word_Addition)
        Call Files_DataIrregularVerb_Read(MUD_IrregularVerb)
        Call Files_DataMember_Read(MUD_Word_Member)
        Call Files_Data_DataFrequency_Read(MUD_Word_Frequency())
        Call Files_DataWordsWords_Read(MUD_Word_Words)
        
        
         '=======================AI==========================
         
    
        MUD_Practice_Start = False
        MUD_Practice_ID = 0
    
        ReDim MUD_ValROM(20)
        For i = 0 To 19
            MUD_ValROM(i) = 0
        Next i
        MUD_ValROM(2) = "0/100"
        
        bEuclidean = True
        
        If m_index >= 2 And m_index <= 4 Then
            nDimensions = 2 ' m_index
        Else
            nDimensions = 2
        End If
    
        s = MUD_DialogueROM(0).Text(0) & MUD_DialogueROM(1).Text(0) & m_index & MUD_DialogueROM(2).Text(0)
        
        DoEvents
        If InStr(s, "\n ") > 0 Then
            s = ReplacePlus(s, "\n ", vbNewLine)
        End If
        this_Graphic.SpriteFont(0).G_String = Trim(s)
        this_Graphic.SpriteFont(0).G_MaxLineLen = 500
        this_Graphic.SpriteFont(0).G_WordWraped = False
        Call Event_Get(this_Graphic, 1)
        
        Call Draw_TilesBoardNote(this_Graphic, "Zork VII", 0, 0, this_AI_text, 1)
        
        DoEvents
        
        
        MUD_UserName = this_Graphic.Player(CurrentPlayerID).mName
        sPath = this_FilePath.Code & MUD_UserName & ".Usr"
        If PathFileExists(sPath) = 0 Then
            With MUD_User
                ReDim .Words(UBound(MUD_Word_Words))
            End With
            Call Files_MUDUser_Save(MUD_UserName, MUD_User)
        Else
            Call Files_MUDUser_Read(MUD_UserName, MUD_User)
            With MUD_User
                .RegCount = .RegCount + 1
            End With
            'Call Files_MUDUser_Save(MUD_UserName, MUD_User)
        End If
'
    '
        Call Files_Data_Factor_Read(MUD_FactorWords)
        Call Files_Data_FactorIndex_Read(MUD_Factor_Index)
        Call Files_Data_Treasure_Read(MUD_Treasure)
        Call Files_Data_Rooms_Read(MUD_Rooms)
        Call Files_Data_Practice_Read(MUD_Practice)
        Call WordsControl
        
    
    this_Graphic.AILoaded = True
    
    Else
        If m_index >= 2 And m_index <= 4 Then
            nDimensions = m_index
        Else
            nDimensions = 4
        End If
        s = MUD_DialogueROM(0).Text(0) & MUD_DialogueROM(1).Text(0) & m_index & MUD_DialogueROM(2).Text(0)
        
        DoEvents
        If InStr(s, "\n ") > 0 Then
            s = ReplacePlus(s, "\n ", vbNewLine)
        End If
        this_Graphic.SpriteFont(0).G_String = Trim(s)
        this_Graphic.SpriteFont(0).G_MaxLineLen = 520
        this_Graphic.SpriteFont(0).G_WordWraped = False
        Call Event_Get(this_Graphic, 1)
        
        Call Draw_TilesBoardNote(this_Graphic, "Zork VII", 0, 0, this_AI_text, 1)
    
        MUD_UserName = this_Graphic.Player(CurrentPlayerID).mName
        sPath = this_FilePath.Code & MUD_UserName & ".Usr"
        If PathFileExists(sPath) = 0 Then
            With MUD_User
                ReDim .Words(UBound(MUD_Word_Words))
            End With
            Call Files_MUDUser_Save(MUD_UserName, MUD_User)
        Else
            Call Files_MUDUser_Read(MUD_UserName, MUD_User)
            With MUD_User
                .RegCount = .RegCount + 1
            End With
            'Call Files_MUDUser_Save(MUD_UserName, MUD_User)
        End If
    End If
'
    Call MUD_Run
'
End Sub

'=============================================================
'Describe:Map Generator in Four-Dimensional Space £¨This function needs a lot of modification.£©
'Author:   James L. Dean & Chen Lyu
'Parameter:
'=============================================================
Private Sub MUD_Run()
    Dim strLine        As String
    Dim filz           As String
    Dim desc2          As String



    
    strWayOut = ""
    
    Dim i As Integer
    For i = 0 To 9
        MUD_User.HisScore(i) = MUD_User.HisScore(i + 1)
        MUD_ValROM(8 + i) = MUD_User.level(0 + i)
    Next i
  
    
    
    nTreasures = UBound(MUD_Treasure)
    nRooms = UBound(MUD_Rooms)
  
 
    nTreasure1 = 0
    nRoom1 = 0
    
    
'    Call Files_Data_Treasure_Save(MUD_Treasure)
'    Call Files_Data_Rooms_Save(MUD_Rooms)
    
    nMaxWidth = 1 + Int(CDbl(2 * nRooms) ^ (1# / CDbl(nDimensions)))
    bWidthsFound = False

    Do While Not bWidthsFound
        nDimension1 = 0
        nVolume = 1

        Do While nDimension1 < nDimensions
            Randomize
            nWidth(nDimension1) = nMaxWidth - Int(2# * Rnd)
            nVolume = nVolume * nWidth(nDimension1)
            nDimension1 = nDimension1 + 1
        Loop

        If nVolume > nRooms Then
            bWidthsFound = True
        End If

    Loop

    nDimension1 = nDimensions

    Do While nDimension1 < 4
        nWidth(nDimension1) = 1
        nDimension1 = nDimension1 + 1
    Loop

    nRoom1 = 1

    Do While nRoom1 < nRooms
        Randomize
        nRoom2 = 1 + Int(CDbl(nRooms - 1) * Rnd)
        strLine = MUD_Rooms(nRoom1).strDescription
        MUD_Rooms(nRoom1).strDescription = MUD_Rooms(nRoom2).strDescription
        MUD_Rooms(nRoom2).strDescription = strLine
        nRoom1 = nRoom1 + 1
    Loop

    nXCoordinate = 0

    Do While nXCoordinate < nWidth(0)
        nYCoordinate = 0

        Do While nYCoordinate < nWidth(1)
            nZCoordinate = 0

            Do While nZCoordinate < nWidth(2)
                nTCoordinate = 0

                Do While nTCoordinate < nWidth(3)
                    nCell(nXCoordinate, nYCoordinate, nZCoordinate, nTCoordinate) = -1
                    nTCoordinate = nTCoordinate + 1
                Loop

                nZCoordinate = nZCoordinate + 1
            Loop

            nYCoordinate = nYCoordinate + 1
        Loop

        nXCoordinate = nXCoordinate + 1
    Loop

    nXCoordinate = 0
    nYCoordinate = 0
    nZCoordinate = 0
    nTCoordinate = 0
    nCoordinate(0) = nXCoordinate
    nCoordinate(1) = nYCoordinate
    nCoordinate(2) = nZCoordinate
    nCoordinate(3) = nTCoordinate
    nRoom1 = 0
    nRoom2 = 0
    nCell(0, 0, 0, 0) = 0

    Do While nRoom1 < (nRooms - 1)
        bDirectionFound = False

        Do While Not bDirectionFound
            Randomize
            nDirection1 = Int(2# * Rnd)
            Randomize
            nDimension1 = Int(CDbl(nDimensions) * Rnd)

            If bEuclidean Then
                If nCoordinate(nDimension1) + 2 * nDirection1 - 1 >= 0 Then
                    If nCoordinate(nDimension1) + 2 * nDirection1 - 1 < nWidth(nDimension1) Then
                        bDirectionFound = True
                    End If
                End If

            Else
                bDirectionFound = True
            End If

        Loop
        MUD_Rooms(nRoom2).bConnected(nDimension1, nDirection1) = True
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

        If nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3)) < 0 Then
            nRoom1 = nRoom1 + 1
            nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3)) = nRoom1
        End If

        nRoom2 = nCell(nCoordinateNext(0), nCoordinateNext(1), nCoordinateNext(2), nCoordinateNext(3))
        MUD_Rooms(nRoom2).bConnected(nDimension1, 1 - nDirection1) = True

        nCoordinate(0) = nCoordinateNext(0)
        nCoordinate(1) = nCoordinateNext(1)
        nCoordinate(2) = nCoordinateNext(2)
        nCoordinate(3) = nCoordinateNext(3)
    Loop

    nTreasure1 = 0

    Do While nTreasure1 < nTreasures
        Randomize
        MUD_Treasure(nTreasure1).nTreasureRoom = 1 + Int(CDbl(nRooms - 1) * Rnd)
        MUD_Treasure(nTreasure1).nGuardRoom = MUD_Treasure(nTreasure1).nTreasureRoom
        bWeaponRoomFound = False

        Do While Not bWeaponRoomFound
            Randomize
            MUD_Treasure(nTreasure1).nWeaponRoom = 1 + Int(CDbl(nRooms - 1) * Rnd)

            If MUD_Treasure(nTreasure1).nWeaponRoom <> MUD_Treasure(nTreasure1).nTreasureRoom Then
                bWeaponRoomFound = True
            End If

        Loop

        nTreasure1 = nTreasure1 + 1
    Loop

    bInitialized = True
    Call GameUpdate(MUD_ValROM)
 

End Sub

'=============================================================
'Describe:Uninstall AI program
'Author:   James L. Dean & Chen Lyu
'Parameter:
'=============================================================
Public Sub AI_Unload()
  Dim Response As Long
  If bInitialized Then
    If nScore < 20 Then
      Response = MsgBox("Your score ranks you as a beginner.", vbOKOnly, "You scored " & CStr(nScore) & " out of 100 points.")
    Else
      If nScore < 40 Then
        Response = MsgBox("Your score ranks you as a novice adventurer.", vbOKOnly, "You scored " & CStr(nScore) & " out of 100 points.")
      Else
        If nScore < 60 Then
          Response = MsgBox("Your score ranks you as a seasoned explorer.", vbOKOnly, "You scored " & CStr(nScore) & " out of 100 points.")
        Else
          If nScore < 80 Then
            Response = MsgBox("Your score ranks you as a grissly old prospector.", vbOKOnly, "You scored " & CStr(nScore) & " out of 100 points.")
          Else
            Response = MsgBox("Your score ranks you as an expert treasure hunter;  there is no higher rating.", vbOKOnly, "You scored " & CStr(nScore) & " out of 100 points.")
          End If
        End If
      End If
    End If
  End If
End Sub
