Attribute VB_Name = "m_File"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' File Processing Functions

'====================================================================================================================================


Option Explicit

'=============================================================
'Describe:Read Code Files
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub CodeFileLoad(this_Graphic As m_Graphic, Optional m_FileName As String)
    Dim s
    Dim i As Integer
    
    With this_Graphic
    
    
        If Trim(m_FileName) <> "" Then
            .Code.Autorun = False
            .Code.Text = ""
            For i = 0 To 99
                Steps(i).Direction = -1
                Steps(i).Visible = False
            Next i
            For i = 0 To 3
                .Map.Windows(i).Visible = True
            Next i
            Steps(0).Position.x = .Player(CurrentPlayerID).Info.C_Position.x
            Steps(0).Position.y = .Player(CurrentPlayerID).Info.C_Position.y
            For i = 0 To 63
                .Map.GameEvents(i).m_Description = ""
            Next
            CurrentGameEvents = 63
'
            .Code.Text = FileRead(this_FilePath.Code & Trim(m_FileName))
            .Code.mName = m_FileName

            
            this_FilePath.CodeName = m_FileName
            flagChange = False
            If Trim(.Code.Text) <> "" Then
                For Each s In Split(.Code.Text, vbCrLf)
                     If LCase(Trim(s)) = "game.autorun" Then
                        .Code.Autorun = True
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    Call fMainForm.File_thisLoad
End Sub

'=============================================================
'Describe:New Code File
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub FileNew()
    fMainForm.txtSource.Text = ""
    this_FilePath.CodeName = ""
    flagChange = False
End Sub

'=============================================================
'Describe:Save the current file
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub FileSave()
    If this_FilePath.CodeName = "" Then
        Call FileSaveAs
    Else
        FileWrite this_FilePath.CodeName
    End If
End Sub

'=============================================================
'Describe:Save the current file
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub FileSaveAs()
    With fMainForm.dlgCommonDialog
        .FileName = ""
        .Filter = "Plarn(*.pl)|*.pl"
        .ShowSave
        If .FileName = "" Then
            Exit Sub
        Else
            FileWrite .FileName
            this_FilePath.CodeName = .FileName
        End If
    End With
End Sub

Public Sub FileWrite(ByVal Name As String)
    Dim Fso As New FileSystemObject
    Dim f As TextStream
    Set f = Fso.OpenTextFile(Name, ForWriting, True)
    f.Write fMainForm.txtSource.Text
    f.Close
    flagChange = False
End Sub

Public Function FileRead(ByVal Name As String)
    Dim Fso As New FileSystemObject
    Dim f As TextStream
    Set f = Fso.OpenTextFile(Name, ForReading, True)
    FileRead = f.ReadAll
    f.Close
    flagChange = False
End Function

'=============================================================
'Describe:Save the map file
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Files_Data_Map_Save(this_Graphic As m_Graphic, this_FilePath As m_FilePath, m_Name As String)

    Dim Filenr As Integer
    Dim s As String
    
    Select Case m_Name
    
        Case "Map"
            s = this_FilePath.Code & m_Name & ".map"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Write As #Filenr
                DoEvents
                Put #Filenr, , this_Graphic.Map.Letter
                DoEvents
            Close #Filenr
        Case "Player"
            s = this_FilePath.CourseMap & m_Name & ".player"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Write As #Filenr
                DoEvents
                Put #Filenr, , this_Graphic.Player
                DoEvents
            Close #Filenr
        Case "Story"
            s = this_FilePath.Story & m_Name & ".Story"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Write As #Filenr
                DoEvents
                'Put #Filenr, , this_Graphic.Story
                DoEvents
            Close #Filenr
            
    End Select

End Sub

'=============================================================
'Describe:Read the map file
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Files_Data_Map_Read(this_Graphic As m_Graphic, this_FilePath As m_FilePath, m_Name As String)
    Dim Filenr As Integer
    Dim s As String
    Dim i As Long
    
    Select Case m_Name
    
        Case "Map"
            s = this_FilePath.Code & m_Name & ".map"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Read As #Filenr
                DoEvents
                Get #Filenr, , this_Graphic.Map.Letter
                DoEvents
            Close #Filenr
            
        Case "Player"
            s = this_FilePath.CourseMap & m_Name & ".player"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Read As #Filenr
                DoEvents
                Get #Filenr, , this_Graphic.Player
                DoEvents
            Close #Filenr
            
        Case "Story"
            s = this_FilePath.Story & m_Name & ".Story"
            Filenr = FreeFile
            DoEvents
            Open s For Binary Access Read As #Filenr
                DoEvents
                'Get #Filenr, , this_Graphic.Story
                DoEvents
            Close #Filenr
         
    End Select
 

End Sub
