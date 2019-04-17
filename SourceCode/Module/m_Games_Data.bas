Attribute VB_Name = "m_Games_Data"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

'Functions for handling text displayed in a form

'====================================================================================================================================


Option Explicit

Public Sub Draw_TilesBoardNote(this_Graphic As m_Graphic, m_text As String, SpriteFontID As Integer, SpriteFontLen As Integer, m_Position As m_Position, Optional m_Style As Byte, Optional IsSpriteFont As Boolean)
    Dim s As String
    Dim i As Integer
    Dim m_Chr As String
    Dim i_Chr As Integer
    Dim Ascii() As Byte
    Dim a, b, c As Byte
    
    With this_Graphic
    
        If m_Style > 3 Then
            If IsSpriteFont = True Then
                .SpriteFont(SpriteFontID).G_String = m_text
                .SpriteFont(SpriteFontID).x = m_Position.x
                .SpriteFont(SpriteFontID).y = m_Position.y
                .SpriteFont(SpriteFontID).G_WordWraped = False
                .SpriteFont(SpriteFontID).Visiable = True
                .SpriteFont(SpriteFontID).G_MaxLineLen = SpriteFontLen
            Else
                .WindowsFont(SpriteFontID).G_String = m_text
                .WindowsFont(SpriteFontID).x = m_Position.x
                .WindowsFont(SpriteFontID).y = m_Position.y
                .WindowsFont(SpriteFontID).G_WordWraped = False
                .WindowsFont(SpriteFontID).Visiable = True
                .WindowsFont(SpriteFontID).G_MaxLineLen = SpriteFontLen
            End If

        Else
        
            'm_Text = UCase("A")
                If InStr(m_text, "\n ") > 0 Then
                    m_text = ReplacePlus(m_text, "\n ", vbNewLine)
                End If
                Ascii() = StrConv(m_text, vbFromUnicode)
     
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
            
                'Check keyword characters
                For i = 1 To Len(m_text)
                    m_Chr = Asc(UCase(Chr(Ascii(i - 1))))
                    i_Chr = Ascii(i - 1)
                    

                    'Check keyword characters
                    If i_Chr = 124 Then 'If the character is "|"
                    
                    Else
                        Select Case m_Chr
    
                            Case 65 To 90
                                .Map.Letters(m_Style, m_Chr - 65).MapPosition(i).x = m_Position.x + i - 1
                                .Map.Letters(m_Style, m_Chr - 65).MapPosition(i).y = m_Position.y
                            Case 33 To 64
                                .Map.Letter(m_Chr - 33).MapPosition(i).x = m_Position.x + i - 1
                                .Map.Letter(m_Chr - 33).MapPosition(i).y = m_Position.y
                            Case 91 To 96
                                .Map.Letter(m_Chr - 59).MapPosition(i).x = m_Position.x + i - 1
                                .Map.Letter(m_Chr - 59).MapPosition(i).y = m_Position.y
                            Case 123 To 126
    
                                .Map.Letter(m_Chr - 85).MapPosition(i).x = m_Position.x + i - 1
                                .Map.Letter(m_Chr - 85).MapPosition(i).y = m_Position.y
                        End Select
                    End If
                Next i
            End If
            
    End With
End Sub


 


