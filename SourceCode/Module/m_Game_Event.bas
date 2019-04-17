Attribute VB_Name = "m_Game_Event"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************

'====================================================================Function description====================================================

' This function describes the whole game event library.

'====================================================================================================================================



Option Explicit

'=============================================================
'Describe:Get Game Events
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Event_Get(m_Graphic As m_Graphic, m_EventID As Byte)
    
    Call Events_Activate(m_Graphic, m_EventID)

End Sub

'=============================================================
'Describe:Game Event Base
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Events_Activate(m_Graphic As m_Graphic, m_index As Byte)
Dim i As Byte

With m_Graphic
    Select Case m_index
        Case 0

            
        Case 1
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            

            For i = 1 To 1
            .Map.Windows(i).Visible = False
            .SpriteFont(i).Visiable = .Map.Windows(1).Visible
            
            Next i
            .SpriteFont(4).Visiable = False
            .SpriteFont(5).Visiable = False
            .Map.Windows(0).Visible = True
            .SpriteFont(0).Visiable = .Map.Windows(0).Visible
            Debug.Print .SpriteFont(0).G_String
        Case 2
            .Map.Windows(0).Visible = False
            .SpriteFont(0).Visiable = .Map.Windows(0).Visible
            .SpriteFont(4).Visiable = False
            .SpriteFont(5).Visiable = False
            .Map.Windows(1).Visible = True
            .SpriteFont(1).Visiable = .Map.Windows(1).Visible
            fMainForm.Txt_Input.Text = ""
            fMainForm.Txt_Input.Visible = True
        
        
        Case 3
             For i = 0 To 1
                .Map.Windows(i).Visible = False
                .SpriteFont(i).Visiable = .Map.Windows(i).Visible
            Next i
            .SpriteFont(4).Visiable = False
            .SpriteFont(5).Visiable = False
            .Map.Windows(2).Visible = True
            .SpriteFont(2).Visiable = .Map.Windows(2).Visible
            fMainForm.Txt_Input.Text = ""
            fMainForm.Txt_Input.Visible = False
            
        Case 6
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            

            For i = 1 To 1
            .Map.Windows(i).Visible = False
            .SpriteFont(i).Visiable = .Map.Windows(1).Visible
            
            Next i
            .SpriteFont(4).Visiable = True
            .SpriteFont(5).Visiable = False
            .Map.Windows(0).Visible = True
            .SpriteFont(0).Visiable = False
 
        Case 7
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            

            For i = 1 To 1
            .Map.Windows(i).Visible = False
            .SpriteFont(i).Visiable = .Map.Windows(1).Visible
            
            Next i
            .SpriteFont(4).Visiable = False
            .SpriteFont(5).Visiable = True
            .Map.Windows(0).Visible = True
            .SpriteFont(0).Visiable = False
            
        
        Case 40
            Call Emo(m_Graphic, 0)
        Case 41
            Call Emo(m_Graphic, 1)
        Case 42
            Call Emo(m_Graphic, 2)

        
        
        Case 50 '穿越打开文件
            
            
            If .Map.GameEvents(m_index).m_Description <> "" Then
                  Call CodeFileLoad(this_Graphic, .Map.GameEvents(m_index).m_Description & ".pl")
            End If
            
                    
        Case 51 '显示对话框，带输入框
            If fMainForm.Filelist.Visible = True Then fMainForm.Filelist.Visible = False
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            .Map.Windows(1).Visible = True
            fMainForm.Txt_Input.Left = .Map.Windows(1).DrawPosition.x + 70
            fMainForm.Txt_Input.Top = .Map.Windows(1).DrawPosition.y + 340
            fMainForm.Txt_Input.Visible = True
            fMainForm.Txt_Input.SetFocus
            fMainForm.Txt_Input.Text = ""
            
            
        Case 52 '显示对话框，不输入框
            If fMainForm.Filelist.Visible = True Then fMainForm.Filelist.Visible = False
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            .Map.Windows(1).Visible = True
            .SpriteFont(1).Visiable = True
            fMainForm.Txt_Input.Visible = False
            
            
        Case 61 '打开小文件

            'fMainForm.Filelist.Refresh
            .Map.Windows(1).Visible = True
            fMainForm.Filelist.Left = .Map.Windows(1).DrawPosition.x + 70
            fMainForm.Filelist.Top = .Map.Windows(1).DrawPosition.y + 140
            fMainForm.Filelist.Visible = True
            
        Case 62 '关闭小文件
            .Map.Windows(1).Visible = False
            If fMainForm.Filelist.Visible = True Then fMainForm.Filelist.Visible = False
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            
        Case 63
            For i = 0 To 1
                .Map.Windows(i).Visible = False
                .SpriteFont(i).Visiable = .Map.Windows(i).Visible
            Next i
            .SpriteFont(4).Visiable = False
            .SpriteFont(5).Visiable = False
            If fMainForm.Filelist.Visible = True Then fMainForm.Filelist.Visible = False
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
            If fMainForm.Txt_Input.Visible = True Then fMainForm.Txt_Input.Visible = False
        
    End Select
    

End With

    
End Sub

Private Sub Emo(m_Graphic As m_Graphic, m_index As Integer)
    Dim a, b As Integer
    With m_Graphic
    
    

        For b = 42 To 51
            For a = 0 To 31
                .Map.Letter(b).MapPosition(a).x = -1
                .Map.Letter(b).MapPosition(a).y = -1
            Next a
        Next b
        
                .Map.Letter(44).MapPosition(0).x = 8
                .Map.Letter(44).MapPosition(0).y = 2
                .Map.Letter(44).MapPosition(1).x = 7
                .Map.Letter(44).MapPosition(1).y = 2
                .Map.Letter(44).MapPosition(2).x = 9
                .Map.Letter(44).MapPosition(2).y = 2
                .Map.Letter(44).MapPosition(3).x = 6
                .Map.Letter(44).MapPosition(3).y = 2
                .Map.Letter(44).MapPosition(4).x = 10
                .Map.Letter(44).MapPosition(4).y = 2
                .Map.Letter(44).MapPosition(5).x = 5
                .Map.Letter(44).MapPosition(5).y = 3
                .Map.Letter(44).MapPosition(6).x = 11
                .Map.Letter(44).MapPosition(6).y = 3
                .Map.Letter(44).MapPosition(7).x = 4
                .Map.Letter(44).MapPosition(7).y = 4
                .Map.Letter(44).MapPosition(8).x = 12
                .Map.Letter(44).MapPosition(8).y = 4
                .Map.Letter(44).MapPosition(9).x = 3
                .Map.Letter(44).MapPosition(9).y = 5
                .Map.Letter(44).MapPosition(10).x = 13
                .Map.Letter(44).MapPosition(10).y = 5
                .Map.Letter(44).MapPosition(11).x = 3
                .Map.Letter(44).MapPosition(11).y = 6
                .Map.Letter(44).MapPosition(12).x = 13
                .Map.Letter(44).MapPosition(12).y = 6
                .Map.Letter(44).MapPosition(13).x = 3
                .Map.Letter(44).MapPosition(13).y = 7
                .Map.Letter(44).MapPosition(14).x = 13
                .Map.Letter(44).MapPosition(14).y = 7
                .Map.Letter(44).MapPosition(15).x = 3
                .Map.Letter(44).MapPosition(15).y = 8
                .Map.Letter(44).MapPosition(16).x = 13
                .Map.Letter(44).MapPosition(16).y = 8
                .Map.Letter(44).MapPosition(17).x = 3
                .Map.Letter(44).MapPosition(17).y = 9
                .Map.Letter(44).MapPosition(18).x = 13
                .Map.Letter(44).MapPosition(18).y = 9
                .Map.Letter(44).MapPosition(19).x = 4
                .Map.Letter(44).MapPosition(19).y = 10
                .Map.Letter(44).MapPosition(20).x = 12
                .Map.Letter(44).MapPosition(20).y = 10
                .Map.Letter(44).MapPosition(21).x = 5
                .Map.Letter(44).MapPosition(21).y = 11
                .Map.Letter(44).MapPosition(22).x = 11
                .Map.Letter(44).MapPosition(22).y = 11
                .Map.Letter(44).MapPosition(23).x = 6
                .Map.Letter(44).MapPosition(23).y = 12
                .Map.Letter(44).MapPosition(24).x = 10
                .Map.Letter(44).MapPosition(24).y = 12
                .Map.Letter(44).MapPosition(25).x = 7
                .Map.Letter(44).MapPosition(25).y = 12
                .Map.Letter(44).MapPosition(26).x = 9
                .Map.Letter(44).MapPosition(26).y = 12
                .Map.Letter(44).MapPosition(27).x = 8
                .Map.Letter(44).MapPosition(27).y = 12
                
        Select Case m_index
        
        
            Case 0

                .Map.Letter(49).MapPosition(0).x = 7
                .Map.Letter(49).MapPosition(0).y = 5
                .Map.Letter(49).MapPosition(1).x = 9
                .Map.Letter(49).MapPosition(1).y = 5
                .Map.Letter(49).MapPosition(2).x = 7
                .Map.Letter(49).MapPosition(2).y = 6
                .Map.Letter(49).MapPosition(3).x = 9
                .Map.Letter(49).MapPosition(3).y = 6
                .Map.Letter(49).MapPosition(4).x = 5
                .Map.Letter(49).MapPosition(4).y = 7
                .Map.Letter(49).MapPosition(5).x = 11
                .Map.Letter(49).MapPosition(5).y = 7
                .Map.Letter(49).MapPosition(6).x = 5
                .Map.Letter(49).MapPosition(6).y = 8
                .Map.Letter(49).MapPosition(7).x = 11
                .Map.Letter(49).MapPosition(7).y = 8
                .Map.Letter(49).MapPosition(8).x = 6
                .Map.Letter(49).MapPosition(8).y = 9
                .Map.Letter(49).MapPosition(9).x = 10
                .Map.Letter(49).MapPosition(9).y = 9
                .Map.Letter(49).MapPosition(10).x = 7
                .Map.Letter(49).MapPosition(10).y = 10
                .Map.Letter(49).MapPosition(11).x = 9
                .Map.Letter(49).MapPosition(11).y = 10
                .Map.Letter(49).MapPosition(12).x = 8
                .Map.Letter(49).MapPosition(12).y = 10
            Case 1
                .Map.Letter(50).MapPosition(0).x = 6
                .Map.Letter(50).MapPosition(0).y = 6
                .Map.Letter(50).MapPosition(1).x = 10
                .Map.Letter(50).MapPosition(1).y = 6
                .Map.Letter(50).MapPosition(2).x = 7
                .Map.Letter(50).MapPosition(2).y = 6
                .Map.Letter(50).MapPosition(3).x = 9
                .Map.Letter(50).MapPosition(3).y = 6
                .Map.Letter(50).MapPosition(4).x = 5
                .Map.Letter(50).MapPosition(4).y = 8
                .Map.Letter(50).MapPosition(5).x = 11
                .Map.Letter(50).MapPosition(5).y = 8
                .Map.Letter(50).MapPosition(6).x = 6
                .Map.Letter(50).MapPosition(6).y = 8
                .Map.Letter(50).MapPosition(7).x = 10
                .Map.Letter(50).MapPosition(7).y = 8
                .Map.Letter(50).MapPosition(8).x = 7
                .Map.Letter(50).MapPosition(8).y = 9
                .Map.Letter(50).MapPosition(9).x = 9
                .Map.Letter(50).MapPosition(9).y = 9
                .Map.Letter(50).MapPosition(10).x = 8
                .Map.Letter(50).MapPosition(10).y = 9
            Case 2
                .Map.Letter(50).MapPosition(0).x = 7
                .Map.Letter(50).MapPosition(0).y = 5
                .Map.Letter(50).MapPosition(1).x = 9
                .Map.Letter(50).MapPosition(1).y = 5
                .Map.Letter(50).MapPosition(2).x = 6
                .Map.Letter(50).MapPosition(2).y = 6
                .Map.Letter(50).MapPosition(3).x = 10
                .Map.Letter(50).MapPosition(3).y = 6
                .Map.Letter(50).MapPosition(4).x = 5
                .Map.Letter(50).MapPosition(4).y = 9
                .Map.Letter(50).MapPosition(5).x = 11
                .Map.Letter(50).MapPosition(5).y = 9
                .Map.Letter(50).MapPosition(6).x = 6
                .Map.Letter(50).MapPosition(6).y = 8
                .Map.Letter(50).MapPosition(7).x = 10
                .Map.Letter(50).MapPosition(7).y = 8
                .Map.Letter(50).MapPosition(8).x = 7
                .Map.Letter(50).MapPosition(8).y = 8
                .Map.Letter(50).MapPosition(9).x = 9
                .Map.Letter(50).MapPosition(9).y = 8
                .Map.Letter(50).MapPosition(10).x = 8
                .Map.Letter(50).MapPosition(10).y = 8

                
        End Select
    End With
End Sub
