Attribute VB_Name = "m_Games_Graphic"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

'Function Library of game animation

'====================================================================================================================================



Option Explicit


'=============================================================
'Describe:Game initialization, mainly dealing with the graphics part of the game
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Game_Initialize(this_hdc As Long, this_FilePath As m_FilePath)     '初始化窗体
    Call Game_MapDataLoad(this_FilePath)
    EventBreak = False
 
    
    With this_Graphic

        
        .Buffer.BackBuffer = CreateCompatibleDC(this_hdc)
        .Buffer.BackBufferBmp = CreateCompatibleBitmap(this_hdc, .Screen.Width, .Screen.Height)
        .Buffer.OldBackBufferDC = SelectObject(.Buffer.BackBuffer, .Buffer.BackBufferBmp)
        
        .Buffer.TileSetBmp(0) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(0) = SelectObject(.Buffer.TileSetBmp(0), LoadPicture(this_FilePath.Graphics & .GraphicFiles(0).FilesName))
        
        .Buffer.TileSetBmp(1) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(1) = SelectObject(.Buffer.TileSetBmp(1), LoadPicture(this_FilePath.Graphics & .GraphicFiles(1).FilesName))
        
        .Buffer.TileSetBmp(2) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(2) = SelectObject(.Buffer.TileSetBmp(2), LoadPicture(this_FilePath.Graphics & .GraphicFiles(2).FilesName))
        
        .Buffer.TileSetBmp(3) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(3) = SelectObject(.Buffer.TileSetBmp(3), LoadPicture(this_FilePath.Graphics & .GraphicFiles(3).FilesName))
        
        .Buffer.TileSetBmp(4) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(4) = SelectObject(.Buffer.TileSetBmp(4), LoadPicture(this_FilePath.Graphics & .GraphicFiles(4).FilesName))
        
        .Buffer.TileSetBmp(5) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(5) = SelectObject(.Buffer.TileSetBmp(5), LoadPicture(this_FilePath.Graphics & .GraphicFiles(5).FilesName))
        
        .Buffer.TileSetBmp(6) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(6) = SelectObject(.Buffer.TileSetBmp(6), LoadPicture(this_FilePath.Graphics & .GraphicFiles(6).FilesName))
        
        .Buffer.TileSetBmp(7) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(7) = SelectObject(.Buffer.TileSetBmp(7), LoadPicture(this_FilePath.Graphics & .GraphicFiles(7).FilesName))
        
        
        .Buffer.TileSetBmp(8) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(8) = SelectObject(.Buffer.TileSetBmp(8), fMainForm.img_Font.Picture)
        
        .Buffer.TileSetBmp(9) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(9) = SelectObject(.Buffer.TileSetBmp(9), fMainForm.img_Event.Picture)
        
        .Buffer.TileSetBmp(10) = CreateCompatibleDC(this_hdc)
        .Buffer.OldTilesetBmpDC(10) = SelectObject(.Buffer.TileSetBmp(10), LoadPicture(this_FilePath.Graphics & .GraphicFiles(10).FilesName))
        
        
        
    End With
    
    
End Sub


'=============================================================
'Describe:Main Functions of Graphic Drawing
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub GameDraw(this_hdc As Long, this_Graphic As m_Graphic, this_Switch As m_Switch)
    Call Draw_Tiles1(this_Graphic, this_Switch)
    Call Draw_Tiles2(this_Graphic, this_Switch)
    Call Draw_Tiles_Font(this_Graphic) '绘制基于窗体的文字
    
    Call Draw_Tiles3(this_Graphic, this_Switch)
    Call Draw_Tiles4(this_Graphic, this_Switch) '  特效层
    Call Draw_Tiles5(this_Graphic, this_Switch)
    Call Draw_Tiles9(this_Graphic)
    Call Events(this_Graphic, this_Switch)
    Call Draw_Tiles_Font(this_Graphic, True)
    Call Draw_DisplayDigital_Timer(this_Graphic, 790, 40, this_Switch)
    
    
    BitBlt this_hdc, this_Graphic.Posi.x, this_Graphic.Posi.y, this_Graphic.Screen.Width, this_Graphic.Screen.Height, this_Graphic.Buffer.BackBuffer, 0, 0, vbSrcCopy    '将地图绘制在屏幕上

End Sub

'=============================================================
'Describe:Draw the tiles of the game
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles1(this_Graphic As m_Graphic, m_Switch As m_Switch)   '
    Dim i, j, II, jj, a As Integer
    Dim f As Byte
    Dim s As Integer
    s = -14
    With this_Graphic
        .Player(0).Pic.x = .Player(0).Pic_Temp(0).x
        .Player(0).Pic.y = .Player(0).Pic_Temp(0).y
                                
        BitBlt .Buffer.BackBuffer, 0, 0, .Screen.Width, .Screen.Height, .Buffer.TileSetBmp(0), 0, 0, vbSrcCopy
            If .Map.Tiles(.Map.TileID).GraphicPosition.Width = 0 Then .Map.Tiles(.Map.TileID).GraphicPosition.Width = 1
            If .Map.Tiles(.Map.TileID).GraphicPosition.Height = 0 Then .Map.Tiles(.Map.TileID).GraphicPosition.Height = 1
            For i = 0 To (.Map.TilesInfo.Width / .Map.Tiles(.Map.TileID).GraphicPosition.Width) - 1
                For j = 0 To (.Map.TilesInfo.Height / .Map.Tiles(.Map.TileID).GraphicPosition.Height) - 1
                   BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tiles(.Map.TileID).GraphicPosition.Width, .Map.TilesInfo.y + j * .Map.Tiles(.Map.TileID).GraphicPosition.Height, .Map.Tiles(.Map.TileID).GraphicPosition.Width, .Map.Tiles(.Map.TileID).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Tiles(.Map.TileID).GraphicID), .Map.Tiles(.Map.TileID).GraphicPosition.x, .Map.Tiles(.Map.TileID).GraphicPosition.y, vbSrcCopy
                Next j
            Next i
            
            If m_Switch.Pathway = True Then
                For i = 0 To 15
                    For j = 0 To 13
                        If .Map.BlockTiles(i, j) = 1 Then
                            '调整地砖方向的
                            f = 6
                            jj = j
                            II = i
                            If jj = 0 Then jj = 1
                            If II = 0 Then II = 1
                            
                            Select Case .Map.BlockTiles(II, jj - 1)
                                Case 1
                                    Select Case .Map.BlockTiles(II - 1, jj)
                                        Case 1
                                            Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 4
                                                        Case 0
                                                            f = 7
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            If i = 0 Then
                                                                f = 3
                                                            Else
                                                                f = 5
                                                            End If
                                                        Case 0
                                                            f = 8
                                                    End Select
                                            End Select
                                        Case 0
                                              Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                            End Select
                                        End Select
                                Case 0
                                     Select Case .Map.BlockTiles(II - 1, jj)
                                        Case 1
                                            Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 1
                                                        Case 0
                                                            f = 7
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 2
                                                        Case 0
                                                            f = 8
                                                    End Select
                                            End Select
                                        Case 0
                                              Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 0
                                                        Case 0
                                                            f = 6
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                            End Select
                                            
                                        End Select
                            End Select
                            
                            If j = 13 Then
                                If i = 0 Then
                                    f = 6
                                Else
                                    If .Map.BlockTiles(i, 12) = 1 Then
                                        If .Map.BlockTiles(i - 1, 12) = 0 Then
                                            If .Map.BlockTiles(i + 1, 12) = 1 Then
                                                f = 6
                                            Else
                                                f = 5
                                            End If
                                        Else
                                            If .Map.BlockTiles(i + 1, 12) = 1 Then
                                                f = 7
                                            Else
                                                f = 8
                                            End If
                                        End If
                                    Else
                                        f = 7
                                    End If
                                End If
                            End If
                            
                            If m_Switch.PathwayWithBlock = True Then
                                .Map.Block(i, j) = 1
                            Else
                                .Map.Block(i, j) = 0
                            End If
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.TilesInfo.y - 10 + j * .Map.BlockTilesInfo(f).GraphicPosition.Height, .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.BlockTilesInfo(f).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.BlockTilesInfo(f).GraphicID), .Map.BlockTilesInfo(f).GraphicPosition.x, .Map.BlockTilesInfo(f).GraphicPosition.y, vbSrcCopy
                            '自动填充地砖的下方
                            If f > 5 And f < 9 Then
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.TilesInfo.y - 10 + (j + 1) * .Map.BlockTilesInfo(f).GraphicPosition.Height, .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.BlockTilesInfo(f + 3).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.BlockTilesInfo(f).GraphicID), .Map.BlockTilesInfo(f + 3).GraphicPosition.x, .Map.BlockTilesInfo(f + 3).GraphicPosition.y, vbSrcCopy
                            End If

                            
                            If .Player(CurrentPlayerID).Info.C_Position.x = i And .Player(CurrentPlayerID).Info.C_Position.y = j Then
                                .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                            End If
                            
                        End If


                            
                    Next j
                Next i
            End If
            
            If m_Switch.Letters = True Then
                For i = 0 To 3
                    For j = 0 To 25
                        For a = 0 To 31
                            If .Map.Letters(i, j).MapPosition(a).x = -1 And .Map.Letters(i, j).MapPosition(a).y = -1 Then
                            
                            Else
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Letters(i, j).MapPosition(a).x * .Map.Tile_Object.Width, .Map.TilesInfo.y - 10 + .Map.Letters(i, j).MapPosition(a).y * .Map.Tile_Object.Height, .Map.Letters(i, j).GraphicPosition.Width, .Map.Letters(i, j).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Letters(i, j).GraphicID), .Map.Letters(i, j).GraphicPosition.x, .Map.Letters(i, j).GraphicPosition.y, vbSrcCopy
                            
                            
                                If .Player(CurrentPlayerID).Info.C_Position.x = .Map.Letters(i, j).MapPosition(a).x And .Player(CurrentPlayerID).Info.C_Position.y = .Map.Letters(i, j).MapPosition(a).y Then
                                    .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                    .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                                End If
                            End If
                        Next a
                    Next j
                Next i
                       
                For i = 0 To 51
                    For a = 0 To 31
                        If .Map.Letter(i).MapPosition(a).x = -1 And .Map.Letter(i).MapPosition(a).y = -1 Then
                        
                        Else
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Letter(i).MapPosition(a).x * .Map.Tile_Object.Width, .Map.TilesInfo.y - 10 + .Map.Letter(i).MapPosition(a).y * .Map.Tile_Object.Height, .Map.Letter(i).GraphicPosition.Width, .Map.Letter(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Letter(i).GraphicID), .Map.Letter(i).GraphicPosition.x, .Map.Letter(i).GraphicPosition.y, vbSrcCopy
                        
                        
                            If .Player(CurrentPlayerID).Info.C_Position.x = .Map.Letter(i).MapPosition(a).x And .Player(CurrentPlayerID).Info.C_Position.y = .Map.Letter(i).MapPosition(a).y Then
                                .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                            End If
                        End If
                    Next a
                Next i
                
            End If
            
            
            
           If m_Switch.Steps = True Then
               
                For i = 1 To 99
                    If Steps(i).Visible = True Then
                        If Steps(i - 1).Position.y <> -1 And Steps(i - 1).Position.x <> -1 Then
                        
                            Select Case Steps(i).Direction
                                Case 0 '向下
                                    Steps(i).Position.y = Steps(i - 1).Position.y + 1
                                    Steps(i).Position.x = Steps(i - 1).Position.x
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                    If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                        s = -24
                                    Else
                                        s = -14
                                    End If
                                    End If
                                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                                Case 1 '向左
                                    Steps(i).Position.y = Steps(i - 1).Position.y
                                    Steps(i).Position.x = Steps(i - 1).Position.x - 1
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                               
                                Case 2 '向右
                            
                                    Steps(i).Position.y = Steps(i - 1).Position.y
                                    Steps(i).Position.x = Steps(i - 1).Position.x + 1
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                                
                                Case 3 ' 向上
                                    Steps(i).Position.y = Steps(i - 1).Position.y - 1
                                    Steps(i).Position.x = Steps(i - 1).Position.x
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                            
                            End Select
                    
                        End If
                    End If
                Next i
           
           End If
    End With
    
    
    
End Sub

'=============================================================
'Describe:Draw the object layer of the game
'Author:  Chenlyu
'Parameter:
'=============================================================

Public Sub Draw_Tiles2(this_Graphic As m_Graphic, m_Switch As m_Switch) '
    Dim i, j As Integer
    Dim M As m_GraphicPosition
    Dim s As Integer
    Dim Sx As Integer
    s = 20
    
    With this_Graphic
        If m_Switch.Object = True Then
            For i = 0 To 31
                For j = 0 To 31
                    Sx = 10
                    M.x = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.x
                    M.y = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.y
                    M.Height = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.Height
                    M.Width = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.Width
            
                    If .Map.Objects(i).MapPosition(j).x = -1 And .Map.Objects(i).MapPosition(j).y = -1 Then
'                        If i = 15 Then
'                             Debug.Print "1"
'                        End If
                    Else
'                        If i = 15 Then
'                             Debug.Print "1"
'                        End If
                        
                        If .Map.Objects(i).LayerNum = 0 Then
'                            If .Map.Objects(i).GraphicID = 1 Then
'                                Debug.Print "1"
'                            End If
                            If .Map.BlockTiles(.Map.Objects(i).MapPosition(j).x, .Map.Objects(i).MapPosition(j).y) = 1 Then
                                s = 20
                            Else
                                s = 10
                            End If
                            If i <= 9 Then Sx = 0
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Objects(i).MapPosition(j).x * .Map.Tile_Object.Width - Sx, .Map.TilesInfo.y + .Map.Objects(i).MapPosition(j).y * .Map.Tile_Object.Height - s, .Map.Objects(i).GraphicPosition.Width, .Map.Objects(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Objects(i).GraphicID), .Map.Objects(i).GraphicPosition.x + M.x * M.Width, .Map.Objects(i).GraphicPosition.y + M.y * M.Height, vbSrcAnd
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Objects(i).MapPosition(j).x * .Map.Tile_Object.Width - Sx, .Map.TilesInfo.y + .Map.Objects(i).MapPosition(j).y * .Map.Tile_Object.Height - s, .Map.Objects(i).GraphicPosition.Width, .Map.Objects(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Objects(i).GraphicID), .Map.Objects(i).GraphicPosition.x, .Map.Objects(i).GraphicPosition.y, vbSrcPaint
                        End If
                        
                        Select Case i
                        
                            Case 9
                                .Map.Block(.Map.Objects(i).MapPosition(j).x, .Map.Objects(i).MapPosition(j).y) = 1
                                .Map.Block(.Map.Objects(i).MapPosition(j).x + 1, .Map.Objects(i).MapPosition(j).y) = 1
                                .Map.Block(.Map.Objects(i).MapPosition(j).x + 2, .Map.Objects(i).MapPosition(j).y) = 1
                            Case 25
                                 If (.Map.Objects(i).MapPosition(j).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(i).MapPosition(j).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(i).MapPosition(j).y = .Player(CurrentPlayerID).Info.C_Position.y - 1 Then
                                    .Map.Objects(i).MapPosition(j).x = -1
                                    .Map.Objects(i).MapPosition(j).y = -1
                                 End If
                            Case 30
                                 If (.Map.Objects(i).MapPosition(j).x = .Player(CurrentPlayerID).Info.C_Position.x Or .Map.Objects(i).MapPosition(j).x = .Player(CurrentPlayerID).Info.C_Position.x - 1) And .Map.Objects(i).MapPosition(j).y = .Player(CurrentPlayerID).Info.C_Position.y - 1 Then
                                    .Map.Objects(i).MapPosition(j).x = -1
                                    .Map.Objects(i).MapPosition(j).y = -1
                                 End If
                                 
                            Case Else
                            
                        End Select
                    End If
                Next j
            Next i
        End If
    End With
End Sub


'=============================================================
'Describe:Draw the object layer of the game This layer mainly draws objects for occlusion.
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles5(this_Graphic As m_Graphic, m_Switch As m_Switch) '
    Dim i, j As Integer
    Dim M As m_GraphicPosition
    Dim Sx As Integer
    Dim s As Integer
    
    s = 20
    
    With this_Graphic
        If m_Switch.Object = True Then
            For i = 0 To 31
                For j = 0 To 31
                    Sx = 10
                    M.x = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.x
                    M.y = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.y
                    M.Height = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.Height
                    M.Width = .GraphicFiles(.Map.Objects(i).GraphicID).BlackPicPixel.Width
                    
                    
                    If .Map.Objects(i).LayerNum = 1 And .Map.Objects(i + 1).LayerNum = 0 Then
                        
                        .Map.Objects(i).MapPosition(j).x = .Map.Objects(i + 1).MapPosition(j).x
                        .Map.Objects(i).MapPosition(j).y = .Map.Objects(i + 1).MapPosition(j).y
                        
                        If .Map.Objects(i).MapPosition(j).x = -1 And .Map.Objects(i).MapPosition(j).y = -1 Then
                        Else
                            If .Map.BlockTiles(.Map.Objects(i).MapPosition(j).x, .Map.Objects(i).MapPosition(j).y) = 1 Then
                                s = 20
                            Else
                                s = 10
                            End If
                            If i <= 9 Then Sx = 0
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Objects(i).MapPosition(j).x * .Map.Tile_Object.Width - Sx, .Map.TilesInfo.y + .Map.Objects(i).MapPosition(j).y * .Map.Tile_Object.Height - s, .Map.Objects(i).GraphicPosition.Width, .Map.Objects(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Objects(i).GraphicID), .Map.Objects(i).GraphicPosition.x + M.x * M.Width, .Map.Objects(i).GraphicPosition.y + M.y * M.Height, vbSrcAnd
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Objects(i).MapPosition(j).x * .Map.Tile_Object.Width - Sx, .Map.TilesInfo.y + .Map.Objects(i).MapPosition(j).y * .Map.Tile_Object.Height - s, .Map.Objects(i).GraphicPosition.Width, .Map.Objects(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Objects(i).GraphicID), .Map.Objects(i).GraphicPosition.x, .Map.Objects(i).GraphicPosition.y, vbSrcPaint
                            
                        End If
                    End If
                Next j
            Next i
        End If
    End With
End Sub

'=============================================================
'Describe:Character Layer of the Game
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles3(this_Graphic As m_Graphic, m_Switch As m_Switch) '
    Dim i, j As Integer
    Dim a As Integer
    a = 6
    Dim m_PlayerPosition As m_Position
    Dim M As m_GraphicPosition
    
    With this_Graphic
        If m_Switch.Player = True Then
        For i = 0 To 1
            M.x = .GraphicFiles(.Player(i).GraphicID).BlackPicPixel.x
            M.y = .GraphicFiles(.Player(i).GraphicID).BlackPicPixel.y
            M.Height = .GraphicFiles(.Player(i).GraphicID).BlackPicPixel.Height
            M.Width = .GraphicFiles(.Player(i).GraphicID).BlackPicPixel.Width
                    
            If .Player(i).Info.Alive = True Then
                
                If .Player(i).Info.EventTimer = 0 Or .Player(i).Info.EventBreak = True Then
                    .Player(i).Info.MoveTimer = 0
                    BitBlt .Buffer.BackBuffer, .Player(i).Info.C_Position.x * .Map.Tile_Object.Width + .Player(i).Pic.x, .Player(i).Info.C_Position.y * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer + M.x * M.Width, .Player(i).Info.MoveDirection * .Player(i).Tile.Height + M.y * M.Height, vbSrcAnd
                    BitBlt .Buffer.BackBuffer, .Player(i).Info.C_Position.x * .Map.Tile_Object.Width + .Player(i).Pic.x, .Player(i).Info.C_Position.y * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer, .Player(i).Info.MoveDirection * .Player(i).Tile.Height, vbSrcPaint
                If .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y) <> 0 And CurrentGameEvents = 63 Then
                    CurrentGameEvents = .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y)
                    Call Event_Get(this_Graphic, CurrentGameEvents)
                End If
                                        
            Else
                    If CurrentGameEvents = 63 Then CurrentGameEvents = 0
                    If .Player(i).Info.G_Event(CurrentEventTimer) < 4 Then
                        m_PlayerPosition.x = 0
                        m_PlayerPosition.y = 0
                        
                        If .Player(i).Info.C_Position.x > 0 Then m_PlayerPosition.x = .Player(i).Info.C_Position.x - 1
                        If .Player(i).Info.C_Position.y > 0 Then m_PlayerPosition.y = .Player(i).Info.C_Position.y - 1
                        
'                        CurrentGameEvents = 63
'                        Call Event_Get(this_Graphic, CurrentGameEvents)

'                        If .Map.Events(.Player(i).Info.C_Position.X + 1, .Player(i).Info.C_Position.Y) = 0 And .Map.Events(m_PlayerPosition.X, .Player(i).Info.C_Position.Y) = 0 And .Map.Events(.Player(i).Info.C_Position.X, .Player(i).Info.C_Position.Y + 1) = 0 And .Map.Events(.Player(i).Info.C_Position.X, m_PlayerPosition.Y) = 0 Then
'
'                            If CurrentGameEvents <> 63 Then
'                                CurrentGameEvents = 63
'                                Call Event_Get(this_Graphic, CurrentGameEvents)
'                            End If
'                        End If

                        

                        
                        .Player(i).Info.MoveDirection = .Player(i).Info.G_Event(CurrentEventTimer)
                        Select Case .Player(i).Info.MoveDirection
                        
                            Case MoveRight
                                If .Player(i).Info.PositionTimer < 1 Then
                                        If .Map.Events(.Player(i).Info.C_Position.x + 1, .Player(i).Info.C_Position.y) > 0 And .Map.Events(.Player(i).Info.C_Position.x + 1, .Player(i).Info.C_Position.y) <> CurrentGameEvents Then
                                            CurrentGameEvents = .Map.Events(.Player(i).Info.C_Position.x + 1, .Player(i).Info.C_Position.y)
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        ElseIf .Map.Events(.Player(i).Info.C_Position.x + 1, .Player(i).Info.C_Position.y) = 0 Then
                                            CurrentGameEvents = 63
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        End If
                                        
                                        '碰到墙面后的状态
                                        If .Player(i).Info.C_Position.x < ((.Map.TilesInfo.Width - 42) / .Map.Tile_Object.Width) And .Map.Block(.Player(i).Info.C_Position.x + 1, .Player(i).Info.C_Position.y) = False Then
                                            .Player(i).Info.PositionTimer = .Player(i).Info.PositionTimer + .Player(i).Info.MoveSpeed
                                        Else
                                            If EventBreak = True Then
                                                .Player(i).Info.MoveTimer = 0
                                                .Player(i).Info.PositionTimer = 0
                                            Else
                                                If CurrentEventTimer < 1022 And CurrentEventTimer < .Player(i).Info.EventTimer Then
                                                    CurrentEventTimer = CurrentEventTimer + 1
                                                    If CurrentEventTimer = .Player(i).Info.EventTimer Then
                                                        CurrentEventTimer = CurrentEventTimer
                                                        .Player(i).Info.EventBreak = True
                                                    End If
                                                End If
                                                    
                                            End If
                                        End If
                                Else
                                    Call Player_Event_MinusMove(.Player(i).Info.MoveDirection, CByte(i))
                                    .Player(i).Info.C_Position.x = .Player(i).Info.C_Position.x + 1
                                    .Player(i).Info.PositionTimer = 0
                                End If
                            
                            Case MoveDown
                                If .Player(i).Info.PositionTimer < 1 Then
                                    If .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y + 1) > 0 And .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y + 1) <> CurrentGameEvents Then
                                        CurrentGameEvents = .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y + 1)
                                        Call Event_Get(this_Graphic, CurrentGameEvents)
                                    ElseIf .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y + 1) = 0 Then
                                        CurrentGameEvents = 63
                                        Call Event_Get(this_Graphic, CurrentGameEvents)
                                    End If
                                        
                                    If .Player(i).Info.C_Position.y < ((.Map.TilesInfo.Height - 42) / .Map.Tile_Object.Height) And .Map.Block(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y + 1) = False Then
                                        .Player(i).Info.PositionTimer = .Player(i).Info.PositionTimer + .Player(i).Info.MoveSpeed
                                    Else
                                        If EventBreak = True Then
                                            .Player(i).Info.MoveTimer = 0
                                            .Player(i).Info.PositionTimer = 0
                                        Else
                                                If CurrentEventTimer < 1022 And CurrentEventTimer < .Player(i).Info.EventTimer Then
                                                    CurrentEventTimer = CurrentEventTimer + 1
                                                    If CurrentEventTimer = .Player(i).Info.EventTimer Then
                                                        CurrentEventTimer = CurrentEventTimer
                                                        .Player(i).Info.EventBreak = True
                                                    End If
                                                End If
                                        End If
                                    End If
                                Else
                                    Call Player_Event_MinusMove(.Player(i).Info.MoveDirection, CByte(i))
                                    .Player(i).Info.C_Position.y = .Player(i).Info.C_Position.y + 1
                                    .Player(i).Info.PositionTimer = 0
                                End If
                            Case MoveLeft
                                 If .Player(i).Info.PositionTimer > -1 Then
                                    If .Player(i).Info.C_Position.x - 1 >= 0 Then
                                        If .Map.Events(.Player(i).Info.C_Position.x - 1, .Player(i).Info.C_Position.y) > 0 And .Map.Events(.Player(i).Info.C_Position.x - 1, .Player(i).Info.C_Position.y) <> CurrentGameEvents Then
                                            CurrentGameEvents = .Map.Events(.Player(i).Info.C_Position.x - 1, .Player(i).Info.C_Position.y)
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        ElseIf .Map.Events(.Player(i).Info.C_Position.x - 1, .Player(i).Info.C_Position.y) = 0 Then
                                            CurrentGameEvents = 63
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        End If
                                    End If
                                    
                                    If .Player(i).Info.C_Position.x > ((.Map.TilesInfo.x) / .Map.Tile_Object.Width) Then
                                        If .Map.Block(.Player(i).Info.C_Position.x - 1, .Player(i).Info.C_Position.y) = False Then
                                            .Player(i).Info.PositionTimer = .Player(i).Info.PositionTimer - .Player(i).Info.MoveSpeed
                                        Else
                                            GoTo MoveLeft1:
                                        End If
                                    Else
MoveLeft1:
                                        If EventBreak = True Then
                                            .Player(i).Info.MoveTimer = 0
                                            .Player(i).Info.PositionTimer = 0
                                        Else
                                            If CurrentEventTimer < 1022 And CurrentEventTimer < .Player(i).Info.EventTimer Then
                                                    CurrentEventTimer = CurrentEventTimer + 1
                                                    If CurrentEventTimer = .Player(i).Info.EventTimer Then
                                                        CurrentEventTimer = CurrentEventTimer
                                                        .Player(i).Info.EventBreak = True
                                                    End If
                                            End If
                                        End If
                                    End If
                                Else
                                    Call Player_Event_MinusMove(.Player(i).Info.MoveDirection, CByte(i))
                                    .Player(i).Info.C_Position.x = .Player(i).Info.C_Position.x - 1
                                    .Player(i).Info.PositionTimer = 0

                                End If
                            
                            Case MoveUp
                                    
                                If .Player(i).Info.PositionTimer > -1 Then
                                    If .Player(i).Info.C_Position.y - 1 >= 0 Then
                                        If .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y - 1) > 0 And .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y - 1) <> CurrentGameEvents Then
                                            CurrentGameEvents = .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y - 1)
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        ElseIf .Map.Events(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y - 1) = 0 Then
                                            CurrentGameEvents = 63
                                            Call Event_Get(this_Graphic, CurrentGameEvents)
                                        End If
                                    End If
                                    If .Player(i).Info.C_Position.y > ((.Map.TilesInfo.y - 84) / .Map.Tile_Object.Height) Then
                                        If .Map.Block(.Player(i).Info.C_Position.x, .Player(i).Info.C_Position.y - 1) = False Then
                                            .Player(i).Info.PositionTimer = .Player(i).Info.PositionTimer - .Player(i).Info.MoveSpeed
                                        Else
                                            GoTo MoveUp1
                                        End If
                                    Else
MoveUp1:
                                        If EventBreak = True Then
                                            .Player(i).Info.MoveTimer = 0
                                            .Player(i).Info.PositionTimer = 0
                                        Else
                                            If CurrentEventTimer < 1022 And CurrentEventTimer < .Player(i).Info.EventTimer Then
                                                CurrentEventTimer = CurrentEventTimer + 1
                                                If CurrentEventTimer = .Player(i).Info.EventTimer Then
                                                    CurrentEventTimer = CurrentEventTimer
                                                    .Player(i).Info.EventBreak = True
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    Call Player_Event_MinusMove(.Player(i).Info.MoveDirection, CByte(i))
                                    .Player(i).Info.C_Position.y = .Player(i).Info.C_Position.y - 1
                                    .Player(i).Info.PositionTimer = 0
                        
                                End If
                        
                        End Select
                    End If
                    
                    If .Player(i).Info.PositionTimer = 0 Then
                        If this_Graphic.Player(CurrentPlayerID).Info.RoadXY = True Then Call Player_Move_Test(this_Graphic, CByte(i))
                    End If
                    
                    If .Player(i).Info.MagicBall = False Then
                        '判断人物的帧动画
                        If .Player(i).Info.MoveTimer < 3 Then
                            .Player(i).Info.MoveTimer = .Player(i).Info.MoveTimer + 1
                        Else
                           .Player(i).Info.MoveTimer = 0
                        End If
                        
                        Select Case .Player(i).Info.MoveDirection
                            Case MoveRight, MoveLeft
                                BitBlt .Buffer.BackBuffer, (.Player(i).Info.C_Position.x + .Player(i).Info.PositionTimer) * .Map.Tile_Object.Width + .Player(i).Pic.x, .Player(i).Info.C_Position.y * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer + M.x * M.Width, .Player(i).Info.MoveDirection * .Player(i).Tile.Height + M.y * M.Height, vbSrcAnd
                                BitBlt .Buffer.BackBuffer, (.Player(i).Info.C_Position.x + .Player(i).Info.PositionTimer) * .Map.Tile_Object.Width + .Player(i).Pic.x, .Player(i).Info.C_Position.y * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer, .Player(i).Info.MoveDirection * .Player(i).Tile.Height, vbSrcPaint
                            Case MoveDown, MoveUp
                                BitBlt .Buffer.BackBuffer, (.Player(i).Info.C_Position.x) * .Map.Tile_Object.Width + .Player(i).Pic.x, (.Player(i).Info.C_Position.y + .Player(i).Info.PositionTimer) * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer + M.x * M.Width, .Player(i).Info.MoveDirection * .Player(i).Tile.Height + M.y * M.Height, vbSrcAnd
                                BitBlt .Buffer.BackBuffer, (.Player(i).Info.C_Position.x) * .Map.Tile_Object.Width + .Player(i).Pic.x, (.Player(i).Info.C_Position.y + .Player(i).Info.PositionTimer) * .Map.Tile_Object.Height + .Player(i).Pic.y, .Player(i).Tile.Width, .Player(i).Tile.Height, .Buffer.TileSetBmp(.Player(i).GraphicID), .Player(i).Tile.Width * .Player(i).Info.MoveTimer, .Player(i).Info.MoveDirection * .Player(i).Tile.Height, vbSrcPaint
                         End Select

                    Else

                    
                    End If
                End If
                
                  
            End If
             
        Next i
        End If
    End With
End Sub

'=============================================================
'Describe:Major effects layers for the game
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles4(this_Graphic As m_Graphic, m_Switch As m_Switch) '
    Dim i, j As Integer
    Dim M As m_GraphicPosition
    Dim a, b, c As Byte
    Dim k, h As Byte
    Dim s(2) As Integer
    
    s(0) = 100 - 42
    s(1) = 100 - 42
    With this_Graphic
        If m_Switch.Effect = True Then
       
   
            For i = 0 To 7
                Select Case i
                    Case 0
                        s(0) = 100 - 42
                        s(1) = 100 - 52
                    Case 1
                        s(0) = 100 - 52
                        s(1) = 100 - 20
                    Case 2
                        s(0) = 100 - 52
                        s(1) = 100 - 20
                    Case 3
                        s(0) = 100 - 72
                        s(1) = 100 + 10
                    Case 4
                        s(0) = 100 - 72
                        s(1) = 100 + 10
                    Case 5
                        s(0) = 100 - 72
                        s(1) = 100 + 10
                End Select
                M.x = .GraphicFiles(.Map.Effect(i).GraphicID).BlackPicPixel.x
                M.y = .GraphicFiles(.Map.Effect(i).GraphicID).BlackPicPixel.y
                M.Height = .GraphicFiles(.Map.Effect(i).GraphicID).BlackPicPixel.Height
                M.Width = .GraphicFiles(.Map.Effect(i).GraphicID).BlackPicPixel.Width
                
                For k = 0 To 1
                        If .Map.Effect(i).Visible(k) = True Then
                                     
                            If .Map.Effect(i).FrameCount = 0 Then .Map.Effect(i).FrameCount = 1
                            
                            If .Map.Effect(i).FrameRun(k) = True Then
                                h = .EffectTimer(i) * .Map.Effect(i).Timer(k) Mod (.Map.Effect(i).FrameCount * 2)
                                c = .Map.Effect(i).Matrix.x
1                                Select Case .Map.Effect(i).Matrix.y
                                    Case -2
                                        
                                        Select Case h
                                            Case 0 To c - 1
                                                a = 0
                                                b = h
                                            Case c To 2 * c - 1
                                                a = 1
                                                b = h - c
                                            Case 2 * c To 3 * c - 1
                                                a = 1
                                                b = Abs(-1 * h + 3 * c - 1)
                                            Case 3 * c To 4 * c - 1
                                                a = 0
                                                b = Abs(-1 * h + 4 * c - 1)
                                        End Select
                                        
                                    Case -1
    
                                        Select Case h
                                            Case 0 To c - 1
                                                a = 0
                                                b = h
                                            Case c To 2 * c - 1
                                                a = 0
                                                b = Abs(-1 * h + 2 * c - 1)
                                        End Select
                                        
                                    Case 22
                                    
                                        Select Case h
                                            Case 0 To c - 1
                                                a = 0
                                                b = h
                                            Case c To 2 * c - 1
                                                a = 1
                                                b = h - c
                                            Case 2 * c To 3 * c - 1
                                                a = 5
                                                b = Abs(-1 * h + 3 * c - 1)
                                            Case 3 * c To 4 * c - 1
                                                a = 4
                                                b = Abs(-1 * h + 4 * c - 1)
                                        End Select
                                    Case 11
                                        h = .EffectTimer(i) * .Map.Effect(i).Timer(k) Mod (.Map.Effect(i).FrameCount)
                                        Select Case h
                                            Case 0 To c - 1
                                                a = 0
                                                b = h
                                            Case c To 2 * c - 1
                                                a = 1
                                                b = h - c
                                                If h = 2 * c - 1 Then .Map.Effect(i).FrameRun(k) = False
                                        End Select
                                        
                                    Case 12
                                        h = .EffectTimer(i) * .Map.Effect(i).Timer(k) Mod (.Map.Effect(i).FrameCount)
                                        Select Case h
                                            Case 0 To c - 1
                                                a = 1
                                                b = Abs(-1 * h + c - 1)
                                            Case c To 2 * c - 1
                                                a = 0
                                                b = Abs(-1 * h + 2 * c - 1)
                                                If h = 2 * c - 1 Then .Map.Effect(i).FrameRun(k) = False
                                        End Select
                                End Select
    
                        
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Effect(i).DrawPosition(k).x * .Map.Tile_Object.Width - s(0), .Map.TilesInfo.y + .Map.Effect(i).DrawPosition(k).y * .Map.Tile_Object.Height - s(1), .Map.Effect(i).MatrixGraphic.Width, .Map.Effect(i).MatrixGraphic.Height, .Buffer.TileSetBmp(.Map.Effect(i).GraphicID), .Map.Effect(i).LoadPosition.x + b * .Map.Effect(i).LoadPosition.Width + M.x * M.Width, .Map.Effect(i).LoadPosition.y + a * .Map.Effect(i).LoadPosition.Height + M.y * M.Height, vbSrcAnd
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Effect(i).DrawPosition(k).x * .Map.Tile_Object.Width - s(0), .Map.TilesInfo.y + .Map.Effect(i).DrawPosition(k).y * .Map.Tile_Object.Height - s(1), .Map.Effect(i).MatrixGraphic.Width, .Map.Effect(i).MatrixGraphic.Height, .Buffer.TileSetBmp(.Map.Effect(i).GraphicID), .Map.Effect(i).LoadPosition.x + b * .Map.Effect(i).LoadPosition.Width, .Map.Effect(i).LoadPosition.y + a * .Map.Effect(i).LoadPosition.Height, vbSrcPaint
                               
                            Else
                                
                            End If
                            
                        End If
                Next k
            Next i
        End If
    End With
End Sub


'=============================================================
'Describe:Layers that mainly draw game forms
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles9(this_Graphic As m_Graphic) '
    Dim i, j As Integer
    Dim M As m_GraphicPosition
    With this_Graphic
        i = 3
        Do Until i < 0
 
            M.x = .GraphicFiles(.Map.Windows(i).GraphicID).BlackPicPixel.x
            M.y = .GraphicFiles(.Map.Windows(i).GraphicID).BlackPicPixel.y
            M.Height = .GraphicFiles(.Map.Windows(i).GraphicID).BlackPicPixel.Height
            M.Width = .GraphicFiles(.Map.Windows(i).GraphicID).BlackPicPixel.Width
            
            If i = 2 Then
                Select Case .Player(CurrentPlayerID).Info.MoveDirection
                
                    Case 0
                        .Map.Windows(i).DrawPosition.x = (.Player(CurrentPlayerID).Info.C_Position.x) * .Map.Tile_Object.Width - .Map.Windows(i).LoadPosition.Width / 2 + 20
                        .Map.Windows(i).DrawPosition.y = (.Player(CurrentPlayerID).Info.C_Position.y + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Height - .Map.Windows(i).LoadPosition.Height - 60
                    Case 1
                        .Map.Windows(i).DrawPosition.x = (.Player(CurrentPlayerID).Info.C_Position.x + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Width - .Map.Windows(i).LoadPosition.Width / 2 + 20
                        .Map.Windows(i).DrawPosition.y = (.Player(CurrentPlayerID).Info.C_Position.y) * .Map.Tile_Object.Height - .Map.Windows(i).LoadPosition.Height - 60
                    Case 2
                        .Map.Windows(i).DrawPosition.x = (.Player(CurrentPlayerID).Info.C_Position.x + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Width - .Map.Windows(i).LoadPosition.Width / 2 + 20
                        .Map.Windows(i).DrawPosition.y = (.Player(CurrentPlayerID).Info.C_Position.y) * .Map.Tile_Object.Height - .Map.Windows(i).LoadPosition.Height - 60
                    Case 3
                        .Map.Windows(i).DrawPosition.x = (.Player(CurrentPlayerID).Info.C_Position.x) * .Map.Tile_Object.Width - .Map.Windows(i).LoadPosition.Width / 2 + 20
                        .Map.Windows(i).DrawPosition.y = (.Player(CurrentPlayerID).Info.C_Position.y + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Height - .Map.Windows(i).LoadPosition.Height - 60
                  
                
                End Select
                If this_Switch.Talk = True Then
                    .SpriteFont(2).x = .Map.Windows(2).DrawPosition.x + 30
                    .SpriteFont(2).y = .Map.Windows(2).DrawPosition.y + 92
    '                .Map.Windows(i).DrawPosition.X = (.Player(CurrentPlayerID).Info.C_Position.X + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Width - .Map.Windows(i).LoadPosition.Width / 2 + 20
    '                .Map.Windows(i).DrawPosition.Y = (.Player(CurrentPlayerID).Info.C_Position.Y + .Player(CurrentPlayerID).Info.PositionTimer) * .Map.Tile_Object.Height - .Map.Windows(i).LoadPosition.Height - 60
                End If
            End If
            .SpriteFont(i).Visiable = .Map.Windows(i).Visible
            If .Map.Windows(0).Visible = True Then
                .SpriteFont(1).Visiable = False
                .SpriteFont(2).Visiable = False
                If .SpriteFont(5).Visiable = True Then
                    .SpriteFont(0).Visiable = False
                    .SpriteFont(4).Visiable = False
                End If
                If .SpriteFont(4).Visiable = True Then
                    .SpriteFont(0).Visiable = False
                    .SpriteFont(5).Visiable = False
                End If
            Else
                
            End If
            
            If .Map.Windows(i).Visible = True Then
                
                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Windows(i).DrawPosition.x, .Map.TilesInfo.y + .Map.Windows(i).DrawPosition.y, .Map.Windows(i).LoadPosition.Width, .Map.Windows(i).LoadPosition.Height, .Buffer.TileSetBmp(.Map.Windows(i).GraphicID), .Map.Windows(i).LoadPosition.x + M.x * M.Width, .Map.Windows(i).LoadPosition.y + M.y * M.Height, vbSrcAnd
                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Windows(i).DrawPosition.x, .Map.TilesInfo.y + .Map.Windows(i).DrawPosition.y, .Map.Windows(i).LoadPosition.Width, .Map.Windows(i).LoadPosition.Height, .Buffer.TileSetBmp(.Map.Windows(i).GraphicID), .Map.Windows(i).LoadPosition.x, .Map.Windows(i).LoadPosition.y, vbSrcPaint
            
            End If
            i = i - 1
        Loop
    End With
End Sub


'=============================================================
'Describe:Processing Functions for Disallowing Passage on Maps  Special effect function
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Events(this_Graphic As m_Graphic, m_Switch As m_Switch) '
    Dim i, j As Integer
    Dim a As Integer
    Dim Tile32 As Byte
    
    a = 5
    Tile32 = 32
    With this_Graphic
        If m_Switch.Event = True Then
            For i = 0 To 15
                For j = 0 To 13
                    If .Map.Events(i, j) > 0 Then
                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tile_Object.Width + a, .Map.TilesInfo.y + j * .Map.Tile_Object.Height + a, Tile32, Tile32, .Buffer.TileSetBmp(9), Tile32 * .Map.GameEvents(.Map.Events(i, j)).PicPosition.x + 512, Tile32 * .Map.GameEvents(.Map.Events(i, j)).PicPosition.y, vbSrcAnd
                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tile_Object.Width + a, .Map.TilesInfo.y + j * .Map.Tile_Object.Height + a, Tile32, Tile32, .Buffer.TileSetBmp(9), Tile32 * .Map.GameEvents(.Map.Events(i, j)).PicPosition.x, Tile32 * .Map.GameEvents(.Map.Events(i, j)).PicPosition.y, vbSrcPaint
                    End If
                Next j
            Next i
        Else
            
        End If
        
         If m_Switch.Block = True Then
            For i = 0 To 15
                For j = 0 To 13
                    'If .Map.Block(i, j) = 1 Then
                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tile_Object.Width + a, .Map.TilesInfo.y + j * .Map.Tile_Object.Height + a, Tile32, Tile32, .Buffer.TileSetBmp(9), Tile32 * .Map.GameEvents(CInt(.Map.Block(i, j))).PicPosition.x + 512, Tile32 * .Map.GameEvents(CInt(.Map.Block(i, j))).PicPosition.y, vbSrcAnd
                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tile_Object.Width + a, .Map.TilesInfo.y + j * .Map.Tile_Object.Height + a, Tile32, Tile32, .Buffer.TileSetBmp(9), Tile32 * .Map.GameEvents(CInt(.Map.Block(i, j))).PicPosition.x, Tile32 * .Map.GameEvents(CInt(.Map.Block(i, j))).PicPosition.y, vbSrcPaint
                    'End If
                Next j
            Next i
        Else
        
        End If
        
        
    End With
End Sub

'=============================================================
'Describe: Layer of text drawn on a map
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Draw_Tiles_Font(this_Graphic As m_Graphic, Optional IsSpriteFont As Boolean)      '
    Dim a As Integer
 

    With this_Graphic
            If IsSpriteFont = True Then
                For a = 0 To UBound(.SpriteFont) - 1
                    If .SpriteFont(a).Visiable = True Then
                        If .SpriteFont(a).G_MaxLineLen > SpriteFont_MaxLineLen Then .SpriteFont(a).G_MaxLineLen = SpriteFont_MaxLineLen
                        Call FontDraw(this_Graphic, a, .SpriteFont(a).x, .SpriteFont(a).y, .SpriteFont(a).G_MaxLineLen, IsSpriteFont)
                    End If
                Next a
            Else
                 For a = 0 To UBound(.WindowsFont) - 1
                    If .WindowsFont(a).Visiable = True Then
                        If .WindowsFont(a).G_MaxLineLen > SpriteFont_MaxLineLen Then .WindowsFont(a).G_MaxLineLen = SpriteFont_MaxLineLen
                        Call FontDraw(this_Graphic, a, .WindowsFont(a).x, .WindowsFont(a).y, .WindowsFont(a).G_MaxLineLen, IsSpriteFont)
                    End If
                Next a
            
            End If
    
    End With

End Sub



'=============================================================
'Describe: Layer of text drawn on a map
'Author:  Chenlyu
'Parameter:
'=============================================================

Private Sub FontDraw(this_Graphic As m_Graphic, SpriteFontID As Integer, DRLeft As Integer, DRTop As Integer, DRLen As Integer, Optional IsSpriteFont As Boolean)  '画字库
Dim FontText As String
Const BubbleSectionSize As Long = 6 ' 泡泡图形的宽和高的大小

With this_Graphic
    If IsSpriteFont = True Then
         If .SpriteFont(SpriteFontID).G_WordWraped = False Then
            .SpriteFont(SpriteFontID).G_String = Engine_WordWrap(this_Graphic.Font_Default, .SpriteFont(SpriteFontID).G_String, DRLen)
            .SpriteFont(SpriteFontID).G_WordWraped = True
        End If
        FontText = .SpriteFont(SpriteFontID).G_String
    Else
         If .WindowsFont(SpriteFontID).G_WordWraped = False Then
            .WindowsFont(SpriteFontID).G_String = Engine_WordWrap(this_Graphic.Font_Default, .WindowsFont(SpriteFontID).G_String, DRLen)
            .WindowsFont(SpriteFontID).G_WordWraped = True
        End If
        FontText = .WindowsFont(SpriteFontID).G_String
    End If


    '最后渲染文字
    Call Render_Text(this_Graphic, FontText, DRLeft + BubbleSectionSize, DRTop + BubbleSectionSize)
End With
End Sub

'=============================================================
'Describe: Layer of text drawn on a map
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer

Dim i As Integer
Dim a As Integer
    '确定文本有内容
    If LenB(Text) = 0 Then Exit Function
    
    '循环
    For i = 1 To Len(Text)
        a = Asc(Mid$(Text, i, 1))
        If a >= 0 And a < 256 Then
        Else
        a = 0
        '宽度
            
        End If
        GetTextWidth = GetTextWidth + UseFont.HeaderInfo.CharWidth(a)
    Next i

End Function



'=============================================================
'Describe: Layer of text drawn on a map
'Author:  Chenlyu
'Parameter:
'=============================================================

Private Sub Render_Text(this_Graphic As m_Graphic, this_Text As String, ByVal x As Integer, ByVal y As Integer)

 
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim V As Single
Dim i As Long
Dim j As Long

Dim YOffset As Single
 
With this_Graphic
                    
    '检查文本
    If LenB(this_Text) = 0 Then Exit Sub
 
 

    '将文本转化成数组
    TempStr = Split(this_Text, vbCrLf)
    
    
    '如果没有换行符，则循环
    For i = 0 To UBound(TempStr)
        If Len(Trim(TempStr(i))) > 0 Then
            YOffset = i * .Font_Default.CharHeight
            Count = 0
        
            '将字符转换成ascii值
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
        
            '循环字符
            For j = 1 To Len(TempStr(i))

                '检查关键字符
                If Ascii(j - 1) = 124 Then '如果字符是 "|"
 
                Else
                    Row = (Ascii(j - 1) - .Font_Default.HeaderInfo.BaseCharOffset) \ .Font_Default.RowPitch
                    u = ((Ascii(j - 1) - .Font_Default.HeaderInfo.BaseCharOffset) - (Row * .Font_Default.RowPitch)) * .Font_Default.ColFactor
                    V = Row * .Font_Default.RowFactor
                    
                    BitBlt .Buffer.BackBuffer, x + Count, y + YOffset, .Font_Default.HeaderInfo.CharWidth(Ascii(j - 1)), .Font_Default.CharHeight, .Buffer.TileSetBmp(8), u * .Font_Default.HeaderInfo.BitmapWidth + 512, V * .Font_Default.HeaderInfo.BitmapHeight, vbSrcAnd
                    BitBlt .Buffer.BackBuffer, x + Count, y + YOffset, .Font_Default.HeaderInfo.CharWidth(Ascii(j - 1)), .Font_Default.CharHeight, .Buffer.TileSetBmp(8), u * .Font_Default.HeaderInfo.BitmapWidth, V * .Font_Default.HeaderInfo.BitmapHeight, vbSrcPaint
                    
  
                    '在位置的渲染转到的下一个角色
                    Count = Count + .Font_Default.HeaderInfo.CharWidth(Ascii(j - 1))
                
                End If
                
            Next j
            
        End If
    Next i
    
End With
End Sub

'=============================================================
'Describe: Layer of text drawn on a map
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function Engine_WordWrap(ByRef UseFont As CustomFont, ByVal Text As String, ByVal MaxLineLen As Integer) As String

Dim TempSplit() As String
Dim TSLoop As Long
Dim LastSpace As Long
Dim Size As Long
Dim i, j As Long
Dim b As Long
Dim a, d As Integer
Dim c() As String
Dim s As String

    '字符太短
    If Len(Text) < 2 Then
        Engine_WordWrap = Text
        Exit Function
    End If
    
    '检查是否有换行
    
    Text = ReplacePlus(Text, vbNewLine & vbNewLine, vbNewLine)
    
    TempSplit = Split(Text, vbNewLine)
    
    ReDim c(32767)
    d = 0
    For TSLoop = 0 To UBound(TempSplit)
        
        '清除新行的值
        Size = 0
        b = 1
        LastSpace = 1
        
        '加上 vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        '仅检查“ ”和“_”
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            '在每个字符中查找
            For i = 1 To Len(TempSplit(TSLoop))
            
                '如果是空格，则先存储以便于后期处理
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": LastSpace = i
                    Case "_": LastSpace = i
                    Case "-": LastSpace = i
                End Select
    
                '不算 "|" 字符的总值
                If Not Mid$(TempSplit(TSLoop), i, 1) = "|" Then
                    a = Asc(Mid$(TempSplit(TSLoop), i, 1))
                    If a >= 0 And a < 256 Then
                    
                    Else
                        a = 0
                    End If
                    Size = Size + UseFont.HeaderInfo.CharWidth(a)
                End If
                
                '检查是否太大
                If Size > MaxLineLen Then
                    If i - LastSpace > 4 Then
                        'Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                        If d > 14 Then
                            For j = 0 To 13
                                c(i) = c(j + 1)
                            Next j
                            c(14) = Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                        Else
                            c(d) = Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                            d = d + 1
                        End If
                        
                        
                        b = i - 1
                        Size = 0
                        
                    Else
                        'Engine_WordWrap = Engine_WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, LastSpace - B)) & vbNewLine
                        If d > 14 Then
                            For j = 0 To 13
                                c(j) = c(j + 1)
                            Next j
                            c(14) = Trim$(Mid$(TempSplit(TSLoop), b, Abs(LastSpace - b))) & vbNewLine
                        Else
                            c(d) = Trim$(Mid$(TempSplit(TSLoop), b, Abs(LastSpace - b))) & vbNewLine
                            d = d + 1
                        End If
                        
                        b = LastSpace + 1
                        Size = GetTextWidth(UseFont, Mid$(TempSplit(TSLoop), LastSpace, i - LastSpace))
                        
                    End If
                End If
                
                '处理剩余的
                If i = Len(TempSplit(TSLoop)) Then
                    If b <> i Then
'                        Engine_WordWrap = Engine_WordWrap & Mid$(TempSplit(TSLoop), B, i)
                        If d > 14 Then
                            For j = 0 To 13
                                c(j) = c(j + 1)
                            Next j
                            c(14) = Mid$(TempSplit(TSLoop), b, i)
                        Else
                            c(d) = Mid$(TempSplit(TSLoop), b, i)
                            d = d + 1
                        End If
            
                    End If
                End If
            Next i
        Else
            ' Trim(TempSplit(TSLoop))
'            Engine_WordWrap = Engine_WordWrap & Trim(TempSplit(TSLoop))
            If d > 14 Then
                For i = 0 To 13
                    c(i) = c(i + 1)
                Next i
                c(14) = Trim(TempSplit(TSLoop))
            Else
                c(d) = Trim(TempSplit(TSLoop))
                d = d + 1
            End If
            
        End If
    Next TSLoop


        
        Engine_WordWrap = ""
        For i = 0 To 99
             Engine_WordWrap = Engine_WordWrap & c(i)
        Next i


End Function



'=============================================================
'Describe: Timetable functions
'Author:  Chenlyu
'Parameter:
'=============================================================

Private Sub Draw_DisplayDigital_Timer(this_Graphic As m_Graphic, myX As Integer, myY As Integer, this_Switch As m_Switch)
    Dim SPLT() As String
    Dim i As Integer
    Dim s As String
    Dim PIC_WIDTH_PIX    As Integer
    Dim PIC_HEIGHT_PIX   As Integer
    
    Dim Pic_pix As m_Position
    
    PIC_WIDTH_PIX = 14
    PIC_HEIGHT_PIX = 20
    
    Pic_pix.x = 256
    Pic_pix.y = 556
    
    If this_Switch.Timer = False Then Exit Sub
    
    With this_Graphic
    
    s = Right(0 & .Clock.myHour, 2) & ":" & Right(0 & .Clock.myMinute, 2) & ":00"
    
    ReDim SPLT(1 To Len(s))
    For i = 1 To Len(s)
        SPLT(i) = Mid(s, i, 1)
    Next i
    
    
    BitBlt .Buffer.BackBuffer, myX, myY, 134, 46, .Buffer.TileSetBmp(9), Pic_pix.x, Pic_pix.y - 46, vbSrcAnd
    BitBlt .Buffer.BackBuffer, myX, myY, 134, 46, .Buffer.TileSetBmp(9), Pic_pix.x, Pic_pix.y - 46, vbSrcPaint
    
    For i = 1 To UBound(SPLT())

        Select Case UCase(SPLT(i))
            Case "0"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(0 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(0 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint
            Case "1"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(1 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(1 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "2"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(2 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(2 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "3"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(3 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(3 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "4"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(4 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(4 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "5"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(5 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(5 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "6"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(6 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(6 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "7"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(7 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(7 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "8"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(8 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(8 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "9"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(9 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(9 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case " "
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(10 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(10 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case ":"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(11 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(11 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case ";"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(12 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(12 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "-"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(13 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(13 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "'"
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(14 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(14 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case "."
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(15 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(15 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

            Case ","
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(16 * PIC_WIDTH_PIX) + 512, Pic_pix.y, vbSrcAnd
               BitBlt .Buffer.BackBuffer, myX + 8 + (i - 1) * PIC_WIDTH_PIX, myY + 19, PIC_WIDTH_PIX, PIC_HEIGHT_PIX, .Buffer.TileSetBmp(9), Pic_pix.x + CLng(16 * PIC_WIDTH_PIX), Pic_pix.y, vbSrcPaint

        End Select

    Next i
    
    

    End With
    
 
End Sub

'=============================================================
'Describe: Motion Testing of Players
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Player_Move_Test(this_Graphic As m_Graphic, this_i As Byte)
    Dim X0 As Integer
    Dim Y0 As Integer
    Dim M As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim OKtest As Integer
    Dim mm As Integer, nn As Integer
    
    
    With this_Graphic.Player(this_i).Info
    

    
    
    If .G_Position.x <> .C_Position.x Or .G_Position.y <> .C_Position.y Then

        Call FindPath(.C_Position.x, .C_Position.y, .G_Position.x, .G_Position.y, .RoadData, .Roading, True)
        
        If .RoadData >= 0 Then
            .RodeTry = False
            
            If (.RoadData <= -1) And (.C_Position.x <> .Next_Position.x) And (.Next_Position.y <> .C_Position.y) Then GoTo NoWays '当出现无法抵达目的地的情况时直接退出该判断

            If .RoadData >= 0 Then
                X0 = .Roading(.RoadData).x
                Y0 = .Roading(.RoadData).y
                .RoadData = .RoadData - 1
            ElseIf .RoadData = 0 Then
                X0 = .G_Position.x
                Y0 = .G_Position.y
            End If
                        
                M = X0 - .C_Position.x
                n = Y0 - .C_Position.y
                If M = -1 Then
                    .MoveDirection = MoveLeft
                End If
                If M = 1 Then
                    .MoveDirection = MoveRight
                End If
                If n = -1 Then
                    .MoveDirection = MoveUp
                End If
                If n = 1 Then
                    .MoveDirection = MoveDown
                End If
                
                Call Player_Event_AddMove(.MoveDirection, this_i)
                .Next_Position.x = X0 '.G_Position.X
                .Next_Position.y = Y0 '.G_Position.Y
                Exit Sub
            End If

NoWays:
        If .RoadData = -1 Then
            If (.C_Position.x = .Next_Position.x) And (.C_Position.y = .Next_Position.y) Then
                '准确到达目的地
                
            Else
                .RodeTry = True
                Randomize
                i = Int((2 * Rnd))
                i = i * 2
                i = 1 - i
                Randomize
                j = Int((2 * Rnd))
                j = j * 2
                j = 1 - j
                If (.Next_Position.x + i <= 16) And (.Next_Position.y + j <= 14) And (.Next_Position.x + i >= 0) And (.Next_Position.y + j >= 0) Then
                    .Next_Position.x = .Next_Position.x + i
                    .Next_Position.y = .Next_Position.y + j
                End If
            End If
        End If
    
    Else
        .RodeTry = True
        .RoadXY = False
    End If
    
    
'    .C_Position.X = .Next_Position.X
'    .C_Position.Y = .Next_Position.Y
    
End With
End Sub

'=============================================================
'Describe: Move function after event is clicked
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Player_Event_ClicktoMove(m_MoveEvent As Byte, m_PlayerID As Byte)
With this_Graphic.Player(m_PlayerID).Info
    Select Case m_MoveEvent
    
        Case 0
             .G_Position.y = .C_Position.y + 1
             .G_Position.x = .C_Position.x
        Case 1
            .G_Position.x = .C_Position.x - 1
            .G_Position.y = .C_Position.y
        Case 2
            .G_Position.x = .C_Position.x + 1
            .G_Position.y = .C_Position.y
        Case 3
            .G_Position.y = .C_Position.y - 1
            .G_Position.x = .C_Position.x
    End Select
    
    .Next_Position.x = .G_Position.x
    .Next_Position.y = .G_Position.y
    .RoadXY = False
End With
    Call Player_Event_AddMove(m_MoveEvent, m_PlayerID)
End Sub

'=============================================================
'Describe: Move function after event is clicked
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Player_Event_AddMove(m_MoveEvent As Byte, m_PlayerID As Byte)
    With this_Graphic.Player(m_PlayerID)
        .Info.G_Event(.Info.EventTimer) = m_MoveEvent
        .Info.EventTimer = .Info.EventTimer + 1
        If .Info.PositionTimer = 0 Then .Info.EventBreak = False
    End With
End Sub

'=============================================================
'Describe: Move function after event is clicked
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Player_Event_MinusMove(m_MoveEvent As Byte, m_PlayerID As Byte)
    With this_Graphic.Player(m_PlayerID)
        If EventBreak = True Then .Info.EventBreak = True
        If CurrentEventTimer < 1022 And CurrentEventTimer < .Info.EventTimer Then
            CurrentEventTimer = CurrentEventTimer + 1
            If CurrentEventTimer = .Info.EventTimer Then
                CurrentEventTimer = CurrentEventTimer
                .Info.EventBreak = True
            End If
        Else
            .Info.EventBreak = True
        End If
         
        Select Case m_MoveEvent
        
            Case MoveRight
            
            Case MoveDown
            
            Case MoveLeft
            
            Case MoveUp
        
        End Select
        
    End With
End Sub

