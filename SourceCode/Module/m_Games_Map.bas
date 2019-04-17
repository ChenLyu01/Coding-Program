Attribute VB_Name = "m_Games_Map"
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
'Describe:The corresponding map needs to be loaded
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Sub Game_MapDataLoad(this_FilePath As m_FilePath)
    Dim a, b, i, j As Integer
    Dim s As String
    Dim t() As String
    
    CurrentPlayerID = 0
    
    
    CurrentEventTimer = 0
    CurrentGameEvents = 0
    this_Switch.Object = True
    this_Switch.Player = True
    this_Switch.Effect = True
    this_Switch.Pathway = True
    this_Switch.Letters = True
    this_Switch.Steps = True
    this_Switch.PathwayWithBlock = False
    this_Switch.Talk = False

'    this_Switch.Timer = True
    
    
    

        
        
    With this_Graphic
    
        For i = 0 To 15
            For j = 0 To 13
                .Map.Events(i, j) = 0
                .Map.Block(i, j) = 0
            Next j
        Next i
        
        .Map.BlockTilesID = 0
        

        
        .Font_Default.HeaderInfo.CellHeight = 32
        .Font_Default.HeaderInfo.CellWidth = 32
        .Font_Default.HeaderInfo.BitmapWidth = 512
        .Font_Default.HeaderInfo.BitmapHeight = 512
        
        .Font_Default.CharHeight = Abs(.Font_Default.HeaderInfo.CellHeight)
        .Font_Default.RowPitch = .Font_Default.HeaderInfo.BitmapWidth / .Font_Default.HeaderInfo.CellWidth
        .Font_Default.ColFactor = .Font_Default.HeaderInfo.CellWidth / .Font_Default.HeaderInfo.BitmapWidth
        .Font_Default.RowFactor = .Font_Default.HeaderInfo.CellHeight / .Font_Default.HeaderInfo.BitmapHeight
        
        
        s = Trim("11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|11|3|4|4|8|7|10|8|2|5|5|6|7|4|4|4|5|8|7|8|8|8|8|8|8|8|8|4|4|7|7|7|6|11|9|7|8|8|7|7|9|10|4|6|8|7|10|9|9|7|9|8|7|8|9|9|12|9|8|7|5|5|5|7|7|6|7|7|6|7|6|4|7|8|4|4|7|4|12|8|6|7|6|6|5|5|8|8|10|7|8|5|6|4|6|8|11|8|11|3|6|6|9|6|6|6|15|7|5|11|11|7|11|11|3|3|6|6|5|7|9|6|10|5|5|10|11|5|8|3|4|7|7|7|8|4|5|6|10|6|7|7|4|10|7|5|7|6|6|6|6|6|4|6|6|6|7|12|12|12|6|9|9|9|9|9|9|12|8|7|7|7|7|4|4|4|4|8|9|9|9|9|9|9|7|9|9|9|9|9|8|7|7|7|7|7|7|7|7|10|6|6|6|6|6|4|4|4|4|6|8|6|6|6|6|6|7|6|8|8|8|8|8|7|8|")
        
        t = Split(s, "|")
        For j = 0 To 255
            .Font_Default.HeaderInfo.CharWidth(j) = t(j) * 2
        Next j
    
    
        .GraphicFiles(0).FilesName = "LessonOne.bmp"
        .GraphicFiles(0).BlackPicPixel.x = 1
        .GraphicFiles(0).BlackPicPixel.y = 0
        .GraphicFiles(0).BlackPicPixel.Width = 245
        
        .GraphicFiles(1).FilesName = "Object2.gif"
        .GraphicFiles(1).BlackPicPixel.x = 1
        .GraphicFiles(1).BlackPicPixel.y = 0
        .GraphicFiles(1).BlackPicPixel.Width = 494
        
        .GraphicFiles(2).FilesName = "Player1.gif"
        .GraphicFiles(2).BlackPicPixel.x = 1
        .GraphicFiles(2).BlackPicPixel.y = 0
        .GraphicFiles(2).BlackPicPixel.Width = 512
        
        .GraphicFiles(3).FilesName = "Player.gif"
        .GraphicFiles(3).BlackPicPixel.x = 0
        .GraphicFiles(3).BlackPicPixel.y = 1
        .GraphicFiles(3).BlackPicPixel.Height = 128
        
        .GraphicFiles(4).FilesName = "Mana.gif"
        .GraphicFiles(4).BlackPicPixel.x = 0
        .GraphicFiles(4).BlackPicPixel.y = 1
        .GraphicFiles(4).BlackPicPixel.Height = 256
        
        .GraphicFiles(5).FilesName = "Windows.gif"
        .GraphicFiles(5).BlackPicPixel.x = 0
        .GraphicFiles(5).BlackPicPixel.y = 1
        .GraphicFiles(5).BlackPicPixel.Height = 520
        
        
        .GraphicFiles(6).FilesName = "Effect1.gif"
        .GraphicFiles(6).BlackPicPixel.x = 0
        .GraphicFiles(6).BlackPicPixel.y = 1
        .GraphicFiles(6).BlackPicPixel.Height = 256
        
        .GraphicFiles(7).FilesName = "letter.bmp"
        .GraphicFiles(7).BlackPicPixel.x = 0
        .GraphicFiles(7).BlackPicPixel.y = 0
        .GraphicFiles(7).BlackPicPixel.Height = 0
        
        .GraphicFiles(10).FilesName = "jiequn.gif"
        .GraphicFiles(10).BlackPicPixel.x = 1
        .GraphicFiles(10).BlackPicPixel.y = 0
        .GraphicFiles(10).BlackPicPixel.Width = 223
        
        .Screen.Width = 1024
        .Screen.Height = 768
        
        .Posi.x = 0
        .Posi.y = 0
        
        a = 0
        For j = 0 To 3
            For i = 0 To 25
                For b = 0 To 31
                    .Map.Letters(j, i).GraphicID = 7
                    .Map.Letters(j, i).GraphicPosition.x = i * 42
                    .Map.Letters(j, i).GraphicPosition.y = j * 52
                    .Map.Letters(j, i).GraphicPosition.Width = 42
                    .Map.Letters(j, i).GraphicPosition.Height = 52
                    .Map.Letters(j, i).MapPosition(b).y = -1
                    .Map.Letters(j, i).MapPosition(b).x = -1
                Next b
                
                If a < 52 Then
                    For b = 0 To 31
                        .Map.Letter(a).GraphicID = 7
                        .Map.Letter(a).GraphicPosition.x = i * 42
                        .Map.Letter(a).GraphicPosition.y = (j + 4) * 52
                        .Map.Letter(a).GraphicPosition.Width = 42
                        .Map.Letter(a).GraphicPosition.Height = 52
                        .Map.Letter(a).MapPosition(b).x = -1
                        .Map.Letter(a).MapPosition(b).y = -1
                    Next b
                End If
                a = a + 1
            Next i
         Next j
        
        
'**************************************************************************************
        .Player(0).GraphicID = 2
        .Player(0).Tile.Width = 128
        .Player(0).Tile.Height = 128
        .Player(0).Pic.x = -46 + 26
        .Player(0).Pic.y = -80 + 78
        .Player(0).Pic_Temp(0).x = .Player(0).Pic.x
        .Player(0).Pic_Temp(0).y = .Player(0).Pic.y
        .Player(0).Pic_Temp(1).x = .Player(0).Pic.x
        .Player(0).Pic_Temp(1).y = .Player(0).Pic.y - 10
        

       
        '箭头
    For i = 0 To 99
        Steps(i).GraphicID = 9
        Steps(i).Arrow(1).x = 292
        Steps(i).Arrow(1).y = 422
        Steps(i).Arrow(1).Width = 55
        Steps(i).Arrow(1).Height = 55

        Steps(i).Arrow(3).x = 292 + 55
        Steps(i).Arrow(3).y = 422
        Steps(i).Arrow(3).Width = 55
        Steps(i).Arrow(3).Height = 55

        Steps(i).Arrow(2).x = 292 + 55 * 2
        Steps(i).Arrow(2).y = 422
        Steps(i).Arrow(2).Width = 55
        Steps(i).Arrow(2).Height = 55

        Steps(i).Arrow(0).x = 292 + 55 * 3
        Steps(i).Arrow(0).y = 422
        Steps(i).Arrow(0).Width = 55
        Steps(i).Arrow(0).Height = 55
        
        Steps(i).Direction = -1
    Next i
    
        
'        Steps(0).Direction = 0
'        Steps(1).Direction = 0
'        Steps(2).Direction = 0
'        Steps(3).Direction = 0
'        Steps(4).Direction = 1
'        Steps(5).Direction = 0
'        Steps(6).Direction = 0
'        Steps(7).Direction = 0
'        Steps(8).Direction = 2
'        Steps(9).Direction = 2
'        Steps(10).Direction = 2
'        Steps(11).Direction = 3
'        Steps(12).Direction = 3
'        Steps(13).Direction = 3
'
        '地砖0
         .Map.Tiles_Posi(0).x = 278
         .Map.Tiles_Posi(0).y = 540
         
        '地砖1
         .Map.Tiles_Posi(2).x = 404
         .Map.Tiles_Posi(2).y = 541
          
        '地砖2
         .Map.Tiles_Posi(1).x = 529
         .Map.Tiles_Posi(1).y = 540

        .Map.Tile_Map.Width = 84
        .Map.Tile_Map.Height = 84
        .Map.Tile_Object.Width = 42
        .Map.Tile_Object.Height = 42

        .Map.TilesInfo.x = 26
        .Map.TilesInfo.y = 88
        .Map.TilesInfo.Width = 672
        .Map.TilesInfo.Height = 588


        .Map.Tiles(1).GraphicPosition.x = 26
        .Map.Tiles(1).GraphicPosition.y = 592
        .Map.Tiles(1).GraphicPosition.Width = 84
        .Map.Tiles(1).GraphicPosition.Height = 84
        .Map.Tiles(2).GraphicPosition.x = 110
        .Map.Tiles(2).GraphicPosition.y = 592
        .Map.Tiles(2).GraphicPosition.Width = 84
        .Map.Tiles(2).GraphicPosition.Height = 84
        .Map.Tiles(3).GraphicPosition.x = 194
        .Map.Tiles(3).GraphicPosition.y = 592
        .Map.Tiles(3).GraphicPosition.Width = 84
        .Map.Tiles(3).GraphicPosition.Height = 84

        .Map.Windows(2).Visible = True
        .Map.Windows(2).GraphicID = 10
        .Map.Windows(2).DrawPosition.x = 0
        .Map.Windows(2).DrawPosition.y = 0
        .Map.Windows(2).LoadPosition.x = 0
        .Map.Windows(2).LoadPosition.y = 0
        .Map.Windows(2).LoadPosition.Width = 200
        .Map.Windows(2).LoadPosition.Height = 120
        
        .SpriteFont(2).G_String = "Chen Lyu"
        .SpriteFont(2).x = .Map.Windows(2).DrawPosition.x + 70
        .SpriteFont(2).y = .Map.Windows(2).DrawPosition.y + 130
        .SpriteFont(2).G_WordWraped = False
        .SpriteFont(2).Visiable = .Map.Windows(2).Visible
        .SpriteFont(2).G_MaxLineLen = 200
        

        .Map.Windows(0).Visible = True
        .Map.Windows(0).GraphicID = 5
        .Map.Windows(0).DrawPosition.x = 20
        .Map.Windows(0).DrawPosition.y = 50
        .Map.Windows(0).LoadPosition.x = 0
        .Map.Windows(0).LoadPosition.y = 0
        .Map.Windows(0).LoadPosition.Width = 600
        .Map.Windows(0).LoadPosition.Height = 520


        .SpriteFont(0).G_String = ""
        .SpriteFont(0).x = .Map.Windows(0).DrawPosition.x + 70
        .SpriteFont(0).y = .Map.Windows(0).DrawPosition.y + 110
        .SpriteFont(0).G_WordWraped = False
        .SpriteFont(0).Visiable = .Map.Windows(0).Visible
        .SpriteFont(0).G_MaxLineLen = 500
        
        
        .SpriteFont(4).G_String = ""
        .SpriteFont(4).x = .Map.Windows(0).DrawPosition.x + 70
        .SpriteFont(4).y = .Map.Windows(0).DrawPosition.y + 110
        .SpriteFont(4).G_WordWraped = False
        .SpriteFont(4).Visiable = .Map.Windows(0).Visible
        .SpriteFont(4).G_MaxLineLen = 500
        
        .SpriteFont(5).G_String = ""
        .SpriteFont(5).x = .Map.Windows(0).DrawPosition.x + 70
        .SpriteFont(5).y = .Map.Windows(0).DrawPosition.y + 110
        .SpriteFont(5).G_WordWraped = False
        .SpriteFont(5).Visiable = .Map.Windows(0).Visible
        .SpriteFont(5).G_MaxLineLen = 500
        
        
        .Map.Windows(1).Visible = False
        .Map.Windows(1).GraphicID = 5
        .Map.Windows(1).DrawPosition.x = 130
        .Map.Windows(1).DrawPosition.y = 50
        .Map.Windows(1).LoadPosition.x = 600
        .Map.Windows(1).LoadPosition.y = 0
        .Map.Windows(1).LoadPosition.Width = 386
        .Map.Windows(1).LoadPosition.Height = 363
        
        .SpriteFont(1).G_String = ""
        .SpriteFont(1).x = .Map.Windows(1).DrawPosition.x + 70
        .SpriteFont(1).y = .Map.Windows(1).DrawPosition.y + 130
        .SpriteFont(1).G_WordWraped = False
        .SpriteFont(1).Visiable = .Map.Windows(1).Visible
        .SpriteFont(1).G_MaxLineLen = 300
        
        
        fMainForm.Txt_Input.Left = .Map.Windows(1).DrawPosition.x + 60
        fMainForm.Txt_Input.Top = .Map.Windows(1).DrawPosition.y + 300
        
        




        .Map.Effect(0).GraphicID = 6
        .Map.Effect(0).FrameCount = 14
        .Map.Effect(0).Matrix.x = 7 '每一行有多少帧动画
        .Map.Effect(0).Matrix.y = -2 '带回放效果的2行动画
        .Map.Effect(0).MatrixGraphic.Width = 128
        .Map.Effect(0).MatrixGraphic.Height = 128
        .Map.Effect(0).LoadPosition.x = 0
        .Map.Effect(0).LoadPosition.y = 0
        .Map.Effect(0).LoadPosition.Height = 128
        .Map.Effect(0).LoadPosition.Width = 128
        
        .Map.Effect(1).GraphicID = 3
        .Map.Effect(1).FrameCount = 8
        .Map.Effect(1).Matrix.x = 8
        .Map.Effect(1).Matrix.y = -1 '带回放效果的1行动画
        .Map.Effect(1).MatrixGraphic.Width = 128
        .Map.Effect(1).MatrixGraphic.Height = 128
        .Map.Effect(1).LoadPosition.x = 0
        .Map.Effect(1).LoadPosition.y = 0
        .Map.Effect(1).LoadPosition.Height = 128
        .Map.Effect(1).LoadPosition.Width = 128
        
        .Map.Effect(2).GraphicID = 3
        .Map.Effect(2).FrameCount = 8
        .Map.Effect(2).Matrix.x = 8
        .Map.Effect(2).Matrix.y = -1 '带回放效果的1行动画
        .Map.Effect(2).MatrixGraphic.Width = 128
        .Map.Effect(2).MatrixGraphic.Height = 128
        .Map.Effect(2).LoadPosition.x = 0
        .Map.Effect(2).LoadPosition.y = 256
        .Map.Effect(2).LoadPosition.Height = 128
        .Map.Effect(2).LoadPosition.Width = 128
        
        
        .Map.Effect(3).GraphicID = 4
        .Map.Effect(3).FrameCount = 16
        .Map.Effect(3).Matrix.x = 8
        .Map.Effect(3).Matrix.y = 22 '不带回放效果的1行动画
        .Map.Effect(3).MatrixGraphic.Width = 512
        .Map.Effect(3).MatrixGraphic.Height = 128
        .Map.Effect(3).LoadPosition.x = 0
        .Map.Effect(3).LoadPosition.y = 0
        .Map.Effect(3).LoadPosition.Height = 128
        .Map.Effect(3).LoadPosition.Width = 512
        
        
        .Map.Effect(4).GraphicID = 4
        .Map.Effect(4).FrameCount = 16
        .Map.Effect(4).Matrix.x = 8
        .Map.Effect(4).Matrix.y = 11 '不带回放效果的1行动画
        .Map.Effect(4).MatrixGraphic.Width = 512
        .Map.Effect(4).MatrixGraphic.Height = 128
        .Map.Effect(4).LoadPosition.x = 0
        .Map.Effect(4).LoadPosition.y = 0
        .Map.Effect(4).LoadPosition.Height = 128
        .Map.Effect(4).LoadPosition.Width = 512
    
        .Map.Effect(5).GraphicID = 4
        .Map.Effect(5).FrameCount = 16
        .Map.Effect(5).Matrix.x = 8
        .Map.Effect(5).Matrix.y = 12 '不带回放效果的1行动画
        .Map.Effect(5).MatrixGraphic.Width = 512
        .Map.Effect(5).MatrixGraphic.Height = 128
        .Map.Effect(5).LoadPosition.x = 0
        .Map.Effect(5).LoadPosition.y = 512
        .Map.Effect(5).LoadPosition.Height = 128
        .Map.Effect(5).LoadPosition.Width = 512
        
'**************************************************************************************


        
        '在系统默认的事件图里的图片的位置
        '***********************************************************************************
        .Map.GameEvents(60).m_Name = "New File"
        .Map.GameEvents(60).PicPosition.x = 13
        .Map.GameEvents(60).PicPosition.y = 0
        .Map.GameEvents(60).m_Description = "Create a New File"
        
        .Map.GameEvents(61).m_Name = "Open File"
        .Map.GameEvents(61).PicPosition.x = 13
        .Map.GameEvents(61).PicPosition.y = 0
        .Map.GameEvents(60).m_Description = "Open a Exist File"
        
        .Map.GameEvents(62).m_Name = "Close File"
        .Map.GameEvents(62).PicPosition.x = 13
        .Map.GameEvents(62).PicPosition.y = 0
        .Map.GameEvents(62).m_Description = "Close this File"
         
        .Map.GameEvents(1).m_Name = "Block = True"
        .Map.GameEvents(1).PicPosition.x = 8  '禁止通行的图标
        .Map.GameEvents(1).PicPosition.y = 8
        
        .Map.GameEvents(0).m_Name = "Block = False"
        .Map.GameEvents(0).PicPosition.x = 13 '允许通行的图标
        .Map.GameEvents(0).PicPosition.y = 8
        
        .Map.GameEvents(51).m_Name = "Close File"
        .Map.GameEvents(51).PicPosition.x = 14
        .Map.GameEvents(51).PicPosition.y = 0
        .Map.GameEvents(51).m_Description = "Close this File"
        
        .Map.GameEvents(52).m_Name = "Close File"
        .Map.GameEvents(52).PicPosition.x = 14
        .Map.GameEvents(52).PicPosition.y = 0
        .Map.GameEvents(52).m_Description = "Close this File"
        
                
        .Map.GameEvents(50).m_Name = "Open File"
        .Map.GameEvents(50).PicPosition.x = 13
        .Map.GameEvents(50).PicPosition.y = 0
        .Map.GameEvents(50).m_Description = "plarn"
        
         .Map.GameEvents(10).m_Name = "Carry some thing"
        .Map.GameEvents(10).PicPosition.x = 12
        .Map.GameEvents(10).PicPosition.y = 1
        .Map.GameEvents(10).m_Description = "object"
        
        .Map.GameEvents(59).m_Name = "Files"
        .Map.GameEvents(59).PicPosition.x = 12
        .Map.GameEvents(59).PicPosition.y = 0
        .Map.GameEvents(59).m_Description = "Open File"
        
        .Map.GameEvents(63).m_Name = "Files"
        .Map.GameEvents(63).PicPosition.x = 15
        .Map.GameEvents(63).PicPosition.y = 0
        .Map.GameEvents(63).m_Description = "Open File"
        
        .Map.GameEvents(40).m_Name = "Emo"
        .Map.GameEvents(40).PicPosition.x = 10
        .Map.GameEvents(40).PicPosition.y = 7
        .Map.GameEvents(40).m_Description = "Emo"
        
        .Map.GameEvents(41).m_Name = "Emo"
        .Map.GameEvents(41).PicPosition.x = 10
        .Map.GameEvents(41).PicPosition.y = 7
        .Map.GameEvents(41).m_Description = "Emo"
        
        .Map.GameEvents(42).m_Name = "Emo"
        .Map.GameEvents(42).PicPosition.x = 10
        .Map.GameEvents(42).PicPosition.y = 7
        .Map.GameEvents(42).m_Description = "Emo"
        
        '***********************************************************************************
    
        For i = 0 To 11
            .Map.BlockTilesInfo(i).GraphicID = 0
            .Map.BlockTilesInfo(i).GraphicPosition.Width = .Map.Tile_Object.Width
            If i < 9 Then
                .Map.BlockTilesInfo(i).GraphicPosition.Height = .Map.Tile_Object.Width
            Else
                .Map.BlockTilesInfo(i).GraphicPosition.Height = 10
            End If
        Next i

        
        
        b = .Map.BlockTilesID
       
        a = 0
        For j = 0 To 3
            For i = 0 To 2
                .Map.BlockTilesInfo(a).GraphicPosition.x = .Map.Tiles_Posi(b).x + .Map.BlockTilesInfo(a).GraphicPosition.Width * i
                If j < 3 Then
                    .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(a).GraphicPosition.Height * j
                Else
                    .Map.BlockTilesInfo(a).GraphicPosition.y = .Map.Tiles_Posi(b).y + .Map.BlockTilesInfo(0).GraphicPosition.Height * j
                End If
                a = a + 1
            Next i
        Next j

        
'**************************************************************************************
         '地砖颜色
        .Map.TileID = 1
        
        '穿越门
        .Map.Objects(0).GraphicID = 0
        .Map.Objects(0).GraphicPosition.x = 8 + .Map.TilesInfo.x
        .Map.Objects(0).GraphicPosition.y = 0 + .Map.TilesInfo.y
        .Map.Objects(0).GraphicPosition.Width = 237
        .Map.Objects(0).GraphicPosition.Height = 207 - 77
        .Map.Objects(0).LayerNum = 1

        '穿越门的柱子
        .Map.Objects(1).GraphicID = 0
        .Map.Objects(1).GraphicPosition.x = 8 + .Map.TilesInfo.x
        .Map.Objects(1).GraphicPosition.y = 0 + .Map.TilesInfo.y
        .Map.Objects(1).GraphicPosition.Width = 237
        .Map.Objects(1).GraphicPosition.Height = 207
        .Map.Objects(1).LayerNum = 0

        '柱子
        .Map.Objects(11).GraphicID = 0
        .Map.Objects(11).GraphicPosition.x = 3 + .Map.TilesInfo.x
        .Map.Objects(11).GraphicPosition.y = 214 + .Map.TilesInfo.y
        .Map.Objects(11).GraphicPosition.Width = 39
        .Map.Objects(11).GraphicPosition.Height = 63

        '石头的图标
        .Map.Objects(2).GraphicID = 0
        .Map.Objects(2).GraphicPosition.x = 41 + .Map.TilesInfo.x
        .Map.Objects(2).GraphicPosition.y = 229 + .Map.TilesInfo.y
        .Map.Objects(2).GraphicPosition.Width = 47
        .Map.Objects(2).GraphicPosition.Height = 49

        '蓝水晶
        .Map.Objects(3).GraphicID = 0
        .Map.Objects(3).GraphicPosition.x = 90 + .Map.TilesInfo.x
        .Map.Objects(3).GraphicPosition.y = 225 + .Map.TilesInfo.y
        .Map.Objects(3).GraphicPosition.Width = 53
        .Map.Objects(3).GraphicPosition.Height = 52

        '狼人的图标
        .Map.Objects(4).GraphicID = 10
        .Map.Objects(4).GraphicPosition.x = 0
        .Map.Objects(4).GraphicPosition.y = 186
        .Map.Objects(4).GraphicPosition.Width = 84
        .Map.Objects(4).GraphicPosition.Height = 75

        '炸药的图标
        .Map.Objects(5).GraphicID = 0
        .Map.Objects(5).GraphicPosition.x = 197 + .Map.TilesInfo.x
        .Map.Objects(5).GraphicPosition.y = 234 + .Map.TilesInfo.y
        .Map.Objects(5).GraphicPosition.Width = 44
        .Map.Objects(5).GraphicPosition.Height = 44



'        '书籍的图标
        .Map.Objects(6).GraphicID = 0
        .Map.Objects(6).GraphicPosition.x = 0 + .Map.TilesInfo.x
        .Map.Objects(6).GraphicPosition.y = 277 + .Map.TilesInfo.y
        .Map.Objects(6).GraphicPosition.Width = 42
        .Map.Objects(6).GraphicPosition.Height = 42

'
'        '另一本书的图标
        .Map.Objects(7).GraphicID = 0
        .Map.Objects(7).GraphicPosition.x = 44 + .Map.TilesInfo.x
        .Map.Objects(7).GraphicPosition.y = 278 + .Map.TilesInfo.y
        .Map.Objects(7).GraphicPosition.Width = 41
        .Map.Objects(7).GraphicPosition.Height = 42
'
'        '锤子的图标
        .Map.Objects(8).GraphicID = 0
        .Map.Objects(8).GraphicPosition.x = 89 + .Map.TilesInfo.x
        .Map.Objects(8).GraphicPosition.y = 277 + .Map.TilesInfo.y
        .Map.Objects(8).GraphicPosition.Width = 41
        .Map.Objects(8).GraphicPosition.Height = 42
'
'        '栅栏的图标
        .Map.Objects(9).GraphicID = 10
        .Map.Objects(9).GraphicPosition.x = 4
        .Map.Objects(9).GraphicPosition.y = 133
        .Map.Objects(9).GraphicPosition.Width = 138
        .Map.Objects(9).GraphicPosition.Height = 43
'



'        '矿车的图标
        .Map.Objects(10).GraphicID = 0
        .Map.Objects(10).GraphicPosition.x = 0 + .Map.TilesInfo.x
        .Map.Objects(10).GraphicPosition.y = 326 + .Map.TilesInfo.y
        .Map.Objects(10).GraphicPosition.Width = 147
        .Map.Objects(10).GraphicPosition.Height = 118 - 77
        .Map.Objects(10).LayerNum = 1
        
'        '矿车的图标
        .Map.Objects(11).GraphicID = 0
        .Map.Objects(11).GraphicPosition.x = 0 + .Map.TilesInfo.x
        .Map.Objects(11).GraphicPosition.y = 326 + .Map.TilesInfo.y
        .Map.Objects(11).GraphicPosition.Width = 147
        .Map.Objects(11).GraphicPosition.Height = 118
        .Map.Objects(11).LayerNum = 0
        
        
        '
        '炸弹的图标
        .Map.Objects(12).GraphicID = 10
        .Map.Objects(12).GraphicPosition.x = 173 + .Map.TilesInfo.x
        .Map.Objects(12).GraphicPosition.y = 279 + .Map.TilesInfo.y
        .Map.Objects(12).GraphicPosition.Width = 43
        .Map.Objects(12).GraphicPosition.Height = 49
        
        
        '木架的图标
        .Map.Objects(13).GraphicID = 0
        .Map.Objects(13).GraphicPosition.x = 151 + .Map.TilesInfo.x
        .Map.Objects(13).GraphicPosition.y = 352 + .Map.TilesInfo.y
        .Map.Objects(13).GraphicPosition.Width = 89
        .Map.Objects(13).GraphicPosition.Height = 92


        '火山的图标2
        .Map.Objects(31).GraphicID = 1
        .Map.Objects(31).GraphicPosition.x = 187
        .Map.Objects(31).GraphicPosition.y = 501
        .Map.Objects(31).GraphicPosition.Width = 203
        .Map.Objects(31).GraphicPosition.Height = 167

        '食人花的图标
        .Map.Objects(15).GraphicID = 1
        .Map.Objects(15).GraphicPosition.x = 11
        .Map.Objects(15).GraphicPosition.y = 1
        .Map.Objects(15).GraphicPosition.Width = 79
        .Map.Objects(15).GraphicPosition.Height = 92

        '小树的图标
        .Map.Objects(16).GraphicID = 1
        .Map.Objects(16).GraphicPosition.x = 99
        .Map.Objects(16).GraphicPosition.y = 6
        .Map.Objects(16).GraphicPosition.Width = 73
        .Map.Objects(16).GraphicPosition.Height = 87
        
        '红色的植物的图标
        .Map.Objects(17).GraphicID = 1
        .Map.Objects(17).GraphicPosition.x = 182
        .Map.Objects(17).GraphicPosition.y = 3
        .Map.Objects(17).GraphicPosition.Width = 80
        .Map.Objects(17).GraphicPosition.Height = 94
        
        '绿色的植物的图标
        .Map.Objects(18).GraphicID = 1
        .Map.Objects(18).GraphicPosition.x = 276
        .Map.Objects(18).GraphicPosition.y = 3
        .Map.Objects(18).GraphicPosition.Width = 74
        .Map.Objects(18).GraphicPosition.Height = 91
        
        '树桩的图标
        .Map.Objects(19).GraphicID = 1
        .Map.Objects(19).GraphicPosition.x = 211
        .Map.Objects(19).GraphicPosition.y = 94
        .Map.Objects(19).GraphicPosition.Width = 106
        .Map.Objects(19).GraphicPosition.Height = 84
        
        
             '大树的图标
        .Map.Objects(20).GraphicID = 1
        .Map.Objects(20).GraphicPosition.x = 5
        .Map.Objects(20).GraphicPosition.y = 99
        .Map.Objects(20).GraphicPosition.Width = 204
        .Map.Objects(20).GraphicPosition.Height = 222 - 75
        .Map.Objects(20).LayerNum = 1
        
        
          '大树的图标
        .Map.Objects(21).GraphicID = 1
        .Map.Objects(21).GraphicPosition.x = 5
        .Map.Objects(21).GraphicPosition.y = 99
        .Map.Objects(21).GraphicPosition.Width = 204
        .Map.Objects(21).GraphicPosition.Height = 222
        .Map.Objects(21).LayerNum = 0
        
        

        

        '小黄花的图标
        .Map.Objects(22).GraphicID = 1
        .Map.Objects(22).GraphicPosition.x = 216
        .Map.Objects(22).GraphicPosition.y = 179
        .Map.Objects(22).GraphicPosition.Width = 61
        .Map.Objects(22).GraphicPosition.Height = 68
        
        '小红花的图标
        .Map.Objects(23).GraphicID = 1
        .Map.Objects(23).GraphicPosition.x = 188
        .Map.Objects(23).GraphicPosition.y = 417
        .Map.Objects(23).GraphicPosition.Width = 61
        .Map.Objects(23).GraphicPosition.Height = 70
        
        '小草的图标
        .Map.Objects(24).GraphicID = 1
        .Map.Objects(24).GraphicPosition.x = 175
        .Map.Objects(24).GraphicPosition.y = 340
        .Map.Objects(24).GraphicPosition.Width = 81
        .Map.Objects(24).GraphicPosition.Height = 73
        
        '大草的图标
        .Map.Objects(25).GraphicID = 1
        .Map.Objects(25).GraphicPosition.x = 84
        .Map.Objects(25).GraphicPosition.y = 321
        .Map.Objects(25).GraphicPosition.Width = 90
        .Map.Objects(25).GraphicPosition.Height = 88
        
        '宝箱的图标
        .Map.Objects(26).GraphicID = 1
        .Map.Objects(26).GraphicPosition.x = 98
        .Map.Objects(26).GraphicPosition.y = 589
        .Map.Objects(26).GraphicPosition.Width = 76
        .Map.Objects(26).GraphicPosition.Height = 81
        
         '小山的图标
        .Map.Objects(27).GraphicID = 1
        .Map.Objects(27).GraphicPosition.x = 11
        .Map.Objects(27).GraphicPosition.y = 510
        .Map.Objects(27).GraphicPosition.Width = 75
        .Map.Objects(27).GraphicPosition.Height = 86
        
         '蘑菇的图标
        .Map.Objects(28).GraphicID = 1
        .Map.Objects(28).GraphicPosition.x = 412
        .Map.Objects(28).GraphicPosition.y = 352
        .Map.Objects(28).GraphicPosition.Width = 82
        .Map.Objects(28).GraphicPosition.Height = 84
        

        
        '小花1的图标
        .Map.Objects(29).GraphicID = 1
        .Map.Objects(29).GraphicPosition.x = 316
        .Map.Objects(29).GraphicPosition.y = 94
        .Map.Objects(29).GraphicPosition.Width = 64
        .Map.Objects(29).GraphicPosition.Height = 65
        
        '花骨朵的图标
        .Map.Objects(30).GraphicID = 1
        .Map.Objects(30).GraphicPosition.x = 401
        .Map.Objects(30).GraphicPosition.y = 82
        .Map.Objects(30).GraphicPosition.Width = 81
        .Map.Objects(30).GraphicPosition.Height = 89
        




        For a = 0 To 31
            For b = 0 To 31
                .Map.Objects(a).MapPosition(b).x = -1
                .Map.Objects(a).MapPosition(b).y = -1
            Next b
        Next a



        
        .Player(0).Info.Alive = True
        .Player(0).Info.MoveDirection = MoveDown
        .Player(0).Info.MoveSpeed = 0.2
        .Player(0).Info.PositionTimer = 0
        .Player(0).Info.G_Position.x = 6
        .Player(0).Info.G_Position.y = 6
        .Player(0).Info.C_Position.x = 6
        .Player(0).Info.C_Position.y = 6
        .Player(0).Info.MagicBall = False
        .Player(0).Info.MagicBallTimer = 8
        .Player(0).Info.Hp = 100
        .Player(0).Info.Exp = 0
        .Player(0).Info.Mp = 100
        .Player(0).Info.MaxHp = 100
        .Player(0).Info.MaxMp = 100
        .Player(0).Info.Int = 0
        .Player(0).Info.level = 1

        Steps(0).Position.x = .Player(0).Info.C_Position.x
        Steps(0).Position.y = .Player(0).Info.C_Position.y

        

'        .Map.Events(0, 1) = 61
'        .Map.Events(0, 2) = 62
'
'        .Map.Events(0, 3) = 51
'        .Map.Events(0, 4) = 52
'
'        .Map.Events(0, 5) = 59
'        .Map.Events(0, 6) = 63
        
        
'        .Map.Block(0, 1) = 1 '这里用了Event（1）中的图标
'        .Map.Block(0, 2) = 1
'        .Map.Block(0, 3) = 1
'        .Map.Block(3, 1) = 1
'        .Map.Block(3, 2) = 1
'        .Map.Block(3, 3) = 1
'**************************************************************************************
    End With
End Sub

