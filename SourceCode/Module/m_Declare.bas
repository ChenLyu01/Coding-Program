Attribute VB_Name = "m_Declare"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************



'====================================================================Function description====================================================

' Declare all variables

'====================================================================================================================================


Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long

 
Public SoftwareRegPSD As String
Public User_Code As String
Public User_Codeing As Boolean

'User registration information
Public System_Sm_02(4) As String
Public System_Sm_03(4) As String


 'Maximum width of chat bubbles
Public Const BubbleMaxWidth As Long = 140
Public Const SpriteFont_MaxLineLen As Long = 600

Private Type VFH
    BitmapWidth As Long         'Bitmap size
    BitmapHeight As Long
    CellWidth As Long           '
    CellHeight As Long
    BaseCharOffset As Byte      'Start word
    CharWidth(0 To 255) As Byte 'Actual width of characters
End Type

Public Type CustomFont
    HeaderInfo As VFH           'Title
    RowPitch As Integer         'Number of characters per line
    RowFactor As Single         'Texture width occupied by each character
    ColFactor As Single         '
    CharHeight As Byte          '
End Type


Type m_Position
    x As Integer
    y As Integer
    Enable As Boolean
End Type

Type m_RECT
    Top           As Integer
    bottom        As Integer
    Left          As Integer
    Right         As Integer
End Type

Type m_WandH
    Width As Integer
    Height As Integer
End Type

Type m_GraphicPosition
    x As Integer
    y As Integer
    Width As Integer
    Height As Integer
End Type

Type m_Npcplay      'Types that can be deleted
    C_Position    As m_Position '
    G_Position    As m_Position '
    Next_Position As m_Position '
    Roading(100)  As m_Position    '
    RoadData      As Integer '
    RodeTry       As Boolean
    RoadXY        As Boolean
    G_Event(1024) As Byte  '
    MoveDirection As Byte
    Alive         As Boolean '
    SleepTimer    As Integer '
    EventTimer    As Integer '
    EventBreak    As Boolean '
    PositionTimer As Single '
    MoveTimer     As Integer '
    Agi           As Integer '
    Pic           As Integer
    MoveSpeed     As Single  '
    Hp            As Integer '
    MaxHp         As Integer '
    Def           As Integer '
    Str           As Integer '
    Int           As Integer '
    Men           As Integer '
    Att           As Integer '
    AttSpeed      As Integer '
    level         As Integer '
    Exp           As Long    '
    Mp            As Integer '
    MaxMp         As Integer '
    Width         As Integer '
    Height        As Integer '
    HeadPic       As Integer '
    MagicBall     As Boolean
    MagicBallTimer As Integer
End Type

Type m_Clock
    myHour As Byte
    myMinute As Byte
End Type

Type m_MapGraphicPosition
    GraphicID As Byte
    GraphicPosition As m_GraphicPosition
    MapPosition(32) As m_Position
    LayerNum As Byte
    GraphicName As String
End Type

Type m_Buffer
    BackBuffer As Long
    BackBufferBmp As Long
    OldBackBufferDC As Long
    OldTilesetBmpDC(16) As Long
    TileSetBmp(16) As Long
End Type

Type m_Windows
    GraphicID As Byte
    Enable As Boolean
    Visible As Boolean
    LoadPosition As m_GraphicPosition
    DrawPosition As m_Position
End Type

Type m_BlockTile
    GraphicID As Byte
    GraphicPosition As m_GraphicPosition
End Type


Type m_Event
    m_Name As String
    m_Description As String
    PicPosition As m_Position
End Type

Type m_Effect
    GraphicID As Byte
    Visible(2) As Boolean
    FrameRun(2) As Boolean
    Timer(2) As Integer
    FrameCount As Byte
    Matrix As m_Position
    MatrixGraphic As m_WandH
    DrawPosition(2) As m_Position
    LoadPosition As m_GraphicPosition
End Type

Type m_steps
    GraphicID As Byte
    Arrow(4) As m_GraphicPosition
    Direction As Integer
    Position As m_Position
    Visible As Boolean
End Type

Type m_Map
    sName As String
    ID As Integer
    Tiles_Posi(3) As m_Position
    TileID As Byte
    TilesInfo As m_GraphicPosition
    Tile_Map As m_WandH
    Tile_Object As m_WandH
    Tiles(5) As m_MapGraphicPosition
    Letters(4, 26) As m_MapGraphicPosition
    Letter(52) As m_MapGraphicPosition
    BlockTilesInfo(16) As m_BlockTile
    BlockTilesID As Byte
    Objects(32) As m_MapGraphicPosition
    Events(16, 14) As Byte
    Block(16, 14) As Byte
    BlockTiles(16, 14) As Byte
    GameEvents(64) As m_Event
    Windows(4) As m_Windows
    Effect(8) As m_Effect
End Type

Type m_GraphicFiles
    FilesName As String
    BlackPicPixel As m_GraphicPosition
End Type
        
Type m_Player
    GraphicID As Byte
    Tile As m_WandH
    Pic As m_Position
    Pic_Temp(2) As m_Position
    Info As m_Npcplay
    mName As String
End Type

Public Type m_Font
    x As Integer  'Font X coordinates
    y As Integer   'Font Y coordinates
    G_String As String 'Font content
    G_MaxLineLen As Integer 'Font content width
    G_WordWraped As Boolean
    Visiable As Boolean
End Type

Public Type m_Story
    m_Name As String
    Text As String
End Type

Type m_Code
    mName As String
    Text As String
    Autorun As Boolean
    runFlag As Boolean
    Order As String
End Type

Type m_Graphic
    EffectTimer(8) As Single
    Posi  As m_Position
    RECT As m_RECT
    Screen As m_WandH
    Buffer As m_Buffer
    Map As m_Map
    GraphicFiles(16) As m_GraphicFiles
    Player(2) As m_Player
    SpriteFont(100) As m_Font
    WindowsFont(100) As m_Font
    Font_Default As CustomFont   'Default font
    Clock As m_Clock
    Code As m_Code
    AILoaded As Boolean
End Type

Type m_FilePath
    Graphics As String
    Code As String
    CourseMap As String
    CodeName As String
    Story As String
End Type

Type m_Switch
    Object As Boolean
    Block As Boolean
    Event As Boolean
    Player As Boolean
    Effect As Boolean
    Pathway As Boolean
    Debug As Boolean
    Letters As Boolean
    Timer As Boolean
    Steps As Boolean
    PathwayWithBlock As Boolean
    Talk As Boolean
End Type



Public this_AI_text As m_Position

Public Steps(100) As m_steps

 
 

Public this_Graphic As m_Graphic
Public this_FilePath As m_FilePath
Public this_Switch As m_Switch

Public Const MoveLeft = 1
Public Const MoveRight = 2
Public Const MoveUp = 3
Public Const MoveDown = 0
Public CurrentPlayerID As Byte
Public CurrentSteps As Byte
Public CurrentEventTimer As Integer
Public CurrentGameEvents As Byte
Public EventBreak As Boolean
 

