Attribute VB_Name = "Mdl_Declare"
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and James L. Dean
'E-mail: plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Variable declaration, mainly part of AI

'====================================================================================================================================

 

Option Explicit
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "GDI32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function TextOut Lib "GDI32" Alias "TextOutW" (ByVal hDC As OLE_HANDLE, ByVal x&, ByVal y&, ByVal lpString&, ByVal nCount&) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateFontIndirect Lib "GDI32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public oDict As Object ' As Dictionary
Public oDictLV As Object ' As Dictionary
Public oDictLVs(12) As Object ' As Dictionary

Public Const SCREEN_WIDTH_Current As Integer = 800
Public Const SCREEN_HEIGHT_Current As Integer = 600

Public BackBuffer_Black As Long
Public OldBackBufferDC_Black As Long

Public BackBuffer_Other As Long
Public OldBackBufferDC_Other As Long

Public BackBuffer As Long
Public BackBufferBmp As Long
Public OldBackBufferDC As Long

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 50
        x As Integer
        y As Integer
        Text As String
        Clr As Long
        Mdl As Byte
        BkClr As Long
End Type


Public SpriteNameFont(255) As LOGFONT
Public NewFont(255) As Long
Public OldFont(255) As Long
Public WScreen_PrintTmr As Long


Public KeywordFile As String
Public ExtraSearchFile As String


Public Enum tMUD_ResponseType
    Statement = 1
    Question = 2
    Either = 3
    Other = 4
    ExtraSearch = 9
End Enum

Public Type tMUD_Keyword
    KeywordText As String
    KeywordNo As Integer
    KeywordFile As String
    KeywordOrigin As String
End Type

Public Type tMUD_Response
    ResponseText As String      'The response being used
    ResponseSubject As String   'The subject of the current response
    ResponseNo As Integer
    ResponseType As Byte
    Question  As Boolean
    ResponseAction As String
End Type

Public Type tMUD_Files
    Type As String
    Name As String
    Count As Integer
    Response() As tMUD_Response
    Keyword() As tMUD_Keyword
End Type

Public Type tMUD_SimpleResponse
    sReply  As String
    sAction As String
    sQuestion As Boolean
End Type

Public Type tMUD_Word_Frequency
    Words As String
    Frequency As Integer '�ȼ�
    ID As Long
End Type

Public Type tMUD_Word_Words '�ʻ㲿��
    Words As String
    Phonetic As String '����
    Explain As String '����
    Frequency As Integer
    MemonyHelpID As String
    EEDef As String
    Synonyms As String
    UID As Integer
End Type

Public Type tMUD_IrregularVerb '�����򶯴�
    InfinitiveWord As String '����ʱ
    PastTenseWord() As String '��ȥʽ
    PastParticipleWord() As String '��ȥ�ִ�
    Frequency As String  '�ȼ�
    WordsUID As Integer
End Type

Public Type tMUD_Treasure
    strTreasure As String
    strGuard As String
    nGuardRoom As String
    nTreasureRoom As Integer
    nWeaponRoom As Integer
    strWeapon As String
End Type

Public Type tMUD_Rooms
    strDescription  As String
    bVisited As Boolean
    bConnected(4, 2) As Boolean
    bRoomUsed As Boolean
End Type
    
Public Type my_Files
    Type As String
    Name As String
    Count As Integer
End Type

Public Type tDialogue
    ID As Integer
    Step As String
    Text(2) As String
    Type As Byte
End Type

Public Type LogEntry
    sName As String
    sGenerated As String
    sDescription As String
End Type


Public Type tMUD_Word_Addition
    Words As String
    ID As Long
    Addition As String
End Type

Public Type tMUD_Correlation_Index
    AllCorrelation As String
    AllName As String
End Type

Type tMUD_Correlation
    WordsID(30) As Integer  '��¼��صĵ���
    Correlation(30) As Integer '��¼ÿһ�����ʵĹ������,����12|13|14|,�����,2��,3��,4�ŵ��ʶ�������1�ŵ�����
    WordsCount As Byte
    Index As Byte
    SynonymWordExplain As String
End Type

Public Type tMUD_Factor  '�ʻ㲿��
    PrefixID As String
    RootID As String
    SuffixID As String
    UID As Integer
End Type

Public Type tMUD_Factor_Index
    FactorString As String
    EnglishMeanings As String
    NativeMeanings As String
    Class As String
End Type

Public Type tMUD_User_Words
    TypeCount As Long
    CorrectRate As Integer
End Type

Public Type tMUD_User
    Words() As tMUD_User_Words
    level(10) As Long
    HisScore(10) As Long '��ʷ������¼
    RegCount As Long     '��¼����
End Type

Public Type tMUD_Practice '�и��ʻ�
    Material As String '����ʱ
    KeyWordsSentence(2) As String
    KeyWords(2) As String
    Options(2, 4) As String '��ȥʽ
    Answer(2) As String '��ȥ�ִ�
End Type

Public MUD_Practice() As tMUD_Practice
Public MUD_Practice_Start As Boolean
Public MUD_Practice_ID As Integer
Public MUD_UserName As String
Public MUD_UserFiles As String
Public MUD_UserWords() As String
Public MUD_UserWordsLevel() As Long
Public MUD_UserLevelWordsID() As Long

Public MUD_User As tMUD_User
 
'===��Ϸ===
Public MUD_Word_Words() As tMUD_Word_Words
Public MUD_IrregularVerb() As tMUD_IrregularVerb
Public MUD_Word_Addition() As tMUD_Word_Addition
Public MUD_Word_Frequency() As tMUD_Word_Frequency
Public MUD_Word_Member() As String

Public MUD_Correlation_Index() As tMUD_Correlation_Index
Public MUD_CorrelationWords() As tMUD_Correlation
Public MUD_Factor_Index() As tMUD_Factor_Index
Public MUD_FactorWords() As tMUD_Factor


'====����=====
Public MUD_Keyword() As tMUD_Keyword
Public MUD_Response() As tMUD_Response


Public MUD_Files() As tMUD_Files
Public MUDsKey As tMUD_Keyword
Public LastReply As String
Public LastResponse As tMUD_Response
Public LastKeyword As tMUD_Keyword

'===�Ի�==
Public MUD_DialogueROM() As tDialogue
Public MUD_ValROM() As String


'===��Ϸ===
Public MUD_Treasure() As tMUD_Treasure
Public MUD_Rooms() As tMUD_Rooms
Public ReturnFunction As String

Public TriggerKey As Boolean 'Clears text when mouse moves over
Public BreakType As String
Public AboutMe As Boolean


Public bEuclidean As Boolean
Public nDimensions As Long
Public nTop As String
Public nBottom As String
Public bDirectionFound As Boolean
Public bDirectionUsed(255, 4, 2) As Boolean
Public bInitialized As Boolean
Public bPathFound As Boolean
Public bTreasureCarried As Boolean
Public bWeaponRoomFound As Boolean
Public bWidthsFound As Boolean
Public dblScore As Double
Public nCell(15, 15, 15, 15) As Long
Public nCoordinate(4) As Long
Public nCoordinateNext(4) As Long
Public nDimension1 As Long
Public nDimension2 As Long
Public nDirection1 As Long
Public nDirection2 As Long
Public nDirectionsPossible As Long
Public nDirectionsUsed(255) As Long
Public nMaxWidth As Long
Public nMoves As Long
Public nRoom1 As Long
Public nRoom2 As Long
Public nRooms As Long
Public nScore As Long
Public nTCoordinate As Long
Public nTreasure1 As Long
Public nTreasure2 As Long
Public nTreasuresCarried As Long
Public nTreasuresRecovered As Long
Public nTrial As Long
Public nVisited As Long
Public nVolume As Long
Public nWayOutDimension(255) As Long
Public nWayOutDirection(255) As Long
Public nWayOutHead As Long
Public nWayOutPtr As Long
Public nWidth(4) As Long
Public nXCoordinate As Long
Public nYCoordinate As Long
Public nZCoordinate As Long
Public strLine As String
Public strTreasures As String
Public strWayOut As String
Public nTreasures As Integer


Public DebugMode As Boolean
Public Speaking As Boolean
'Systems
Public ConsoleRunning As Boolean

Public Username As String
Public Gender As String


 