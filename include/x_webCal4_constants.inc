<%
' ADO Constants ----------------------------------------------------------
	
' cursors
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

' cursor location
Const adUseServer = 2
Const adUseClient = 3

' locks
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

' commands
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200
Const adExecuteNoRecords = &H00000080

' filters
Const adFilterNone = 0

' Field Constants --------------------------------------------------------
' arGroups
Const g_GROUP_ID = 0
Const g_GROUP_NAME = 1
Const g_VISIBLE = 2
Const g_GROUP_ACCESS = 3

' arScopes
Const g_SCOPE_ID = 0
Const g_SCOPE_NAME = 1
' g_VISIBLE defined above

' event fields
Const g_EVENT_ID = 0
Const g_EVENT_TITLE = 1
Const g_EVENT_RECUR = 2
Const g_EVENT_COLOR = 3
Const g_TIME_START = 4
Const g_TIME_END = 5
Const g_EVENT_DATE = 6
Const g_EVENT_DESC = 7
Const g_EVENT_SKIP_WE = 8

' access levels
Const g_NO_ACCESS = 0
Const g_READ_ACCESS = 1
Const g_ADD_ACCESS = 2
Const g_EDIT_ACCESS = 3
Const g_MGR_ACCESS = 4
Const g_ADMIN_ACCESS = 5

' event scopes
Const g_PRIVATE = 1
Const g_GROUP = 2
Const g_PUBLIC = 3

' date properies
Const g_THIS_DATE = 0		' selected date
Const g_FIRST_DATE = 1		' first date in displayed range
Const g_LAST_DATE = 2		' last date in displayed range
Const g_PREV_DATE = 3		' date from previous month / week / day
Const g_NEXT_DATE = 4		' date in next month / week / day
Const g_FIRST_DAY = 5		' first day number of range
Const g_LAST_DAY = 6		' last day number of range

' week and day grid segment properties
Const g_SEG_MINS = 0		' number of minutes per table cell (segment)
Const g_SEG_START = 1		' day start time measured in segments
Const g_SEG_END = 2			' day end measured in segments
Const g_WEEK = 0
Const g_DAY = 1
Const g_MONTH = 2

' application settings
Const g_COLOR_ID = 0
Const g_LCID = 1
Const g_PASS_LENGTH = 2
Const g_PASS_LIFE = 3
Const g_CACHE_SIZE = 4
Const g_USE_CACHE = 5
Const g_EASY_EDIT = 6
Const g_SHOW_WEEKEND = 7
Const g_DEFAULT_SEG_MINS = 8
Const g_DEFAULT_SEG_START = 9
Const g_DEFAULT_SEG_END = 10
Const g_START_PAGE = 11

' symbol constants
Const g_CHAR_PAPERCLIP = 0
Const g_CHAR_LEFT_ARROW = 1
Const g_CHAR_RIGHT_ARROW = 2
Const g_CHAR_UP_ARROW = 3
Const g_CHAR_DOWN_ARROW = 4
Const g_CHAR_MAG_GLASS = 5
Const g_CHAR_INFO = 6
Const g_CHAR_RECUR = 7
Const g_CHAR_CLOSE = 8
Const g_CHAR_OPEN = 9
Const g_CHAR_QUESTION = 10
Const g_CHAR_CLOCK_1 = 11
Const g_CHAR_CLOCK_2 = 12
Const g_CHAR_CLOCK_3 = 13
Const g_CHAR_CLOCK_4 = 14
Const g_CHAR_CLOCK_5 = 15
Const g_CHAR_CLOCK_6 = 16
Const g_CHAR_CLOCK_7 = 17
Const g_CHAR_CLOCK_8 = 18
Const g_CHAR_CLOCK_9 = 19
Const g_CHAR_CLOCK_10 = 20
Const g_CHAR_CLOCK_11 = 21
Const g_CHAR_CLOCK_12 = 22
Const g_CHAR_LOCK = 23
Const g_CHAR_UNLOCK = 24
Const g_CHAR_XBOX = 25
Const g_CHAR_CHECKBOX = 26
Const g_FONT_CHAR = 0
Const g_FONT_FACE = 1

' query property
Const g_sNO_EVENTS = "noevents"

' cache properties
Const g_CACHE_SQL = 0
Const g_CACHE_HTML = 1
Const g_CACHE_GRID = 2
Const g_CACHE_START_DATE = 3
Const g_CACHE_END_DATE = 4
Const g_CACHE_EXPIRE_DATE = 5
%>