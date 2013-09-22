Option Strict Off
Option Explicit On
Friend Class ctlButton
	Inherits System.Windows.Forms.UserControl
	Public Event MousePointerChange()
	Public Event TooltipIconChange()
	Public Event CaptionChange()
	Public Event FontHighlightColorChange()
	Public Event ToolTipForeColorChange()
	Public Event ToolTipBackColorChange()
	Public Event IconChange()
	Public Event FontColorChange()
	Public Event ShowFocusChange()
	Public Event MaskColorChange()
	Public Event ToolTipTextChange()
	Public Event ValueChange()
	Public Event IconAlignChange()
	Public Event IconSizeChange()
	Public Event TooltipTitleChange()
	Public Event UseCustomColorsChange()
	Public Event HighlightColorChange()
	Public Event NonThemeStyleChange()
	Public Event EnabledChange()
	Public Event UseMaskColorChange()
	Public Event ToolTipTypeChange()
	Public Event CaptionAlignChange()
	Public Event ButtonTypeChange()
	Public Event RoundedBordersByThemeChange()
	Public Event FontChange()
	Public Event StyleChange()
	Public Event BackColorChange()
	Private Const strCurrentVersion As String = "3.5.0"
	Private Const COLOR_BTNFACE As Integer = 15
	Private Const COLOR_BTNSHADOW As Integer = 16
	Private Const COLOR_BTNTEXT As Integer = 18
	Private Const COLOR_HIGHLIGHT As Integer = 13
	Private Const COLOR_WINDOW As Integer = 5
	Private Const COLOR_INFOTEXT As Integer = 23
	Private Const COLOR_INFOBK As Integer = 24
	Private Const BDR_RAISEDOUTER As Integer = &H1s
	Private Const BDR_SUNKENOUTER As Integer = &H2s
	Private Const BDR_RAISEDINNER As Integer = &H4s
	Private Const BDR_SUNKENINNER As Integer = &H8s
	Private Const EDGE_RAISED As Boolean = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
	Private Const EDGE_SUNKEN As Boolean = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
	Private Const BF_LEFT As Integer = &H1s
	Private Const BF_TOP As Integer = &H2s
	Private Const BF_RIGHT As Integer = &H4s
	Private Const BF_BOTTOM As Integer = &H8s
	Private Const BF_RECT As Integer = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
	Private Const WM_MOUSEMOVE As Integer = &H200s
	Private Const WM_MOUSELEAVE As Integer = &H2A3s
	Private Const WM_THEMECHANGED As Integer = &H31As
	Private Const WM_SYSCOLORCHANGE As Integer = &H15s
	Private Const WM_USER As Integer = &H400s
	Private Const GWL_STYLE As Integer = -16
	Private Const WS_CAPTION As Integer = &HC00000
	Private Const WS_THICKFRAME As Integer = &H40000
	Private Const WS_MINIMIZEBOX As Integer = &H20000
	Private Const SWP_REFRESH As Integer = (&H1s Or &H2s Or &H4s Or &H20s)
	Private Const SWP_NOACTIVATE As Integer = &H10s
	Private Const SWP_NOMOVE As Integer = &H2s
	Private Const SWP_NOSIZE As Integer = &H1s
	Private Const SWP_SHOWWINDOW As Integer = &H40s
	Private Const HWND_TOPMOST As Integer = -&H1s
	Private Const CW_USEDEFAULT As Integer = &H80000000
	Private Const ALTERNATE As Integer = 1
	Private Const TTS_NOPREFIX As Integer = &H2s
	Private Const TTF_CENTERTIP As Integer = &H2s
	Private Const TTM_ADDTOOLA As Integer = (WM_USER + 4)
	Private Const TTM_DELTOOLA As Integer = (WM_USER + 5)
	Private Const TTM_SETTIPBKCOLOR As Integer = (WM_USER + 19)
	Private Const TTM_SETTIPTEXTCOLOR As Integer = (WM_USER + 20)
	Private Const TTM_SETTITLE As Integer = (WM_USER + 32)
	Private Const TTS_BALLOON As Integer = &H40s
	Private Const TTS_ALWAYSTIP As Integer = &H1s
	Private Const TTF_SUBCLASS As Integer = &H10s
	Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
	Private Const ALL_MESSAGES As Integer = -1
	Private Const GMEM_FIXED As Integer = 0
	Private Const GWL_WNDPROC As Integer = -4
	Private Const PATCH_04 As Integer = 88
	Private Const PATCH_05 As Integer = 93
	Private Const PATCH_08 As Integer = 132
	Private Const PATCH_09 As Integer = 137
	Private Structure Point
		Dim X As Integer
		Dim Y As Integer
	End Structure
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Top_Renamed As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		'UPGRADE_NOTE: Bottom was upgraded to Bottom_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Bottom_Renamed As Integer
	End Structure
	Private Structure tSubData
		Dim hWnd As Integer
		Dim nAddrSub As Integer
		Dim nAddrOrig As Integer
		Dim nMsgCntA As Integer
		Dim nMsgCntB As Integer
		Dim aMsgTblA() As Integer
		Dim aMsgTblB() As Integer
	End Structure
	
	''Tooltip Window Types
	Private Structure TOOLINFO
		Dim lSize As Integer
		Dim lFlags As Integer
		Dim lHwnd As Integer
		Dim lId As Integer
		Dim lpRect As RECT
		Dim hInstance As Integer
		Dim lpStr As String
		Dim lParam As Integer
	End Structure
	
	'Private Type OSVERSIONINFOEX    'OS Version
	'  dwOSVersionInfoSize             As Long
	'  dwMajorVersion                  As Long
	'  dwMinorVersion                  As Long
	'  dwBuildNumber                   As Long
	'  dwPlatformId                    As Long
	'  szCSDVersion                    As String * 128
	'  wServicePackMajor               As Integer
	'  wServicePackMinor               As Integer
	'  wSuiteMask                      As Integer
	'  wProductType                    As Byte
	'  wReserved                       As Byte
	'End Type
	
	Enum isbStyle 'Styles
		isbNormal = &H0s
		isbSoft = &H1s
		isbFlat = &H2s
		isbJava = &H3s
		isbOfficeXP = &H4s
		isbWindowsXP = &H5s
		isbWindowsTheme = &H6s
		isbPlastik = &H7s
		isbGalaxy = &H8s
		isbKeramik = &H9s
		isbMacOSX = &HAs
	End Enum
	
	Enum isbButtonType
		isbButton = &H0s
		isbCheckBox = &H1s
	End Enum
	
	Enum isbAlign
		isbCenter = &H0s
		isbleft = &H1s
		isbRight = &H2s
		isbTop = &H3s
		isbbottom = &H4s
	End Enum
	
	Private Enum isState
		statenormal = &H1s
		stateHot = &H2s
		statePressed = &H3s
		statedisabled = &H4s
		stateDefaulted = &H5s
	End Enum
	
	'Private Type MSG             'Windows Message Structure
	'  hwnd                            As Long
	'  message                         As Long
	'  wParam                          As Long
	'  lParam                          As Long
	'  time                            As Long
	'  pt                              As POINT
	'End Type
	'
	'Private Type tagTRACKMOUSEEVENT
	'  cbSize                          As Long
	'  dwFlags                         As Long
	'  hwndTrack                       As Long
	'  dwHoverTime                     As Long
	'End Type
	'
	'Private Type TRIVERTEX          'For gradient Drawing
	'  x                               As Long
	'  Y                               As Long
	'  Red                             As Integer
	'  Green                           As Integer
	'  Blue                            As Integer
	'  Alpha                           As Integer
	'End Type
	
	'Private Type GRADIENT_RECT
	'  UpperLeft                       As Long
	'  LowerRight                      As Long
	'End Type
	
	'Private Type GRADIENT_TRIANGLE
	'  Vertex1                         As Long
	'  Vertex2                         As Long
	'  Vertex3                         As Long
	'End Type
	
	'Private Type DRAWTEXTPARAMS 'Required for DrawText
	'  cbSize                          As Long
	'  iTabLength                      As Long
	'  iLeftMargin                     As Long
	'  iRightMargin                    As Long
	'  uiLengthDrawn                   As Long
	'End Type
	
	'Private Type BLENDFUNCTION  'Required for Alphablend API
	'  BlendOp                         As Byte
	'  BlendFlags                      As Byte
	'  SourceConstantAlpha             As Byte
	'  AlphaFormat                     As Byte
	'End Type
	
	'Private Type RGBQUAD
	'  rgbBlue                     As Byte
	'  rgbGreen                    As Byte
	'  rgbRed                      As Byte
	'  rgbReserved                 As Byte
	'End Type
	
	'Private Type BITMAPINFOHEADER
	'  biSize                      As Long
	'  biWidth                     As Long
	'  biHeight                    As Long
	'  biPlanes                    As Integer
	'  biBitCount                  As Integer
	'  biCompression               As Long
	'  biSizeImage                 As Long
	'  biXPelsPerMeter             As Long
	'  biYPelsPerMeter             As Long
	'  biClrUsed                   As Long
	'  biClrImportant              As Long
	'End Type
	
	'Private Type BITMAPINFO
	'  bmiHeader                   As BITMAPINFOHEADER
	'  bmiColors                   As RGBQUAD
	'End Type
	
	'Private Type UxTheme        'Imported from a Cls File from VBAccelerator.com
	'  sClass As String        'And edited to keep the control in a single file.
	'  Part As Long            'I didn't used all the constant definitions where used
	'  State As Long           'in the original file, cuz I don't need them all
	'  hdc As Long             'But I added some others I need, like text offset
	'  hwnd As Long            'properties and UseTheme, to Detect If the draw was
	'  Left As Long            'succesfull or not, and then use classic windows Style
	'  Top As Long             'Drawing.
	'  Width As Long           'All the credits about the usage of UxTheme.dll defined on
	'  Height As Long          'cUxTheme.cls go for Steve at www.vbaccelerator.com
	'  Text As String
	'  TextAlign As Long 'DrawTextFlags
	'  IconIndex As Long
	'  hIml As Long
	'  RaiseError As Boolean
	'  UseThemeSize As Boolean
	'  UseTheme As Boolean
	'  TextOffset As Long
	'  RightTextOffset  As Long
	'End Type
	
	Private Structure ICONINFO
		Dim fIcon As Integer
		Dim xHotspot As Integer
		Dim yHotspot As Integer
		Dim hbmMask As Integer
		Dim hbmColor As Integer
	End Structure
	
	Private Structure BITMAP
		Dim bmType As Integer 'LONG   // Specifies the bitmap type. This member must be zero.
		Dim bmWidth As Integer 'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
		Dim bmHeight As Integer 'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
		Dim bmWidthBytes As Integer 'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
		Dim bmPlanes As Short 'WORD   // Specifies the count of color planes.
		Dim bmBitsPixel As Short 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
		Dim bmBits As Integer 'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
	End Structure
	
	'*************************************************************
	'
	'   Required Enums
	'
	'*************************************************************
	'Private Enum DrawTextAdditionalFlags
	'  DTT_GRAYED = &H1           '// draw a grayed-out string
	'End Enum
	
	'Private Enum THEMESIZE
	'  TS_MIN             '// minimum size
	'  TS_TRUE            '// size without stretching
	'  TS_DRAW            '// size that theme mgr will use to draw part
	'End Enum
	
	Private Enum eMsgWhen
		MSG_AFTER = 1 'Message calls back after the original (previous) WndProc
		MSG_BEFORE = 2 'Message calls back before the original (previous) WndProc
		MSG_BEFORE_AND_AFTER = eMsgWhen.MSG_AFTER Or eMsgWhen.MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
	End Enum
	
	Private Enum TRACKMOUSEEVENT_FLAGS
		TME_HOVER = &H1
		TME_LEAVE = &H2
		TME_QUERY = &H40000000
		TME_CANCEL = &H80000000
	End Enum
	
	Private Structure TRACKMOUSEEVENT_STRUCT
		Dim cbSize As Integer
		Dim dwFlags As Integer 'TRACKMOUSEEVENT_FLAGS
		Dim hwndTrack As Integer
		Dim dwHoverTime As Integer
	End Structure
	
	Public Enum ttIconType
		TTNoIcon = 0
		TTIconInfo = 1
		TTIconWarning = 2
		TTIconError = 3
	End Enum
	
	Public Enum ttStyleEnum
		TTStandard
		TTBalloon
	End Enum
	
	'Private Enum GRADIENT_FILL_RECT
	'  FillHor = GRADIENT_FILL_RECT_H
	'  FillVer = GRADIENT_FILL_RECT_V
	'End Enum
	
	'Private Enum GRADIENT_TO_CORNER
	'  All
	'  TopLeft
	'  TopRight
	'  BottomLeft
	'  BottomRight
	'End Enum
	
	'Private Enum CRADIENT_DIRECTION
	'  DirectionSlash
	'  DirectionBackSlash
	'End Enum
	
	Private Structure RGBQUAD
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
		Dim rgbReserved As Byte
	End Structure
	
	Private Structure BITMAPINFOHEADER
		Dim biSize As Integer
		Dim biWidth As Integer
		Dim biHeight As Integer
		Dim biPlanes As Short
		Dim biBitCount As Short
		Dim biCompression As Integer
		Dim biSizeImage As Integer
		Dim biXPelsPerMeter As Integer
		Dim biYPelsPerMeter As Integer
		Dim biClrUsed As Integer
		Dim biClrImportant As Integer
	End Structure
	
	Private Structure BITMAPINFO
		Dim bmiHeader As BITMAPINFOHEADER
		Dim bmiColors As RGBQUAD
	End Structure
	
	Private Enum DrawTextFlags
		DT_TOP = &H0s
		DT_LEFT = &H0s
		DT_CENTER = &H1s
		DT_RIGHT = &H2s
		DT_VCENTER = &H4s
		DT_BOTTOM = &H8s
		DT_WORDBREAK = &H10s
		DT_SINGLELINE = &H20s
		DT_EXPANDTABS = &H40s
		DT_TABSTOP = &H80s
		DT_NOCLIP = &H100s
		DT_EXTERNALLEADING = &H200s
		DT_CALCRECT = &H400s
		DT_NOPREFIX = &H800s
		DT_INTERNAL = &H1000s
		DT_EDITCONTROL = &H2000s
		DT_PATH_ELLIPSIS = &H4000s
		DT_END_ELLIPSIS = &H8000s
		DT_MODIFYSTRING = &H10000
		DT_RTLREADING = &H20000
		DT_WORD_ELLIPSIS = &H40000
		DT_NOFULLWIDTHCHARBREAK = &H80000
		DT_HIDEPREFIX = &H100000
		DT_PREFIXONLY = &H200000
	End Enum
	
	
	Private Structure RGBTRIPLE
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
	End Structure
	
	'*************************************************************
	'
	'   Required API Declarations
	'
	'*************************************************************
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Declare Function SetWindowText Lib "user32.dll"  Alias "SetWindowTextA"(ByVal hWnd As Integer, ByVal lpString As String) As Integer
	Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Integer
	Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Integer) As Integer
	Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Integer
	'UPGRADE_WARNING: Structure TRACKMOUSEEVENT_STRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Integer
	'UPGRADE_WARNING: Structure TRACKMOUSEEVENT_STRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function TrackMouseEventComCtl Lib "Comctl32"  Alias "_TrackMouseEvent"(ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Integer
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	'Private Declare Function GetModuleHandle _
	'Lib "kernel32" _
	'Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	'Private Declare Function SetCapture _
	'Lib "user32" (ByVal hwnd As Long) As Long
	'Private Declare Function ReleaseCapture _
	'Lib "user32" () As Long
	'Private Declare Function LoadCursor _
	'Lib "user32" _
	'Alias "LoadCursorA" (ByVal hInstance As Long, _
	'ByVal lpCursorName As Long) As Long
	'Private Declare Function SetCursor _
	'Lib "user32" (ByVal hCursor As Long) As Long
	Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	'Private Declare Function StretchBlt _
	'Lib "gdi32" (ByVal hdc As Long, _
	'ByVal x As Long, _
	'ByVal Y As Long, _
	'ByVal nWidth As Long, _
	'ByVal nHeight As Long, _
	'ByVal hSrcDC As Long, _
	'ByVal XSrc As Long, _
	'ByVal YSrc As Long, _
	'ByVal nSrcWidth As Long, _
	'ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CopyRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Integer, ByVal Y As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OffsetRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Integer, ByVal Y As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawText Lib "user32"  Alias "DrawTextA"(ByVal hdc As Integer, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	'Private Declare Function RedrawWindow _
	'Lib "user32" (ByVal hwnd As Long, _
	'lprcUpdate As RECT, _
	'ByVal hrgnUpdate As Long, _
	'ByVal fuRedraw As Long) As Long
	'Private Declare Function AlphaBlend _
	'Lib "msimg32" (ByVal hdcDest As Long, _
	'ByVal nXOriginDest As Long, _
	'ByVal nYOriginDest As Long, _
	'ByVal nWidthDest As Long, _
	'ByVal hHeightDest As Long, _
	'ByVal hdcSrc As Long, _
	'ByVal nXOriginSrc As Long, _
	'ByVal nYOriginSrc As Long, _
	'ByVal nWidthSrc As Long, _
	'ByVal nHeightSrc As Long, ByVal blendFunc As Long) As Boolean
	'Private Declare Function TransparentBlt _
	'Lib "msimg32" (ByVal hdcDest As Long, _
	'ByVal nXOriginDest As Long, _
	'ByVal nYOriginDest As Long, _
	'ByVal nWidthDest As Long, _
	'ByVal hHeightDest As Long, _
	'ByVal hdcSrc As Long, _
	'ByVal nXOriginSrc As Long, _
	'ByVal nYOriginSrc As Long, _
	'ByVal nWidthSrc As Long, _
	'ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Boolean
	'Private Declare Function GetVersion _
	'Lib "kernel32" () As Long
	Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Integer, ByVal pszClassList As Integer) As Integer
	'Private Declare Function CloseThemeData _
	'Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Integer, ByVal lhdc As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef pClipRect As RECT) As Integer
	'Private Declare Function DrawThemeParentBackground _
	'Lib "uxtheme.dll" (ByVal hwnd As Long, _
	'ByVal hdc As Long, _
	'prc As RECT) As Long
	'Private Declare Function GetThemeBackgroundContentRect _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal hdc As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long, _
	'pBoundingRect As RECT, _
	'pContentRect As RECT) As Long
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Integer, ByVal hdc As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef pRegion As Integer) As Integer
	'Private Declare Function DrawThemeText _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal hdc As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long, _
	'ByVal pszText As Long, _
	'ByVal iCharCount As Long, _
	'ByVal dwTextFlag As Long, _
	'ByVal dwTextFlags2 As Long, _
	'pRect As RECT) As Long
	'Private Declare Function DrawThemeIcon _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal hdc As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long, _
	'pRect As RECT, _
	'ByVal hIml As Long, _
	'ByVal iImageIndex As Long) As Long
	'Private Declare Function GetThemePartSize _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal hdc As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long, _
	'prc As RECT, _
	'ByVal eSize As THEMESIZE, _
	'psz As SIZE) As Long
	'Private Declare Function GetThemeTextExtent _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal hdc As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long, _
	'ByVal pszText As Long, _
	'ByVal iCharCount As Long, _
	'ByVal dwTextFlags As DrawTextFlags, _
	'pBoundingRect As RECT, _
	'pExtentRect As RECT) As Long
	'Private Declare Function IsThemePartDefined _
	'Lib "uxtheme.dll" (ByVal hTheme As Long, _
	'ByVal iPartId As Long, _
	'ByVal iStateId As Long) As Long
	'Private Declare Function ImageList_GetImageRect _
	'Lib "comctl32.dll" (ByVal hIml As Long, _
	'ByVal i As Long, _
	'prcImage As RECT) As Long
	Private Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	'Private Declare Function GetModuleFileName _
	'Lib "kernel32" _
	'Alias "GetModuleFileNameA" (ByVal hModule As Long, _
	'ByVal lpFileName As String, _
	'ByVal nSize As Long) As Long
	'Private Declare Function GetTickCount _
	'Lib "kernel32" () As Long
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	'Private Declare Function CreateBitmap _
	'Lib "gdi32" (ByVal nWidth As Long, _
	'ByVal nHeight As Long, _
	'ByVal nPlanes As Long, _
	'ByVal nBitCount As Long, _
	'lpBits As Any) As Long
	Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	'Private Declare Function SetPixel _
	'Lib "gdi32.dll" (ByVal hdc As Long, _
	'ByVal x As Long, _
	'ByVal Y As Long, _
	'ByVal crColor As Long) As Long
	Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function CreateWindowEx Lib "user32"  Alias "CreateWindowExA"(ByVal dwExStyle As Integer, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As Integer, ByVal hMenu As Integer, ByVal hInstance As Integer, ByRef lpParam As Any) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	'Private Declare Function GetWindowRect _
	'Lib "user32" (ByVal hwnd As Long, _
	'lpRect As RECT) As Long
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function GetDC Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Integer, ByVal HPALETTE As Integer, ByRef pccolorref As Integer) As Integer
	'Private Declare Function GetCursorPos _
	'Lib "user32" (lpPoint As POINT) As Long
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Integer, ByRef qrc As RECT, ByVal edge As Integer, ByVal grfFlags As Integer) As Integer
	'UPGRADE_WARNING: Structure Point may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef lpPoint As Point) As Integer
	Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	'Private Declare Function DrawFocusRect _
	'Lib "user32" (ByVal hdc As Long, _
	'lpRect As RECT) As Long
	'Private Declare Function DrawState _
	'Lib "user32" _
	'Alias "DrawStateA" (ByVal hdc As Long, _
	'ByVal hBrush As Long, _
	'ByVal lpDrawStateProc As Long, _
	'ByVal lParam As Long, _
	'ByVal wParam As Long, _
	'ByVal x As Long, _
	'ByVal Y As Long, _
	'ByVal cX As Long, _
	'ByVal cY As Long, ByVal fuFlags As Long) As Long
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'Private Declare Function GetCurrentThemeName _
	'Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
	'ByVal dwMaxNameChars As Long, _
	'ByVal pszColorBuff As Long, _
	'ByVal cchMaxColorChars As Long, _
	'ByVal pszSizeBuff As Long, _
	'ByVal cchMaxSizeChars As Long) As Long
	'Private Declare Function RegisterWindowMessage _
	'Lib "user32" _
	'Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
	'Private Declare Function SystemParametersInfo _
	'Lib "user32" _
	'Alias "SystemParametersInfoA" (ByVal uAction As Long, _
	'ByVal uParam As Long, _
	'ByRef lpvParam As Long, _
	'ByVal fuWinIni As Long) As Long
	Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'Private Declare Function GetActiveWindow _
	'Lib "user32" () As Long
	'Private Declare Function WindowFromPoint _
	'Lib "user32" (ByVal xPoint As Long, _
	'ByVal yPoint As Long) As Long
	Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Integer, ByVal hRgn As Integer, ByVal bRedraw As Boolean) As Integer
	'Private Declare Function CreateEllipticRgn _
	'Lib "gdi32" (ByVal X1 As Long, _
	'ByVal Y1 As Long, _
	'ByVal X2 As Long, _
	'ByVal Y2 As Long) As Long
	Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
	'UPGRADE_WARNING: Structure Point may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreatePolygonRgn Lib "gdi32" (ByRef lpPoint As Point, ByVal nCount As Integer, ByVal nPolyFillMode As Integer) As Integer
	'Private Declare Function CreatePolyPolygonRgn _
	'Lib "gdi32" (lpPoint As POINT, _
	'lpPolyCounts As Long, _
	'ByVal nCount As Long, _
	'ByVal nPolyFillMode As Long) As Long
	Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	'Private Declare Function CombineRgn _
	'Lib "gdi32" (ByVal hDestRgn As Long, _
	'ByVal hSrcRgn1 As Long, _
	'ByVal hSrcRgn2 As Long, _
	'ByVal nCombineMode As Long) As Long
	'Private Declare Function GetVersionEx _
	'Lib "kernel32" _
	'Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
	Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
	'UPGRADE_WARNING: Structure ICONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateIconIndirect Lib "user32" (ByRef piconinfo As ICONINFO) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hdc As Integer) As Integer
	Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Integer, ByVal xLeft As Integer, ByVal yTop As Integer, ByVal hIcon As Integer, ByVal cxWidth As Integer, ByVal cyWidth As Integer, ByVal istepIfAniCur As Integer, ByVal hbrFlickerFreeDraw As Integer, ByVal diFlags As Integer) As Integer
	Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Integer) As Integer
	'UPGRADE_WARNING: Structure ICONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Integer, ByRef piconinfo As ICONINFO) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetObjectAPI Lib "gdi32"  Alias "GetObjectA"(ByVal hObject As Integer, ByVal nCount As Integer, ByRef lpObject As Any) As Integer
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Integer, ByVal hBitmap As Integer, ByVal nStartScan As Integer, ByVal nNumScans As Integer, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Integer) As Integer
	Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Integer, ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal Scan As Integer, ByVal NumScans As Integer, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Integer) As Integer
	
	'*************************************************************
	'
	'   Private variables
	'
	'*************************************************************
	Private m_bFocused As Boolean
	Private m_bVisible As Boolean
	Private m_iState As isState
	Private m_iStyle As isbStyle
	Private m_iNonThemeStyle As isbStyle
	Private m_btnRect As RECT
	Private m_txtRect As RECT
	Private m_lRegion As Integer
	Private m_sCaption As String
	Private m_CaptionAlign As isbAlign
	Private m_IconAlign As isbAlign
	Private m_Icon As System.Drawing.Image
	Private m_Font As System.Drawing.Font
	Private m_IconSize As Integer
	Private m_bEnabled As Boolean
	Private m_bShowFocus As Boolean
	Private m_bUseCustomColors As Boolean
	Private m_lBackColor As Integer
	Private m_lHighlightColor As Integer
	Private m_lFontColor As Integer
	Private m_lFontHighlightColor As Integer
	Private m_sToolTipText As String
	Private m_sTooltiptitle As String
	Private m_lToolTipIcon As ttIconType
	Private m_lToolTipType As ttStyleEnum
	Private m_lttBackColor As Integer
	Private m_lttForeColor As Integer
	Private m_lttCentered As Boolean
	Private m_lttHwnd As Integer
	Private m_ButtonType As isbButtonType
	Private m_Value As Boolean
	Private m_MaskColor As Integer
	Private m_UseMaskColor As Boolean
	Private m_bRoundedBordersByTheme As Boolean
	Private lPrevStyle As Integer
	Private iStyleIconOffset As Integer
	
	'for subclass
	Private sc_aSubData() As tSubData 'Subclass data array
	Private bTrack As Boolean
	Private bTrackUser32 As Boolean
	Private bInCtrl As Boolean
	
	'Auxiliar Variables
	Dim lwFontAlign As Integer
	Dim lPrevButton As Integer
	Dim ttip As TOOLINFO
	
	'*************************************************************
	'
	'   Public Events
	'
	'*************************************************************
	Public Shadows Event Click(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event MouseEnter(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event MouseLeave(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	
	Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Integer, ByRef lng_hWnd As Integer, ByRef uMsg As Integer, ByRef wParam As Integer, ByRef lParam As Integer)
		Select Case uMsg
			Case WM_MOUSEMOVE
				If Not bInCtrl Then
					bInCtrl = True
					Call TrackMouseLeave(lng_hWnd)
					m_iState = isState.stateHot
					Refresh()
					RaiseEvent MouseEnter(Me, Nothing)
					CreateToolTip()
				End If
			Case WM_MOUSELEAVE
				bInCtrl = False
				m_iState = isState.statenormal
				RemoveToolTip()
				Refresh()
				RaiseEvent MouseLeave(Me, Nothing)
			Case WM_SYSCOLORCHANGE
				Refresh()
			Case WM_THEMECHANGED
				Refresh()
		End Select
		Exit Sub
zSubclass_Proc_Error: 
	End Sub
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Subclass_AddMsg(ByVal lng_hWnd As Integer, ByVal uMsg As Integer, Optional ByVal When_Renamed As eMsgWhen = eMsgWhen.MSG_AFTER)
		On Error GoTo Subclass_AddMsg_Error
		With sc_aSubData(zIdx(lng_hWnd))
			If When_Renamed And eMsgWhen.MSG_BEFORE Then
				Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
			End If
			If When_Renamed And eMsgWhen.MSG_AFTER Then
				Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
			End If
		End With
		Exit Sub
Subclass_AddMsg_Error: 
	End Sub
	
	Private Function Subclass_InIDE() As Boolean
		On Error GoTo Subclass_InIDE_Error
		System.Diagnostics.Debug.Assert(zSetTrue(Subclass_InIDE), "")
		Exit Function
Subclass_InIDE_Error: 
	End Function
	
	Private Function Subclass_Start(ByVal lng_hWnd As Integer) As Integer
		On Error GoTo Subclass_Start_Error
		Const CODE_LEN As Integer = 200
		Const FUNC_CWP As String = "CallWindowProcA"
		Const FUNC_EBM As String = "EbMode"
		Const FUNC_SWL As String = "SetWindowLongA"
		Const MOD_USER As String = "user32"
		Const MOD_VBA5 As String = "vba5"
		Const MOD_VBA6 As String = "vba6"
		Const PATCH_01 As Integer = 18
		Const PATCH_02 As Integer = 68
		Const PATCH_03 As Integer = 78
		Const PATCH_06 As Integer = 116
		Const PATCH_07 As Integer = 121
		Const PATCH_0A As Integer = 186
		'UPGRADE_WARNING: Lower bound of array aBuf was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static aBuf(CODE_LEN) As Byte
		Static pCWP As Integer
		Static pEbMode As Integer
		Static pSWL As Integer
		Dim i As Integer
		Dim j As Integer
		Dim nSubIdx As Integer
		Dim sHex As String
		If aBuf(1) = 0 Then
			sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
			i = 1
			Do While j < CODE_LEN
				j = j + 1
				aBuf(j) = Val("&H" & Mid(sHex, i, 2))
				i = i + 2
			Loop 
			If Subclass_InIDE Then
				aBuf(16) = &H90s 'Patch the code buffer to enable the IDE state code
				aBuf(17) = &H90s 'Patch the code buffer to enable the IDE state code
				pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM) 'Get the address of EbMode in vba6.dll
				
				If pEbMode = 0 Then 'Found?
					pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM) 'VB5 perhaps
				End If
			End If
			
			pCWP = zAddrFunc(MOD_USER, FUNC_CWP) 'Get the address of the CallWindowsProc function
			pSWL = zAddrFunc(MOD_USER, FUNC_SWL) 'Get the address of the SetWindowLongA function
			ReDim sc_aSubData(0) 'Create the first sc_aSubData element
		Else
			nSubIdx = zIdx(lng_hWnd, True)
			
			If nSubIdx = -1 Then 'If an sc_aSubData element isn't being re-cycled
				nSubIdx = UBound(sc_aSubData) + 1 'Calculate the Next 'element
				ReDim Preserve sc_aSubData(nSubIdx) 'Create a new sc_aSubData element
			End If
			
			Subclass_Start = nSubIdx
		End If
		
		With sc_aSubData(nSubIdx)
			.hWnd = lng_hWnd 'Store the hWnd
			.nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN) 'Allocate memory for the machine code WndProc
			.nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub) 'Set our WndProc in place
			Call RtlMoveMemory(.nAddrSub, aBuf(1), CODE_LEN) 'Copy the machine code from the static byte array to the code array in sc_aSubData
			Call zPatchRel(.nAddrSub, PATCH_01, pEbMode) 'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
			Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig) 'Original WndProc address for CallWindowProc, call the original WndProc
			Call zPatchRel(.nAddrSub, PATCH_03, pSWL) 'Patch the relative address of the SetWindowLongA api function
			Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig) 'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
			Call zPatchRel(.nAddrSub, PATCH_07, pCWP) 'Patch the relative address of the CallWindowProc api function
			'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me)) 'Patch the address of this object instance into the static machine code buffer
		End With
		
		Exit Function
		
Subclass_Start_Error: 
	End Function
	
	Private Sub Subclass_StopAll()
		On Error Resume Next
		'On Error GoTo Subclass_StopAll_Error
		Dim i As Integer
		i = UBound(sc_aSubData)
		Do While i >= 0
			With sc_aSubData(i)
				If .hWnd <> 0 Then
					Call Subclass_Stop(.hWnd)
				End If
			End With
			i = i - 1
		Loop 
		Exit Sub
Subclass_StopAll_Error: 
	End Sub
	
	Private Sub Subclass_Stop(ByVal lng_hWnd As Integer)
		'On Error GoTo Subclass_Stop_Error
		With sc_aSubData(zIdx(lng_hWnd))
			Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
			Call zPatchVal(.nAddrSub, PATCH_05, 0)
			Call zPatchVal(.nAddrSub, PATCH_09, 0)
			Call GlobalFree(.nAddrSub)
			.hWnd = 0
			.nMsgCntB = 0
			.nMsgCntA = 0
			Erase .aMsgTblB
			Erase .aMsgTblA
		End With
		Exit Sub
Subclass_Stop_Error: 
	End Sub
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub zAddMsg(ByVal uMsg As Integer, ByRef aMsgTbl() As Integer, ByRef nMsgCnt As Integer, ByVal When_Renamed As eMsgWhen, ByVal nAddr As Integer)
		Dim nOff1, nEntry, nOff2 As Integer
		If uMsg = ALL_MESSAGES Then
			nMsgCnt = ALL_MESSAGES
		Else
			Do While nEntry < nMsgCnt
				nEntry = nEntry + 1
				If aMsgTbl(nEntry) = 0 Then
					aMsgTbl(nEntry) = uMsg
					Exit Sub
				ElseIf aMsgTbl(nEntry) = uMsg Then 
					Exit Sub
				End If
			Loop 
			nMsgCnt = nMsgCnt + 1
			'UPGRADE_WARNING: Lower bound of array aMsgTbl was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Preserve aMsgTbl(nMsgCnt)
			aMsgTbl(nMsgCnt) = uMsg
		End If
		
		If When_Renamed = eMsgWhen.MSG_BEFORE Then 'If before
			nOff1 = PATCH_04 'Offset to the Before table
			nOff2 = PATCH_05 'Offset to the Before table entry count
		Else 'Else after
			nOff1 = PATCH_08 'Offset to the After table
			nOff2 = PATCH_09 'Offset to the After table entry count
		End If
		
		If uMsg <> ALL_MESSAGES Then
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1))) 'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
		End If
		
		Call zPatchVal(nAddr, nOff2, nMsgCnt) 'Patch the appropriate table entry count
		Exit Sub
		
zAddMsg_Error: 
	End Sub
	
	'Return the memory address of the passed function in the passed dll
	Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Integer
		'On Error GoTo zAddrFunc_Error
		
		zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
		'  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
		Exit Function
		
zAddrFunc_Error: 
	End Function
	
	'Worker sub for Subclass_DelMsg
	'Private Sub zDelMsg(ByVal uMsg As Long, _
	''                    ByRef aMsgTbl() As Long, _
	''                    ByRef nMsgCnt As Long, _
	''                    ByVal When As eMsgWhen, _
	''                    ByVal nAddr As Long)
	'  'On Error GoTo zDelMsg_Error
	'
	'  Dim nEntry As Long
	'
	'  If uMsg = ALL_MESSAGES Then 'If deleting all messages
	'    nMsgCnt = 0 'Message count is now zero
	'
	'    If When = eMsgWhen.MSG_BEFORE Then 'If before
	'      nEntry = PATCH_05 'Patch the before table message count location
	'    Else 'Else after
	'      nEntry = PATCH_09 'Patch the after table message count location
	'    End If
	'
	'    Call zPatchVal(nAddr, nEntry, 0) 'Patch the table message count to zero
	'  Else 'Else deleteting a specific message
	'
	'    Do While nEntry < nMsgCnt 'For each table entry
	'      nEntry = nEntry + 1
	'
	'      If aMsgTbl(nEntry) = uMsg Then 'If this entry is the message we wish to delete
	'        aMsgTbl(nEntry) = 0 'Mark the table slot as available
	'        Exit Do 'Bail
	'      End If
	'
	'    Loop 'Next 'entry
	'
	'  End If
	'
	'  Exit Sub
	'
	'zDelMsg_Error:
	'End Sub
	
	'Get the sc_aSubData() array index of the passed hWnd
	Private Function zIdx(ByVal lng_hWnd As Integer, Optional ByVal bAdd As Boolean = False) As Integer
		'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
		'On Error GoTo zIdx_Error
		
		zIdx = UBound(sc_aSubData)
		
		Do While zIdx >= 0 'Iterate through the existing sc_aSubData() elements
			
			With sc_aSubData(zIdx)
				
				If .hWnd = lng_hWnd Then
					
					'If the hWnd of this element is the one we're looking for
					If Not bAdd Then 'If we're searching not adding
						Exit Function 'Found
						
					End If
					
				ElseIf .hWnd = 0 Then  'If this an element marked for reuse.
					
					If bAdd Then 'If we're adding
						Exit Function 'Re-use it
						
					End If
				End If
				
			End With
			
			zIdx = zIdx - 1 'Decrement the index
		Loop 
		
		'If we exit here, we're returning -1, no freed elements were found
		Exit Function
		
zIdx_Error: 
	End Function
	
	'Patch the machine code buffer at the indicated offset with the relative address to the target address.
	Private Sub zPatchRel(ByVal nAddr As Integer, ByVal nOffset As Integer, ByVal nTargetAddr As Integer)
		'On Error GoTo zPatchRel_Error
		
		Call RtlMoveMemory(nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
		Exit Sub
		
zPatchRel_Error: 
	End Sub
	
	'Patch the machine code buffer at the indicated offset with the passed value
	Private Sub zPatchVal(ByVal nAddr As Integer, ByVal nOffset As Integer, ByVal nValue As Integer)
		'On Error GoTo zPatchVal_Error
		
		Call RtlMoveMemory(nAddr + nOffset, nValue, 4)
		Exit Sub
		
zPatchVal_Error: 
	End Sub
	
	'Worker function for Subclass_InIDE
	Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
		'On Error GoTo zSetTrue_Error
		
		zSetTrue = True
		bValue = True
		Exit Function
		
zSetTrue_Error: 
	End Function
	
	'*************************************************************
	'
	'added by teee_eeee: unneded pMask Picture Box
	'
	'*************************************************************
	
	Private Sub TransBlt(ByVal DstDC As Integer, ByVal DstX As Integer, ByVal DstY As Integer, ByVal DstW As Integer, ByVal DstH As Integer, ByVal SrcPic As System.Drawing.Image, Optional ByVal TransColor As Integer = -1, Optional ByVal BrushColor As Integer = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)
		
		If DstW = 0 Or DstH = 0 Then Exit Sub
		
		Dim b As Integer
		Dim H As Integer
		Dim F As Integer
		Dim i As Integer
		Dim newW As Integer
		Dim TmpDC As Integer
		Dim TmpBmp As Integer
		Dim TmpObj As Integer
		Dim Sr2DC As Integer
		Dim Sr2Bmp As Integer
		Dim Sr2Obj As Integer
		Dim Data1() As RGBTRIPLE
		Dim Data2() As RGBTRIPLE
		Dim Info As BITMAPINFO
		Dim BrushRGB As RGBTRIPLE
		Dim gCol As Integer
		
		Dim SrcDC As Integer
		Dim tObj As Integer
		Dim ttt As Integer
		
		'UPGRADE_ISSUE: UserControl property ctlButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SrcDC = CreateCompatibleDC(hdc)
		
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstW < 0 Then DstW = MyBase.ScaleX(SrcPic.Width, 8, MyBase.ScaleMode)
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstH < 0 Then DstH = MyBase.ScaleY(SrcPic.Height, 8, MyBase.ScaleMode)
		
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Dim hBrush As Integer
		If SrcPic.Type = 1 Then
			tObj = SelectObject(SrcDC, CInt(CObj(SrcPic)))
		Else
			tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
			hBrush = CreateSolidBrush(System.Drawing.ColorTranslator.ToOle(MaskColor))
			'UPGRADE_ISSUE: Picture property SrcPic.handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawIconEx(SrcDC, 0, 0, SrcPic.Handle, 0, 0, 0, hBrush, &H1s Or &H2s)
			DeleteObject(hBrush)
		End If
		
		TmpDC = CreateCompatibleDC(SrcDC)
		Sr2DC = CreateCompatibleDC(SrcDC)
		TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		TmpObj = SelectObject(TmpDC, TmpBmp)
		Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
		
		ReDim Data1(DstW * DstH * 3 - 1)
		ReDim Data2(UBound(Data1))
		
		With Info.bmiHeader
			.biSize = Len(Info.bmiHeader)
			.biWidth = DstW
			.biHeight = DstH
			.biPlanes = 1
			.biBitCount = 24
		End With
		
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy)
		'UPGRADE_WARNING: Couldn't resolve default property of object Data1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object Data2(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0)
		
		If BrushColor > 0 Then
			BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100s
			BrushRGB.rgbGreen = (BrushColor \ &H100s) Mod &H100s
			BrushRGB.rgbRed = BrushColor And &HFFs
		End If
		
		If Not m_UseMaskColor Then TransColor = -1
		
		newW = DstW - 1
		
		For H = 0 To DstH - 1
			F = H * DstW
			For b = 0 To newW
				i = F + b
				'UPGRADE_ISSUE: UserControl property ctlButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If GetNearestColor(hdc, CInt(Data2(i).rgbRed) + 256 * Data2(i).rgbGreen + 65536 * Data2(i).rgbBlue) <> TransColor Then
					With Data1(i)
						If BrushColor > -1 Then
							If MonoMask Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (CInt(Data2(i).rgbRed) + Data2(i).rgbGreen + Data2(i).rgbBlue) <= 384 Then Data1(i) = BrushRGB
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Data1(i) = BrushRGB
							End If
						Else
							If isGreyscale Then
								gCol = CInt(Data2(i).rgbRed * 0.3) + Data2(i).rgbGreen * 0.59 + Data2(i).rgbBlue * 0.11
								.rgbRed = gCol : .rgbGreen = gCol : .rgbBlue = gCol
							Else
								If XPBlend Then
									.rgbRed = (CInt(.rgbRed) + Data2(i).rgbRed * 2) \ 3
									.rgbGreen = (CInt(.rgbGreen) + Data2(i).rgbGreen * 2) \ 3
									.rgbBlue = (CInt(.rgbBlue) + Data2(i).rgbBlue * 2) \ 3
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Data1(i) = Data2(i)
								End If
							End If
						End If
					End With
					
				End If
				
			Next b
			
		Next H
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Data1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0)
		
		Erase Data1
		Erase Data2
		DeleteObject(SelectObject(TmpDC, TmpObj))
		DeleteObject(SelectObject(Sr2DC, Sr2Obj))
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If SrcPic.Type = 3 Then DeleteObject(SelectObject(SrcDC, tObj))
		DeleteDC(TmpDC) : DeleteDC(Sr2DC)
		DeleteObject(tObj) : DeleteDC(SrcDC)
		
	End Sub
	
	'*************************************************************
	'
	'added by Dennis (dvrdsr) function excerpted vlad memdc class'
	'
	'*************************************************************
	Public Function PaintIconGrayscale(ByVal Dest_hDC As Integer, ByVal hIcon As Integer, Optional ByVal Dest_X As Integer = 0, Optional ByVal Dest_Y As Integer = 0, Optional ByVal Dest_Height As Integer = 0, Optional ByVal Dest_Width As Integer = 0) As Boolean
		'On Error GoTo PaintIconGrayscale_Error
		
		Dim hBMP_Mask As Integer
		Dim hBMP_Image As Integer
		Dim hBMP_Prev As Integer
		Dim hIcon_Temp As Integer
		Dim hDC_Temp As Integer
		
		' Make sure parameters passed are valid
		If Dest_hDC = 0 Or hIcon = 0 Then Exit Function
		
		' Extract the bitmaps from the icon
		If pvGetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function
		' Create a memory DC to work with
		hDC_Temp = CreateCompatibleDC(0)
		
		If hDC_Temp = 0 Then GoTo CleanUp
		
		' Make the image bitmap gradient
		If pvRenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0) = False Then GoTo CleanUp
		' Extract the gradient bitmap out of the DC
		SelectObject(hDC_Temp, hBMP_Prev)
		' Take the newly gradient bitmap and make a gradient icon from it
		hIcon_Temp = pvCreateIconFromBMP(hBMP_Mask, hBMP_Image)
		
		If hIcon_Temp = 0 Then GoTo CleanUp
		
		' Draw the newly created gradient icon onto the specified DC
		If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3s) <> 0 Then
			PaintIconGrayscale = True
		End If
		
CleanUp: 
		DestroyIcon(hIcon_Temp) : hIcon_Temp = 0
		DeleteDC(hDC_Temp) : hDC_Temp = 0
		DeleteObject(hBMP_Mask) : hBMP_Mask = 0
		DeleteObject(hBMP_Image) : hBMP_Image = 0
		Exit Function
		
PaintIconGrayscale_Error: 
	End Function
	
	Private Function pvGetIconBitmaps(ByVal hIcon As Integer, ByRef Return_hBmpMask As Integer, ByRef Return_hBmpImage As Integer) As Boolean
		'On Error GoTo pvGetIconBitmaps_Error
		
		Dim TempICONINFO As ICONINFO
		
		If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
		Return_hBmpMask = TempICONINFO.hbmMask
		Return_hBmpImage = TempICONINFO.hbmColor
		pvGetIconBitmaps = True
		Exit Function
		
pvGetIconBitmaps_Error: 
	End Function
	
	Private Function pvRenderBitmapGrayscale(ByVal Dest_hDC As Integer, ByVal hBitmap As Integer, Optional ByVal Dest_X As Integer = 0, Optional ByVal Dest_Y As Integer = 0, Optional ByVal Srce_X As Integer = 0, Optional ByVal Srce_Y As Integer = 0) As Boolean
		'On Error GoTo pvRenderBitmapGrayscale_Error
		
		Dim TempBITMAP As BITMAP
		Dim hScreen As Integer
		Dim hDC_Temp As Integer
		Dim hBMP_Prev As Integer
		Dim MyCounterX As Integer
		Dim MyCounterY As Integer
		Dim NewColor As Integer
		Dim hNewPicture As Integer
		Dim DeletePic As Boolean
		
		' Make sure parameters passed are valid
		If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
		' Get the handle to the screen DC
		hScreen = GetDC(0)
		
		If hScreen = 0 Then Exit Function
		' Create a memory DC to work with the picture
		hDC_Temp = CreateCompatibleDC(hScreen)
		
		If hDC_Temp = 0 Then GoTo CleanUp
		' If the user specifies NOT to alter the original, then make a copy of it to use
		DeletePic = False
		hNewPicture = hBitmap
		' Select the bitmap into the DC
		hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
		
		' Get the height / width of the bitmap in pixels
		'UPGRADE_WARNING: Couldn't resolve default property of object TempBITMAP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
		If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp
		
		' Loop through each pixel and conver it to it's grayscale equivelant
		For MyCounterX = 0 To TempBITMAP.bmWidth - 1
			For MyCounterY = 0 To TempBITMAP.bmHeight - 1
				NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
				
				If NewColor <> -1 Then
					
					Select Case NewColor
						' If the color is already a grey shade, no need to convert it
						Case System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White), &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
							NewColor = NewColor
						Case Else
							NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + 0.11 * ((NewColor \ 65536) Mod 256)
							NewColor = RGB(NewColor, NewColor, NewColor)
					End Select
					
					SetPixelV(hDC_Temp, MyCounterX, MyCounterY, NewColor)
				End If
				
			Next  'MyCounterY
		Next  'MyCounterX
		
		' Display the picture on the specified hDC
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy)
		pvRenderBitmapGrayscale = True
CleanUp: 
		ReleaseDC(0, hScreen) : hScreen = 0
		SelectObject(hDC_Temp, hBMP_Prev)
		DeleteDC(hDC_Temp) : hDC_Temp = 0
		
		If DeletePic = True Then
			DeleteObject(hNewPicture)
			hNewPicture = 0
		End If
		
		Exit Function
		
pvRenderBitmapGrayscale_Error: 
	End Function
	
	Private Function pvCreateIconFromBMP(ByVal hBMP_Mask As Integer, ByVal hBMP_Image As Integer) As Integer
		'On Error GoTo pvCreateIconFromBMP_Error
		
		Dim TempICONINFO As ICONINFO
		
		If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
		TempICONINFO.fIcon = 1
		TempICONINFO.hbmMask = hBMP_Mask
		TempICONINFO.hbmColor = hBMP_Image
		pvCreateIconFromBMP = CreateIconIndirect(TempICONINFO)
		Exit Function
		
pvCreateIconFromBMP_Error: 
	End Function
	
	'*************************************************************
	'
	'   Private Auxiliar Subs
	'
	'*************************************************************
	'draw a Line Using API call's
	Private Sub APILine(ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer, ByRef lColor As Integer)
		'Use the API LineTo for Fast Drawing
		'On Error GoTo APILine_Error
		
		Dim pt As Point
		Dim hPen, hPenOld As Integer
		hPen = CreatePen(0, 1, lColor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		hPenOld = SelectObject(MyBase.hdc, hPen)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		MoveToEx(MyBase.hdc, X1, Y1, pt)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		LineTo(MyBase.hdc, X2, Y2)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SelectObject(MyBase.hdc, hPenOld)
		DeleteObject(hPen)
		Exit Sub
		
APILine_Error: 
	End Sub
	
	' full version of APILine
	Private Sub APILineEx(ByRef lhdcEx As Integer, ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer, ByRef lColor As Integer)
		'Use the API LineTo for Fast Drawing
		'On Error GoTo APILineEx_Error
		
		Dim pt As Point
		Dim hPen, hPenOld As Integer
		hPen = CreatePen(0, 1, lColor)
		hPenOld = SelectObject(lhdcEx, hPen)
		MoveToEx(lhdcEx, X1, Y1, pt)
		LineTo(lhdcEx, X2, Y2)
		SelectObject(lhdcEx, hPenOld)
		DeleteObject(hPen)
		Exit Sub
		
APILineEx_Error: 
	End Sub
	
	Private Sub APIFillRect(ByRef hdc As Integer, ByRef rc As RECT, ByRef Color As Integer)
		'On Error GoTo APIFillRect_Error
		
		Dim NewBrush As Integer
		NewBrush = CreateSolidBrush(Color)
		Call FillRect(hdc, rc, NewBrush)
		Call DeleteObject(NewBrush)
		Exit Sub
		
APIFillRect_Error: 
	End Sub
	
	Private Sub APIFillRectByCoords(ByRef hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal H As Integer, ByRef Color As Integer)
		'On Error GoTo APIFillRectByCoords_Error
		
		Dim NewBrush As Integer
		Dim tmpRect As RECT
		NewBrush = CreateSolidBrush(Color)
		SetRect(tmpRect, X, Y, X + w, Y + H)
		Call FillRect(hdc, tmpRect, NewBrush)
		Call DeleteObject(NewBrush)
		Exit Sub
		
APIFillRectByCoords_Error: 
	End Sub
	
	Private Function APIRectangle(ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal H As Integer, Optional ByRef lColor As System.Drawing.Color = -1) As Integer
		'On Error GoTo APIRectangle_Error
		
		Dim hPen, hPenOld As Integer
		Dim pt As Point
		hPen = CreatePen(0, 1, System.Drawing.ColorTranslator.ToOle(lColor))
		hPenOld = SelectObject(hdc, hPen)
		MoveToEx(hdc, X, Y, pt)
		LineTo(hdc, X + w, Y)
		LineTo(hdc, X + w, Y + H)
		LineTo(hdc, X, Y + H)
		LineTo(hdc, X, Y)
		SelectObject(hdc, hPenOld)
		DeleteObject(hPen)
		Exit Function
		
APIRectangle_Error: 
	End Function
	
	'Private Sub DrawCtlEdgeByRect(hdc As Long, _
	''                              rt As RECT, _
	''                              Optional Style As Long = EDGE_RAISED, _
	''                              Optional Flags As Long = BF_RECT)
	'  'On Error GoTo DrawCtlEdgeByRect_Error
	'
	'  DrawEdge hdc, rt, Style, Flags
	'  Exit Sub
	'
	'DrawCtlEdgeByRect_Error:
	'End Sub
	
	Private Sub DrawCtlEdge(ByRef hdc As Integer, ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single, Optional ByRef Style As Integer = EDGE_RAISED, Optional ByVal Flags As Integer = BF_RECT)
		'On Error GoTo DrawCtlEdge_Error
		
		Dim R As RECT
		
		With R
			.Left_Renamed = X
			.Top_Renamed = Y
			.Right_Renamed = X + w
			.Bottom_Renamed = Y + H
		End With
		
		DrawEdge(hdc, R, Style, Flags)
		Exit Sub
		
DrawCtlEdge_Error: 
	End Sub
	
	'Blend two colors
	Private Function BlendColors(ByVal lcolor1 As Integer, ByVal lcolor2 As Integer) As Object
		'On Error GoTo BlendColors_Error
		
		BlendColors = RGB(CShort(CShort(lcolor1 And &HFFs) + CShort(lcolor2 And &HFFs)) / 2, CShort(CShort((lcolor1 \ &H100s) And &HFFs) + CShort((lcolor2 \ &H100s) And &HFFs)) / 2, CShort(CShort((lcolor1 \ &H10000) And &HFFs) + CShort((lcolor2 \ &H10000) And &HFFs)) / 2)
		Exit Function
		
BlendColors_Error: 
	End Function
	
	'System color code to long rgb
	Private Function TranslateColor(ByVal lColor As Integer) As Integer
		'On Error GoTo TranslateColor_Error
		
		If OleTranslateColor(lColor, 0, TranslateColor) Then
			TranslateColor = -1
		End If
		
		Exit Function
		
TranslateColor_Error: 
	End Function
	
	'Make Soft a color
	'Private Function SoftColor(lcolor As OLE_COLOR) As OLE_COLOR
	'  'On Error GoTo SoftColor_Error
	'
	'  Dim lRed As OLE_COLOR
	'  Dim lGreen As OLE_COLOR
	'  Dim lBlue As OLE_COLOR
	'  Dim lR As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
	'  lR = (lcolor And &HFF)
	'  lg = ((lcolor And 65280) \ 256)
	'  lb = ((lcolor) And 16711680) \ 65536
	'  lRed = (76 - Int(((lcolor And &HFF) + 32) \ 64) * 19)
	'  lGreen = (76 - Int((((lcolor And 65280) \ 256) + 32) \ 64) * 19)
	'  lBlue = (76 - Int((((lcolor And &HFF0000) \ &H10000) + 32) / 64) * 19)
	'  SoftColor = RGB(lR + lRed, lg + lGreen, lb + lBlue)
	'  Exit Function
	'
	'SoftColor_Error:
	'End Function
	
	Private Function MSOXPShiftColor(ByVal theColor As Integer, Optional ByVal Base As Integer = &HB0s) As Integer
		'On Error GoTo MSOXPShiftColor_Error
		
		Dim Blue, Red, Green As Integer
		Dim Delta As Integer
		Blue = ((theColor \ &H10000) Mod &H100s)
		Green = ((theColor \ &H100s) Mod &H100s)
		Red = (theColor And &HFFs)
		Delta = &HFFs - Base
		Blue = Base + Blue * Delta \ &HFFs
		Green = Base + Green * Delta \ &HFFs
		Red = Base + Red * Delta \ &HFFs
		
		If Red > 255 Then Red = 255
		If Green > 255 Then Green = 255
		If Blue > 255 Then Blue = 255
		MSOXPShiftColor = Red + 256 * Green + 65536 * Blue
		Exit Function
		
MSOXPShiftColor_Error: 
	End Function
	
	'Private Function msSoftColor(lcolor As Long) As Long
	'  'On Error GoTo msSoftColor_Error
	'
	'  Dim lRed As Long
	'  Dim lGreen As Long
	'  Dim lBlue As Long
	'  Dim lR As Long, lg As Long, lb As Long
	'  lR = (lcolor And &HFF)
	'  lg = ((lcolor And 65280) \ 256)
	'  lb = ((lcolor) And 16711680) \ 65536
	'  lRed = (76 - Int(((lcolor And &HFF) + 32) \ 64) * 19)
	'  lGreen = (76 - Int((((lcolor And 65280) \ 256) + 32) \ 64) * 19)
	'  lBlue = (76 - Int((((lcolor And &HFF0000) \ &H10000) + 32) / 64) * 19)
	'  msSoftColor = RGB(lR + lRed, lg + lGreen, lb + lBlue)
	'  Exit Function
	'
	'msSoftColor_Error:
	'End Function
	
	'Offset a color
	Private Function OffsetColor(ByRef lColor As System.Drawing.Color, ByRef lOffset As Integer) As System.Drawing.Color
		'On Error GoTo OffsetColor_Error
		
		Dim lRed As System.Drawing.Color
		Dim lGreen As System.Drawing.Color
		Dim lBlue As System.Drawing.Color
		Dim lg, lR, lb As System.Drawing.Color
		lR = System.Drawing.ColorTranslator.FromOle((lColor And &HFFs))
		lg = System.Drawing.ColorTranslator.FromOle(((lColor And 65280) \ 256))
		lb = System.Drawing.ColorTranslator.FromOle(((lColor) And 16711680) \ 65536)
		lRed = System.Drawing.ColorTranslator.FromOle((lOffset + System.Drawing.ColorTranslator.ToOle(lR)))
		lGreen = System.Drawing.ColorTranslator.FromOle((lOffset + System.Drawing.ColorTranslator.ToOle(lg)))
		lBlue = System.Drawing.ColorTranslator.FromOle((lOffset + System.Drawing.ColorTranslator.ToOle(lb)))
		
		If System.Drawing.ColorTranslator.ToOle(lRed) > 255 Then lRed = System.Drawing.ColorTranslator.FromOle(255)
		If System.Drawing.ColorTranslator.ToOle(lRed) < 0 Then lRed = System.Drawing.ColorTranslator.FromOle(0)
		If System.Drawing.ColorTranslator.ToOle(lGreen) > 255 Then lGreen = System.Drawing.ColorTranslator.FromOle(255)
		If System.Drawing.ColorTranslator.ToOle(lGreen) < 0 Then lGreen = System.Drawing.ColorTranslator.FromOle(0)
		If System.Drawing.ColorTranslator.ToOle(lBlue) > 255 Then lBlue = System.Drawing.ColorTranslator.FromOle(255)
		If System.Drawing.ColorTranslator.ToOle(lBlue) < 0 Then lBlue = System.Drawing.ColorTranslator.FromOle(0)
		OffsetColor = System.Drawing.ColorTranslator.FromOle(RGB(System.Drawing.ColorTranslator.ToOle(lRed), System.Drawing.ColorTranslator.ToOle(lGreen), System.Drawing.ColorTranslator.ToOle(lBlue)))
		Exit Function
		
OffsetColor_Error: 
	End Function
	
	Private Sub DrawCaption()
		'On Error GoTo DrawCaption_Error
		
		Dim lColor, ltmpColor As Integer
		
		If Not m_bUseCustomColors Then
			If m_iState <> isState.statedisabled Then
				lColor = GetSysColor(COLOR_BTNTEXT)
			Else
				lColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.GrayText))
			End If
			
		Else
			
			Select Case m_iState
				Case isState.statenormal
					lColor = m_lFontColor
				Case isState.statedisabled
					lColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.GrayText))
				Case Else
					lColor = m_lFontHighlightColor
			End Select
			
		End If
		
		ltmpColor = System.Drawing.ColorTranslator.ToOle(MyBase.ForeColor)
		MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(lColor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		DrawText(MyBase.hdc, m_sCaption, -1, m_txtRect, lwFontAlign)
		MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(ltmpColor)
		Exit Sub
		
DrawCaption_Error: 
	End Sub
	
	'Private Sub fPaintPicture(ByRef m_Picture As StdPicture, _
	''                          ByVal x As Long, _
	''                          ByVal Y As Long, _
	''                          ByVal w As Long, _
	''                          ByVal h As Long)
	'  'On Error GoTo fPaintPicture_Error
	'
	'  Dim memDC As Long, memDC1 As Long
	'  Dim membitmap As Long
	'  Dim oldW As Long, oldH As Long
	'  'setup w,h vars
	'  oldW = m_Picture.Width: oldH = m_Picture.Height
	'  'create compatible DC
	'  memDC = CreateCompatibleDC(UserControl.hdc)
	'  'create the copy on the
	'  membitmap = SelectObject(memDC, m_Picture.Handle)
	'  'BitBlt memDC, 0, 0, oldW, oldH, vbSrcCopy
	'  StretchBlt UserControl.hdc, x, Y, w, h, memDC, 0, 0, oldW, oldH, vbSrcCopy
	'  BitBlt UserControl.hdc, x, Y, w, h, memDC, 0, 0, vbSrcCopy
	'  Exit Sub
	'
	'fPaintPicture_Error:
	'End Sub
	
	''''''''
	' Under test
	'http://www.visual-basic.com.ar/vbsmart/library/smartnetbutton/smartnetbutton.htm
	'Private Sub fDrawPicture(ByRef m_Picture As StdPicture, _
	''                         ByVal x As Long, _
	''                         ByVal Y As Long, _
	''                         ByVal bShadow As Boolean)
	'  'On Error GoTo fDrawPicture_Error
	'
	'  Dim lFlags As Long
	'  Dim hBrush As Long
	'
	'  Select Case m_Picture.Type
	'    Case vbPicTypeBitmap
	'      lFlags = DST_BITMAP
	'    Case vbPicTypeIcon
	'      lFlags = DST_ICON
	'    Case Else
	'      lFlags = DST_COMPLEX
	'  End Select
	'
	'  If bShadow Then
	'    hBrush = CreateSolidBrush(RGB(128, 128, 128))
	'  End If
	'
	'  DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, x, Y, UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels), UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels), lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
	'
	'  If bShadow Then
	'    DeleteObject hBrush
	'  End If
	'
	'  Exit Sub
	'
	'fDrawPicture_Error:
	'End Sub
	
	Private Sub DrawVGradient(ByRef lEndColor As Integer, ByRef lStartcolor As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
		''Draw a Vertical Gradient in the current HDC
		'On Error GoTo DrawVGradient_Error
		
		Dim dG, dR, dB As Single
		Dim sG, sR, sB As Single
		Dim eG, eR, eB As Single
		Dim ni As Integer
		'lh = UserControl.ScaleHeight
		'lw = UserControl.ScaleWidth
		sR = (lStartcolor And &HFFs)
		sG = (lStartcolor \ &H100s) And &HFFs
		sB = CShort(lStartcolor And &HFF0000) / &H10000
		eR = (lEndColor And &HFFs)
		eG = (lEndColor \ &H100s) And &HFFs
		eB = CShort(lEndColor And &HFF0000) / &H10000
		dR = (sR - eR) / Y2
		dG = (sG - eG) / Y2
		dB = (sB - eB) / Y2
		
		For ni = 0 To Y2
			APILine(X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
		Next  'ni
		
		Exit Sub
		
DrawVGradient_Error: 
	End Sub
	
	Private Sub DrawVGradientEx(ByRef lhdcEx As Integer, ByRef lEndColor As Integer, ByRef lStartcolor As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
		''Draw a Vertical Gradient in the current HDC
		'On Error GoTo DrawVGradientEx_Error
		
		Dim dG, dR, dB As Single
		Dim sG, sR, sB As Single
		Dim eG, eR, eB As Single
		Dim ni As Integer
		'lh = UserControl.ScaleHeight
		'lw = UserControl.ScaleWidth
		sR = (lStartcolor And &HFFs)
		sG = (lStartcolor \ &H100s) And &HFFs
		sB = CShort(lStartcolor And &HFF0000) / &H10000
		eR = (lEndColor And &HFFs)
		eG = (lEndColor \ &H100s) And &HFFs
		eB = CShort(lEndColor And &HFF0000) / &H10000
		dR = (sR - eR) / Y2
		dG = (sG - eG) / Y2
		dB = (sB - eB) / Y2
		
		For ni = 0 To Y2
			APILineEx(lhdcEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
		Next  'ni
		
		Exit Sub
		
DrawVGradientEx_Error: 
	End Sub
	
	Private Sub DrawHGradient(ByRef lEndColor As Integer, ByRef lStartcolor As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
		''Draw a Horizontal Gradient in the current HDC
		'On Error GoTo DrawHGradient_Error
		
		Dim dG, dR, dB As Single
		Dim sG, sR, sB As Single
		Dim eG, eR, eB As Single
		Dim lh, lw As Integer
		Dim ni As Integer
		lh = Y2 - Y
		lw = X2 - X
		sR = (lStartcolor And &HFFs)
		sG = (lStartcolor \ &H100s) And &HFFs
		sB = CShort(lStartcolor And &HFF0000) / &H10000
		eR = (lEndColor And &HFFs)
		eG = (lEndColor \ &H100s) And &HFFs
		eB = CShort(lEndColor And &HFF0000) / &H10000
		dR = (sR - eR) / lw
		dG = (sG - eG) / lw
		dB = (sB - eB) / lw
		
		For ni = 0 To lw
			APILine(X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
		Next  'ni
		
		Exit Sub
		
DrawHGradient_Error: 
	End Sub
	
	Private Sub DrawJavaBorder(ByVal X As Integer, ByVal Y As Integer, ByVal w As Integer, ByVal H As Integer, ByVal lColorShadow As Integer, ByVal lColorLight As Integer, ByVal lColorBack As Integer)
		'On Error GoTo DrawJavaBorder_Error
		
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		APIRectangle(MyBase.hdc, X, Y, w - 1, H - 1, System.Drawing.ColorTranslator.FromOle(lColorShadow))
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		APIRectangle(MyBase.hdc, X + 1, Y + 1, w - 1, H - 1, System.Drawing.ColorTranslator.FromOle(lColorLight))
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, X, Y + H, lColorBack)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, X + w, Y, lColorBack)
		'UPGRADE_WARNING: Couldn't resolve default property of object BlendColors(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, X + 1, Y + H - 1, BlendColors(lColorLight, lColorShadow))
		'UPGRADE_WARNING: Couldn't resolve default property of object BlendColors(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, X + w - 1, Y + 1, BlendColors(lColorLight, lColorShadow))
		Exit Sub
		
DrawJavaBorder_Error: 
	End Sub
	
	Private Function DrawTheme(ByRef sClass As String, ByVal iPart As Integer, ByVal iState As Integer) As Boolean
		Dim hTheme As Integer
		Dim lResult As Integer
		Dim m_btnRect2 As RECT
		Dim hRgn As Integer
		'On Error GoTo NoXP
		
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		hTheme = OpenThemeData(MyBase.Handle.ToInt32, StrPtr(sClass))
		
		If hTheme Then
			If m_bRoundedBordersByTheme Then
				'<--Rounded Region as requested for some themes:
				'Thanks to Dana Seaman-->
				SetRect(m_btnRect2, m_btnRect.Left_Renamed - 1, m_btnRect.Top_Renamed - 1, m_btnRect.Right_Renamed + 1, m_btnRect.Bottom_Renamed + 1)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				lResult = GetThemeBackgroundRegion(hTheme, MyBase.hdc, iPart, iState, m_btnRect2, hRgn)
				SetWindowRgn(Handle.ToInt32, hRgn, True)
				'free the memory.
				DeleteObject(hRgn)
			End If
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			lResult = DrawThemeBackground(hTheme, MyBase.hdc, iPart, iState, m_btnRect, m_btnRect)
			DrawTheme = IIf(lResult, False, True)
		Else
			DrawTheme = False
		End If
		
		Exit Function
		
NoXP: 
		DrawTheme = False
	End Function
	
	Private Function CreateWinXPregion() As Integer
		'On Error GoTo CreateWinXPRegion_Error
		
		Dim pPoligon(8) As Point
		Dim cpPoligon(1) As Integer
		Dim lw, lh As Integer
		lw = MyBase.ClientRectangle.Width
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		cpPoligon(0) = 5
		cpPoligon(1) = 5
		pPoligon(0).X = 0 : pPoligon(0).Y = 1
		pPoligon(1).X = 1 : pPoligon(1).Y = 0
		pPoligon(2).X = lw - 1 : pPoligon(2).Y = 0
		pPoligon(3).X = lw : pPoligon(3).Y = 1
		pPoligon(4).X = lw : pPoligon(4).Y = lh - 2
		pPoligon(5).X = lw - 2 : pPoligon(5).Y = lh
		pPoligon(6).X = 2 : pPoligon(6).Y = lh
		pPoligon(7).X = 0 : pPoligon(7).Y = lh - 2
		'pPoligon(8).x = 0: pPoligon(8).y = lh - 2
		CreateWinXPregion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)
		Exit Function
		
CreateWinXPRegion_Error: 
	End Function
	
	Private Function CreateGalaxyRegion() As Integer
		'On Error GoTo CreateGalaxyRegion_Error
		
		Dim pPoligon(8) As Point
		Dim cpPoligon(1) As Integer
		Dim lw, lh As Integer
		lw = MyBase.ClientRectangle.Width
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		cpPoligon(0) = 5
		cpPoligon(1) = 5
		pPoligon(0).X = 0 : pPoligon(0).Y = 2
		pPoligon(1).X = 2 : pPoligon(1).Y = 0
		pPoligon(2).X = lw - 3 : pPoligon(2).Y = 0
		pPoligon(3).X = lw : pPoligon(3).Y = 3
		pPoligon(4).X = lw : pPoligon(4).Y = lh - 3
		pPoligon(5).X = lw - 3 : pPoligon(5).Y = lh
		pPoligon(6).X = 4 : pPoligon(6).Y = lh
		pPoligon(7).X = 0 : pPoligon(7).Y = lh - 4
		'pPoligon(8).x = 0: pPoligon(8).y = lh - 2
		CreateGalaxyRegion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)
		Exit Function
		
CreateGalaxyRegion_Error: 
	End Function
	
	Private Function CreateMacOSXButtonRegion() As Integer
		'MsgBox "MACOS?"
		'On Error GoTo CreateMacOSXButtonRegion_Error
		
		CreateMacOSXButtonRegion = CreateRoundRectRgn(0, 0, MyBase.ClientRectangle.Width + 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) + 1, 18, 18)
		Exit Function
		
CreateMacOSXButtonRegion_Error: 
	End Function
	
	Public Sub About()
		'On Error GoTo About_Error
		
		m_About.Visible = True
		SetWindowLong(m_About.Handle.ToInt32, GWL_STYLE, lPrevStyle + WS_CAPTION + WS_THICKFRAME + WS_MINIMIZEBOX)
		SetWindowPos(m_About.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW) 'Or SWP_NOACTIVATE
		SetWindowPos(m_About.Handle.ToInt32, 0, 0, 0, 0, 0, SWP_REFRESH)
		SetWindowText(m_About.Handle.ToInt32, "About isButton " & strCurrentVersion)
		SetWindowPos(m_About.Handle.ToInt32, 0, 0, 0, 0, 0, SWP_REFRESH)
		SetParent(m_About.Handle.ToInt32, 0)
		''This is the ocx version about dialog
		'frmAbout.Show vbModal
		Exit Sub
		
About_Error: 
	End Sub
	
	'****************************************************************
	'
	'   Procedures
	'
	'****************************************************************
	Private Sub DrawWinXPButton(ByRef Mode As isState)
		'' This Sub Draws the XPStyle Button
		'On Error GoTo DrawWinXPButton_Error
		
		Dim lhdc As Integer
		Dim tempColor As Integer
		Dim lh, lw As Integer
		Dim lcw, lch As Integer
		Dim lStep As Single
		lw = MyBase.ClientRectangle.Width
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lhdc = MyBase.hdc
		lcw = m_btnRect.Left_Renamed + lw / 2 + 1
		lch = m_btnRect.Top_Renamed + lh / 2
		lStep = 25 / lh
		MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE))
		
		Select Case Mode
			Case isState.statenormal, isState.stateHot
				'Main
				DrawVGradient(&HFBFCFC, &HF0F0F0, 1, 1, lw - 2, 4)
				DrawVGradient(&HF9FAFA, &HEAF0F0, 1, 4, lw - 2, lh - 8)
				DrawVGradient(&HE6EBEB, &HC5D0D6, 1, lh - 4, lw - 2, 3)
				'right
				DrawVGradient(&HFAFBFB, &HDAE2E4, lw - 3, 3, lw - 2, lh - 5)
				DrawVGradient(&HF2F4F5, &HCDD7DB, lw - 2, 3, lw - 1, lh - 5)
				'Border
				APILine(1, 0, lw - 1, 0, &H743C00)
				APILine(0, 1, 0, lh - 1, &H743C00)
				APILine(lw - 1, 1, lw - 1, lh - 1, &H743C00)
				APILine(1, lh - 1, lw - 1, lh - 1, &H743C00)
				'Corners
				SetPixelV(lhdc, 1, 1, &H906E48)
				SetPixelV(lhdc, 1, lh - 2, &H906E48)
				SetPixelV(lhdc, lw - 2, 1, &H906E48)
				SetPixelV(lhdc, lw - 2, lh - 2, &H906E48)
				'External Borders
				SetPixelV(lhdc, 0, 1, &HA28B6A)
				SetPixelV(lhdc, 1, 0, &HA28B6A)
				SetPixelV(lhdc, 1, lh - 1, &HA28B6A)
				SetPixelV(lhdc, 0, lh - 2, &HA28B6A)
				SetPixelV(lhdc, lw - 1, lh - 2, &HA28B6A)
				SetPixelV(lhdc, lw - 2, lh - 1, &HA28B6A)
				SetPixelV(lhdc, lw - 2, 0, &HA28B6A)
				SetPixelV(lhdc, lw - 1, 1, &HA28B6A)
				'Internal Soft
				SetPixelV(lhdc, 1, 2, &HCAC7BF)
				SetPixelV(lhdc, 2, 1, &HCAC7BF)
				SetPixelV(lhdc, 2, lh - 2, &HCAC7BF)
				SetPixelV(lhdc, 1, lh - 3, &HCAC7BF)
				SetPixelV(lhdc, lw - 2, lh - 3, &HCAC7BF)
				SetPixelV(lhdc, lw - 3, lh - 2, &HCAC7BF)
				SetPixelV(lhdc, lw - 3, 1, &HCAC7BF)
				SetPixelV(lhdc, lw - 2, 2, &HCAC7BF)
				
				If Mode = isState.stateHot Then
					APILine(2, 1, lw - 2, 1, &HCFF0FF)
					APILine(2, 2, lw - 2, 2, &H89D8FD)
					APILine(2, lh - 3, lw - 2, lh - 3, &H30B3F8)
					APILine(2, lh - 2, lw - 2, lh - 2, &H1097E5)
					DrawVGradient(&H89D8FD, &H30B3F8, 1, 2, 3, lh - 5)
					DrawVGradient(&H89D8FD, &H30B3F8, lw - 3, 2, lw - 1, lh - 5)
					'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayAsDefault was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				ElseIf (Mode = isState.statenormal And m_bFocused) Or Ambient.DisplayAsDefault Then 
					APILine(2, lh - 2, lw - 2, lh - 2, &HEE8269)
					APILine(2, 1, lw - 2, 1, &HFFE7CE)
					APILine(2, 2, lw - 2, 2, &HF6D4BC)
					APILine(2, lh - 3, lw - 2, lh - 3, &HE4AD89)
					DrawVGradient(&HF6D4BC, &HE4AD89, 1, 2, 3, lh - 5)
					DrawVGradient(&HF6D4BC, &HE4AD89, lw - 3, 2, lw - 1, lh - 5)
				End If
				
			Case isState.statePressed
				' &HC1ccD1 - &HDBE2E3   -&HDCE3E4   -&HC1CCD1   -&HEEF1F2
				'Main
				DrawVGradient(&HC1CCD1, &HDCE3E4, 2, 1, lw - 1, 4)
				DrawVGradient(&HDCE3E4, &HDBE2E3, 2, 4, lw - 1, lh - 8)
				DrawVGradient(&HDBE2E3, &HEEF1F2, 3, lh - 4, lw - 1, 3)
				'left
				DrawVGradient(&HCED8DA, &HDBE2E3, 1, 3, 2, lh - 5)
				DrawVGradient(&HCED8DA, &HDBE2E3, 2, 4, 3, lh - 7)
				'Border
				APILine(1, 0, lw - 1, 0, &H743C00)
				APILine(0, 1, 0, lh - 1, &H743C00)
				APILine(lw - 1, 1, lw - 1, lh - 1, &H743C00)
				APILine(1, lh - 1, lw - 1, lh - 1, &H743C00)
				'Corners
				SetPixelV(lhdc, 1, 1, &H906E48)
				SetPixelV(lhdc, 1, lh - 2, &H906E48)
				SetPixelV(lhdc, lw - 2, 1, &H906E48)
				SetPixelV(lhdc, lw - 2, lh - 2, &H906E48)
				'External Borders
				SetPixelV(lhdc, 0, 1, &HA28B6A)
				SetPixelV(lhdc, 1, 0, &HA28B6A)
				SetPixelV(lhdc, 1, lh - 1, &HA28B6A)
				SetPixelV(lhdc, 0, lh - 2, &HA28B6A)
				SetPixelV(lhdc, lw - 1, lh - 2, &HA28B6A)
				SetPixelV(lhdc, lw - 2, lh - 1, &HA28B6A)
				SetPixelV(lhdc, lw - 2, 0, &HA28B6A)
				SetPixelV(lhdc, lw - 1, 1, &HA28B6A)
				'Internal Soft
				SetPixelV(lhdc, 1, 2, &HCAC7BF)
				SetPixelV(lhdc, 2, 1, &HCAC7BF)
				SetPixelV(lhdc, 2, lh - 2, &HCAC7BF)
				SetPixelV(lhdc, 1, lh - 3, &HCAC7BF)
				SetPixelV(lhdc, lw - 2, lh - 3, &HCAC7BF)
				SetPixelV(lhdc, lw - 3, lh - 2, &HCAC7BF)
				SetPixelV(lhdc, lw - 3, 1, &HCAC7BF)
				SetPixelV(lhdc, lw - 2, 2, &HCAC7BF)
			Case isState.statedisabled
				tempColor = &HEAF4F5
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tempColor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				lhdc = MyBase.hdc
				APIRectangle(lhdc, 0, 0, lw - 1, lh - 1, System.Drawing.ColorTranslator.FromOle(&HBAC7C9))
				tempColor = &HC7D5D8
				SetPixelV(lhdc, 0, 1, tempColor)
				SetPixelV(lhdc, 1, 1, tempColor)
				SetPixelV(lhdc, 1, 0, tempColor)
				SetPixelV(lhdc, 0, lh - 2, tempColor)
				SetPixelV(lhdc, 1, lh - 2, tempColor)
				SetPixelV(lhdc, 1, lh - 1, tempColor)
				SetPixelV(lhdc, lw - 1, 1, tempColor)
				SetPixelV(lhdc, lw - 2, 1, tempColor)
				SetPixelV(lhdc, lw - 2, 0, tempColor)
				SetPixelV(lhdc, lw - 1, lh - 2, tempColor)
				SetPixelV(lhdc, lw - 2, lh - 2, tempColor)
				SetPixelV(lhdc, lw - 2, lh - 1, tempColor)
		End Select
		
		Exit Sub
		
DrawWinXPButton_Error: 
	End Sub
	
	Private Sub DrawCustomWinXPButton(ByRef Mode As isState)
		'On Error GoTo DrawCustomWinXPButton_Error
		
		Dim tmpcolor As Integer
		Dim lh, lw As Integer
		Dim lhdc As Integer
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) : lw = MyBase.ClientRectangle.Width
		
		'Here, we know we will use custom colors
		Select Case Mode
			Case isState.statenormal, isState.stateDefaulted, isState.stateHot
				tmpcolor = m_lBackColor
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				'main gradient
				DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), 1, 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), 1, 2, 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				'Top Lines
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H5s)))
				APILine(1, 2, MyBase.ClientRectangle.Width - 1, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H2s)))
				'Bottom Lines
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H10s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H18s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)))
				'Border
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H80s))
				APILine(2, 0, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				APILine(MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'Border Pixels
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(m_lBackColor), -&H15s))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				
				If Mode = isState.stateDefaulted Or Mode = isState.stateHot Or (m_bFocused And m_bShowFocus) Then 'Or Ambient.DisplayAsDefault  Then
					tmpcolor = IIf((Mode = isState.stateHot), m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
					APILine(2, 1, lw - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H55s)))
					APILine(2, 2, lw - 2, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H45s)))
					APILine(2, lh - 3, lw - 2, lh - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H10s)))
					APILine(2, lh - 2, lw - 2, lh - 2, tmpcolor)
					DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H45s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H10s)), 1, 2, 3, lh - 5)
					DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H45s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H10s)), lw - 3, 2, lw - 1, lh - 5)
				End If
				
			Case isState.statePressed
				tmpcolor = m_lBackColor
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				lhdc = MyBase.hdc
				'Main
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), 2, 1, lw - 1, 4)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), 2, 4, lw - 1, lh - 8)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H5s)), 3, lh - 4, lw - 1, 3)
				'left
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H16s)), 1, 3, 2, lh - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H18s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), 2, 4, 3, lh - 7)
				'External Borders
				'tmpcolor = vbBlue
				SetPixelV(lhdc, 1, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)))
				SetPixelV(lhdc, 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)))
				SetPixelV(lhdc, 2, lh - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)))
				SetPixelV(lhdc, 1, lh - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H10s)))
				SetPixelV(lhdc, lw - 2, lh - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H12s)))
				SetPixelV(lhdc, lw - 3, lh - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H8s)))
				SetPixelV(lhdc, lw - 3, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)))
				SetPixelV(lhdc, lw - 2, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)))
				'Border
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(m_lBackColor), -&H80s))
				APILine(1, 0, lw - 1, 0, tmpcolor)
				APILine(0, 1, 0, lh - 1, tmpcolor)
				APILine(lw - 1, 1, lw - 1, lh - 1, tmpcolor)
				APILine(1, lh - 1, lw - 1, lh - 1, tmpcolor)
				'Corners
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(m_lBackColor), -&H60s))
				SetPixelV(lhdc, 1, 1, tmpcolor)
				SetPixelV(lhdc, 1, lh - 2, tmpcolor)
				SetPixelV(lhdc, lw - 2, 1, tmpcolor)
				SetPixelV(lhdc, lw - 2, lh - 2, tmpcolor)
			Case isState.statedisabled
				tmpcolor = m_lBackColor
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				lhdc = MyBase.hdc
				APIRectangle(lhdc, 0, 0, lw - 1, lh - 1, OffsetColor(System.Drawing.ColorTranslator.FromOle(m_lBackColor), -&H40s))
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(m_lBackColor), -&H35s))
				SetPixelV(lhdc, 0, 1, tmpcolor)
				SetPixelV(lhdc, 1, 1, tmpcolor)
				SetPixelV(lhdc, 1, 0, tmpcolor)
				SetPixelV(lhdc, 0, lh - 2, tmpcolor)
				SetPixelV(lhdc, 1, lh - 2, tmpcolor)
				SetPixelV(lhdc, 1, lh - 1, tmpcolor)
				SetPixelV(lhdc, lw - 1, 1, tmpcolor)
				SetPixelV(lhdc, lw - 2, 1, tmpcolor)
				SetPixelV(lhdc, lw - 2, 0, tmpcolor)
				SetPixelV(lhdc, lw - 1, lh - 2, tmpcolor)
				SetPixelV(lhdc, lw - 2, lh - 2, tmpcolor)
				SetPixelV(lhdc, lw - 2, lh - 1, tmpcolor)
		End Select
		
		Exit Sub
		
DrawCustomWinXPButton_Error: 
	End Sub
	
	Private Sub DrawMacOSXButton()
		'On Error GoTo DrawMacOSXButton_Error
		
		If m_iState = isState.stateHot Or m_iState = isState.stateDefaulted Then ' Or Ambient.DisplayAsDefault Then
			DrawMacOSXButtonHot()
		ElseIf m_iState = isState.statenormal Or m_iState = isState.statedisabled Then 
			
			If m_bFocused Then 'Or Ambient.DisplayAsDefault Then
				DrawMacOSXButtonHot()
			Else
				DrawMacOSXButtonNormal()
			End If
			
		Else 'If m_iState = statePressed Then
			DrawMacOSXButtonPressed()
		End If
		
		Exit Sub
		
DrawMacOSXButton_Error: 
	End Sub
	
	Private Sub DrawMacOSXButtonNormal()
		'On Error GoTo DrawMacOSXButtonNormal_Error
		
		Dim lhdc As Integer
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lhdc = MyBase.hdc
		'Variable vars (real into code)
		Dim lh, lw As Integer
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) : lw = MyBase.ClientRectangle.Width
		Dim tmph, tmpw As Integer
		Dim tmph1, tmpw1 As Integer
		'UPGRADE_ISSUE: UserControl property ctlButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		APIFillRectByCoords(hdc, 18, 11, lw - 34, lh - 19, &HEAE7E8)
		SetPixelV(lhdc, 6, 0, &HFEFEFE) : SetPixelV(lhdc, 7, 0, &HE6E6E6) : SetPixelV(lhdc, 8, 0, &HACACAC) : SetPixelV(lhdc, 9, 0, &H7A7A7A) : SetPixelV(lhdc, 10, 0, &H6C6C6C) : SetPixelV(lhdc, 11, 0, &H6B6B6B) : SetPixelV(lhdc, 12, 0, &H6F6F6F) : SetPixelV(lhdc, 13, 0, &H716F6F) : SetPixelV(lhdc, 14, 0, &H727070) : SetPixelV(lhdc, 15, 0, &H676866) : SetPixelV(lhdc, 16, 0, &H6C6D6B) : SetPixelV(lhdc, 17, 0, &H67696A) : SetPixelV(lhdc, 5, 1, &HEFEFEF) : SetPixelV(lhdc, 6, 1, &H939393) : SetPixelV(lhdc, 7, 1, &H676767) : SetPixelV(lhdc, 8, 1, &H797979) : SetPixelV(lhdc, 9, 1, &HB3B3B3) : SetPixelV(lhdc, 10, 1, &HDBDBDB) : SetPixelV(lhdc, 11, 1, &HEBEDEE) : SetPixelV(lhdc, 12, 1, &HF5F4F6) : SetPixelV(lhdc, 13, 1, &HF5F4F6) : SetPixelV(lhdc, 14, 1, &HF5F4F6) : SetPixelV(lhdc, 15, 1, &HF5F4F6) : SetPixelV(lhdc, 16, 1, &HF5F4F6) : SetPixelV(lhdc, 17, 1, &HF5F4F6)
		SetPixelV(lhdc, 3, 2, &HFEFEFE) : SetPixelV(lhdc, 4, 2, &HE5E5E5) : SetPixelV(lhdc, 5, 2, &H737373) : SetPixelV(lhdc, 6, 2, &H656565) : SetPixelV(lhdc, 7, 2, &H939393) : SetPixelV(lhdc, 8, 2, &HDCDCDC) : SetPixelV(lhdc, 9, 2, &HE9E9E9) : SetPixelV(lhdc, 10, 2, &HF2F1F3) : SetPixelV(lhdc, 11, 2, &HF3F2F4) : SetPixelV(lhdc, 12, 2, &HF2F1F3) : SetPixelV(lhdc, 13, 2, &HF3F2F4) : SetPixelV(lhdc, 14, 2, &HF2F1F3) : SetPixelV(lhdc, 15, 2, &HF3F2F4) : SetPixelV(lhdc, 16, 2, &HF2F1F3) : SetPixelV(lhdc, 17, 2, &HF3F2F4) : SetPixelV(lhdc, 3, 3, &HEEEEEE) : SetPixelV(lhdc, 4, 3, &H717171) : SetPixelV(lhdc, 5, 3, &H6C6C6C) : SetPixelV(lhdc, 6, 3, &H909090) : SetPixelV(lhdc, 7, 3, &HD2D2D2) : SetPixelV(lhdc, 8, 3, &HE3E3E3) : SetPixelV(lhdc, 9, 3, &HECECEC) : SetPixelV(lhdc, 10, 3, &HEDEDED) : SetPixelV(lhdc, 11, 3, &HEEEEEE) : SetPixelV(lhdc, 12, 3, &HEDEDED) : SetPixelV(lhdc, 13, 3, &HEEEEEE) : SetPixelV(lhdc, 14, 3, &HEDEDED) : SetPixelV(lhdc, 15, 3, &HEEEEEE) : SetPixelV(lhdc, 16, 3, &HEDEDED) : SetPixelV(lhdc, 17, 3, &HEEEEEE)
		SetPixelV(lhdc, 2, 4, &HFBFBFB) : SetPixelV(lhdc, 3, 4, &H858585) : SetPixelV(lhdc, 4, 4, &H686868) : SetPixelV(lhdc, 5, 4, &H959595) : SetPixelV(lhdc, 6, 4, &HB1B1B1) : SetPixelV(lhdc, 7, 4, &HDCDCDC) : SetPixelV(lhdc, 8, 4, &HE3E3E3) : SetPixelV(lhdc, 9, 4, &HE3E3E3) : SetPixelV(lhdc, 10, 4, &HEAEAEA) : SetPixelV(lhdc, 11, 4, &HEBEBEB) : SetPixelV(lhdc, 12, 4, &HEBEBEB) : SetPixelV(lhdc, 13, 4, &HEBEBEB) : SetPixelV(lhdc, 14, 4, &HEBEBEB) : SetPixelV(lhdc, 15, 4, &HEBEBEB) : SetPixelV(lhdc, 16, 4, &HEBEBEB) : SetPixelV(lhdc, 17, 4, &HEBEBEB)
		SetPixelV(lhdc, 1, 5, &HFEFEFE) : SetPixelV(lhdc, 2, 5, &HCACACA) : SetPixelV(lhdc, 3, 5, &H696969) : SetPixelV(lhdc, 4, 5, &H949494) : SetPixelV(lhdc, 5, 5, &HA6A6A6) : SetPixelV(lhdc, 6, 5, &HC5C5C5) : SetPixelV(lhdc, 7, 5, &HD8D8D8) : SetPixelV(lhdc, 8, 5, &HE0E0E0) : SetPixelV(lhdc, 9, 5, &HE1E1E1) : SetPixelV(lhdc, 10, 5, &HEAE9EA) : SetPixelV(lhdc, 11, 5, &HE7E7E7) : SetPixelV(lhdc, 12, 5, &HE9E7E8) : SetPixelV(lhdc, 13, 5, &HEBE8EA) : SetPixelV(lhdc, 14, 5, &HEAE7E9) : SetPixelV(lhdc, 15, 5, &HEBE8EA) : SetPixelV(lhdc, 16, 5, &HEAE7E9) : SetPixelV(lhdc, 17, 5, &HEBE8EA)
		SetPixelV(lhdc, 1, 6, &HF9F9F9) : SetPixelV(lhdc, 2, 6, &H808080) : SetPixelV(lhdc, 3, 6, &H878787) : SetPixelV(lhdc, 4, 6, &HA8A8A8) : SetPixelV(lhdc, 5, 6, &HB3B3B3) : SetPixelV(lhdc, 6, 6, &HC6C6C6) : SetPixelV(lhdc, 7, 6, &HDEDEDE) : SetPixelV(lhdc, 8, 6, &HE0E0E0) : SetPixelV(lhdc, 9, 6, &HE2E2E2) : SetPixelV(lhdc, 10, 6, &HE3E2E2) : SetPixelV(lhdc, 11, 6, &HE9EAE9) : SetPixelV(lhdc, 12, 6, &HE9E8E9) : SetPixelV(lhdc, 13, 6, &HEBE8EA) : SetPixelV(lhdc, 14, 6, &HEBE8EA) : SetPixelV(lhdc, 15, 6, &HEBE8EA) : SetPixelV(lhdc, 16, 6, &HEBE8EA) : SetPixelV(lhdc, 17, 6, &HEBE8EA)
		SetPixelV(lhdc, 1, 7, &HE8E8E8) : SetPixelV(lhdc, 2, 7, &H777777) : SetPixelV(lhdc, 3, 7, &H9B9B9B) : SetPixelV(lhdc, 4, 7, &HB1B1B1) : SetPixelV(lhdc, 5, 7, &HB9B9B9) : SetPixelV(lhdc, 6, 7, &HC5C5C5) : SetPixelV(lhdc, 7, 7, &HD6D6D6) : SetPixelV(lhdc, 8, 7, &HE0E0E0) : SetPixelV(lhdc, 9, 7, &HE0E0E0) : SetPixelV(lhdc, 10, 7, &HE7E7E7) : SetPixelV(lhdc, 11, 7, &HE7E7E7) : SetPixelV(lhdc, 12, 7, &HE9E9E9) : SetPixelV(lhdc, 13, 7, &HEAEAEA) : SetPixelV(lhdc, 14, 7, &HEAEAEA) : SetPixelV(lhdc, 15, 7, &HEAEAEA) : SetPixelV(lhdc, 16, 7, &HEAEAEA) : SetPixelV(lhdc, 17, 7, &HEAEAEA)
		SetPixelV(lhdc, 0, 8, &HFDFDFD) : SetPixelV(lhdc, 1, 8, &HC6C6C6) : SetPixelV(lhdc, 2, 8, &H7E7E7E) : SetPixelV(lhdc, 3, 8, &HABABAB) : SetPixelV(lhdc, 4, 8, &HC1C1C1) : SetPixelV(lhdc, 5, 8, &HC1C1C1) : SetPixelV(lhdc, 6, 8, &HCBCBCB) : SetPixelV(lhdc, 7, 8, &HCECECE) : SetPixelV(lhdc, 8, 8, &HD5D5D5) : SetPixelV(lhdc, 9, 8, &HD8D8D8) : SetPixelV(lhdc, 10, 8, &HDADADA) : SetPixelV(lhdc, 11, 8, &HDDDDDD) : SetPixelV(lhdc, 12, 8, &HDEDEDE) : SetPixelV(lhdc, 13, 8, &HE1E1E1) : SetPixelV(lhdc, 14, 8, &HE0E0E0) : SetPixelV(lhdc, 15, 8, &HE1E1E1) : SetPixelV(lhdc, 16, 8, &HE0E0E0) : SetPixelV(lhdc, 17, 8, &HE1E1E1)
		SetPixelV(lhdc, 0, 9, &HFAFAFA) : SetPixelV(lhdc, 1, 9, &HAEAEAE) : SetPixelV(lhdc, 2, 9, &H919191) : SetPixelV(lhdc, 3, 9, &HB9B9B9) : SetPixelV(lhdc, 4, 9, &HC4C4C4) : SetPixelV(lhdc, 5, 9, &HCECECE) : SetPixelV(lhdc, 6, 9, &HD1D1D1) : SetPixelV(lhdc, 7, 9, &HDADADA) : SetPixelV(lhdc, 8, 9, &HDCDCDC) : SetPixelV(lhdc, 9, 9, &HDBDBDB) : SetPixelV(lhdc, 10, 9, &HDFDFDF) : SetPixelV(lhdc, 11, 9, &HE1E3E1) : SetPixelV(lhdc, 12, 9, &HE2E3E2) : SetPixelV(lhdc, 13, 9, &HE5E2E3) : SetPixelV(lhdc, 14, 9, &HE5E2E3) : SetPixelV(lhdc, 15, 9, &HE5E2E3) : SetPixelV(lhdc, 16, 9, &HE5E2E3) : SetPixelV(lhdc, 17, 9, &HE5E2E3)
		SetPixelV(lhdc, 0, 10, &HF7F7F7) : SetPixelV(lhdc, 1, 10, &HA0A0A0) : SetPixelV(lhdc, 2, 10, &H999999) : SetPixelV(lhdc, 3, 10, &HC3C3C3) : SetPixelV(lhdc, 4, 10, &HC9C9C9) : SetPixelV(lhdc, 5, 10, &HD5D5D5) : SetPixelV(lhdc, 6, 10, &HD7D7D7) : SetPixelV(lhdc, 7, 10, &HDFDFDF) : SetPixelV(lhdc, 8, 10, &HE0E0E0) : SetPixelV(lhdc, 9, 10, &HE0E0E0) : SetPixelV(lhdc, 10, 10, &HE4E4E4) : SetPixelV(lhdc, 11, 10, &HE6E8E6) : SetPixelV(lhdc, 12, 10, &HE8E7E7) : SetPixelV(lhdc, 13, 10, &HEAE7E8) : SetPixelV(lhdc, 14, 10, &HEAE7E8) : SetPixelV(lhdc, 15, 10, &HEAE7E8) : SetPixelV(lhdc, 16, 10, &HEAE7E8) : SetPixelV(lhdc, 17, 10, &HEAE7E8)
		SetPixelV(lhdc, 0, 11, &HF5F5F5) : SetPixelV(lhdc, 1, 11, &HA3A3A3) : SetPixelV(lhdc, 2, 11, &H9B9B9B) : SetPixelV(lhdc, 3, 11, &HC6C6C6) : SetPixelV(lhdc, 4, 11, &HD3D3D3) : SetPixelV(lhdc, 5, 11, &HD6D6D6) : SetPixelV(lhdc, 6, 11, &HDDDDDD) : SetPixelV(lhdc, 7, 11, &HE1E1E1) : SetPixelV(lhdc, 8, 11, &HE3E3E3) : SetPixelV(lhdc, 9, 11, &HE6E6E6) : SetPixelV(lhdc, 10, 11, &HE7E8E7) : SetPixelV(lhdc, 11, 11, &HE9EAE9) : SetPixelV(lhdc, 12, 11, &HE8EAE9) : SetPixelV(lhdc, 13, 11, &HE8EBE9) : SetPixelV(lhdc, 14, 11, &HE8EBE9) : SetPixelV(lhdc, 15, 11, &HE8EBE9) : SetPixelV(lhdc, 16, 11, &HE8EBE9) : SetPixelV(lhdc, 17, 11, &HE8EBE9)
		SetPixelV(lhdc, 0, 12, &HF5F5F5) : SetPixelV(lhdc, 1, 12, &HAAAAAA) : SetPixelV(lhdc, 2, 12, &H8E8E8E) : SetPixelV(lhdc, 3, 12, &HD0D0D0) : SetPixelV(lhdc, 4, 12, &HDADADA) : SetPixelV(lhdc, 5, 12, &HDFDFDF) : SetPixelV(lhdc, 6, 12, &HE4E4E4) : SetPixelV(lhdc, 7, 12, &HE6E6E6) : SetPixelV(lhdc, 8, 12, &HE8E8E8) : SetPixelV(lhdc, 9, 12, &HECECEC) : SetPixelV(lhdc, 10, 12, &HEEEFEE) : SetPixelV(lhdc, 11, 12, &HEEF0EF) : SetPixelV(lhdc, 12, 12, &HEEF0EF) : SetPixelV(lhdc, 13, 12, &HEEF1EF) : SetPixelV(lhdc, 14, 12, &HEEF1EF) : SetPixelV(lhdc, 15, 12, &HEEF1EF) : SetPixelV(lhdc, 16, 12, &HEEF1EF) : SetPixelV(lhdc, 17, 12, &HEEF1EF)
		tmph = lh - 22
		SetPixelV(lhdc, 0, tmph + 12, &HF5F5F5) : SetPixelV(lhdc, 1, tmph + 12, &HAAAAAA) : SetPixelV(lhdc, 2, tmph + 12, &H8E8E8E) : SetPixelV(lhdc, 3, tmph + 12, &HD0D0D0) : SetPixelV(lhdc, 4, tmph + 12, &HDADADA) : SetPixelV(lhdc, 5, tmph + 12, &HDFDFDF) : SetPixelV(lhdc, 6, tmph + 12, &HE4E4E4) : SetPixelV(lhdc, 7, tmph + 12, &HE6E6E6) : SetPixelV(lhdc, 8, tmph + 12, &HE8E8E8) : SetPixelV(lhdc, 9, tmph + 12, &HECECEC) : SetPixelV(lhdc, 10, tmph + 12, &HEEEFEE) : SetPixelV(lhdc, 11, tmph + 12, &HEEF0EF) : SetPixelV(lhdc, 12, tmph + 12, &HEEF0EF) : SetPixelV(lhdc, 13, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, 14, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, 15, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, 16, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, 17, tmph + 12, &HEEF1EF)
		SetPixelV(lhdc, 0, tmph + 13, &HF7F7F7) : SetPixelV(lhdc, 1, tmph + 13, &HC2C2C2) : SetPixelV(lhdc, 2, tmph + 13, &H838383) : SetPixelV(lhdc, 3, tmph + 13, &HCFCFCF) : SetPixelV(lhdc, 4, tmph + 13, &HDEDEDE) : SetPixelV(lhdc, 5, tmph + 13, &HE3E3E3) : SetPixelV(lhdc, 6, tmph + 13, &HE8E8E8) : SetPixelV(lhdc, 7, tmph + 13, &HEAEAEA) : SetPixelV(lhdc, 8, tmph + 13, &HEDEDED) : SetPixelV(lhdc, 9, tmph + 13, &HF1F1F1) : SetPixelV(lhdc, 10, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 11, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 12, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 13, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 14, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 15, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 16, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, 17, tmph + 13, &HF2F2F2)
		SetPixelV(lhdc, 0, tmph + 14, &HFBFBFB) : SetPixelV(lhdc, 1, tmph + 14, &HE1E1E1) : SetPixelV(lhdc, 2, tmph + 14, &H818181) : SetPixelV(lhdc, 3, tmph + 14, &HABABAB) : SetPixelV(lhdc, 4, tmph + 14, &HDCDCDC) : SetPixelV(lhdc, 5, tmph + 14, &HE5E5E5) : SetPixelV(lhdc, 6, tmph + 14, &HEDEDED) : SetPixelV(lhdc, 7, tmph + 14, &HEFEFEF) : SetPixelV(lhdc, 8, tmph + 14, &HF1F1F1) : SetPixelV(lhdc, 9, tmph + 14, &HF4F4F4) : SetPixelV(lhdc, 10, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 11, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 12, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 13, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 14, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 15, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 16, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, 17, tmph + 14, &HF5F5F5)
		SetPixelV(lhdc, 0, tmph + 15, &HFEFEFE) : SetPixelV(lhdc, 1, tmph + 15, &HEDEDED) : SetPixelV(lhdc, 2, tmph + 15, &HA0A0A0) : SetPixelV(lhdc, 3, tmph + 15, &H898989) : SetPixelV(lhdc, 4, tmph + 15, &HDEDEDE) : SetPixelV(lhdc, 5, tmph + 15, &HE9E9E9) : SetPixelV(lhdc, 6, tmph + 15, &HEEEEEE) : SetPixelV(lhdc, 7, tmph + 15, &HF4F4F4) : SetPixelV(lhdc, 8, tmph + 15, &HF5F5F5) : SetPixelV(lhdc, 9, tmph + 15, &HFAFAFA) : SetPixelV(lhdc, 10, tmph + 15, &HFFFDFD) : SetPixelV(lhdc, 11, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, 12, tmph + 15, &HFFFDFD) : SetPixelV(lhdc, 13, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, 14, tmph + 15, &HFFFDFD) : SetPixelV(lhdc, 15, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, 16, tmph + 15, &HFFFDFD) : SetPixelV(lhdc, 17, tmph + 15, &HFFFEFE)
		SetPixelV(lhdc, 1, tmph + 16, &HF6F6F6) : SetPixelV(lhdc, 2, tmph + 16, &HD6D6D6) : SetPixelV(lhdc, 3, tmph + 16, &H7B7B7B) : SetPixelV(lhdc, 4, tmph + 16, &H8D8D8D) : SetPixelV(lhdc, 5, tmph + 16, &HE4E4E4) : SetPixelV(lhdc, 6, tmph + 16, &HF0F0F0) : SetPixelV(lhdc, 7, tmph + 16, &HF6F6F6) : SetPixelV(lhdc, 8, tmph + 16, &HFEFEFE) : SetPixelV(lhdc, 9, tmph + 16, &HFEFEFE) : SetPixelV(lhdc, 10, tmph + 16, &HFFFEFE) : SetPixelV(lhdc, 12, tmph + 16, &HFFFEFE) : SetPixelV(lhdc, 14, tmph + 16, &HFFFEFE) : SetPixelV(lhdc, 16, tmph + 16, &HFFFEFE)
		SetPixelV(lhdc, 1, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 2, tmph + 17, &HEDEDED) : SetPixelV(lhdc, 3, tmph + 17, &HBEBEBE) : SetPixelV(lhdc, 4, tmph + 17, &H727272) : SetPixelV(lhdc, 5, tmph + 17, &H898989) : SetPixelV(lhdc, 6, tmph + 17, &HEBEBEB) : SetPixelV(lhdc, 7, tmph + 17, &HF5F5F5) : SetPixelV(lhdc, 8, tmph + 17, &HFCFCFC) : SetPixelV(lhdc, 10, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 11, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 12, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 13, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 14, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 15, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 16, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 17, tmph + 17, &HFDFDFD)
		SetPixelV(lhdc, 2, tmph + 18, &HF9F9F9) : SetPixelV(lhdc, 3, tmph + 18, &HE6E6E6) : SetPixelV(lhdc, 4, tmph + 18, &HB9B9B9) : SetPixelV(lhdc, 5, tmph + 18, &H717171) : SetPixelV(lhdc, 6, tmph + 18, &H787878) : SetPixelV(lhdc, 7, tmph + 18, &HB6B6B6) : SetPixelV(lhdc, 8, tmph + 18, &HF7F7F7) : SetPixelV(lhdc, 9, tmph + 18, &HFCFCFC) : SetPixelV(lhdc, 10, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 11, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 12, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 13, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 14, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 15, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 16, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, 17, tmph + 18, &HFEFEFE)
		SetPixelV(lhdc, 2, tmph + 19, &HFEFEFE) : SetPixelV(lhdc, 3, tmph + 19, &HF8F8F8) : SetPixelV(lhdc, 4, tmph + 19, &HE6E6E6) : SetPixelV(lhdc, 5, tmph + 19, &HC8C8C8) : SetPixelV(lhdc, 6, tmph + 19, &H8E8E8E) : SetPixelV(lhdc, 7, tmph + 19, &H6C6C6C) : SetPixelV(lhdc, 8, tmph + 19, &H757575) : SetPixelV(lhdc, 9, tmph + 19, &H9F9F9F) : SetPixelV(lhdc, 10, tmph + 19, &HC7C7C7) : SetPixelV(lhdc, 11, tmph + 19, &HE9E9E9) : SetPixelV(lhdc, 12, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, 13, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, 14, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, 15, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, 16, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, 17, tmph + 19, &HFBFBFB)
		SetPixelV(lhdc, 3, tmph + 20, &HFEFEFE) : SetPixelV(lhdc, 4, tmph + 20, &HF9F9F9) : SetPixelV(lhdc, 5, tmph + 20, &HECECEC) : SetPixelV(lhdc, 6, tmph + 20, &HDADADA) : SetPixelV(lhdc, 7, tmph + 20, &HC1C1C1) : SetPixelV(lhdc, 8, tmph + 20, &H9D9D9D) : SetPixelV(lhdc, 9, tmph + 20, &H7B7B7B) : SetPixelV(lhdc, 10, tmph + 20, &H5E5E5E) : SetPixelV(lhdc, 11, tmph + 20, &H535353) : SetPixelV(lhdc, 12, tmph + 20, &H4D4D4D) : SetPixelV(lhdc, 13, tmph + 20, &H4B4B4B) : SetPixelV(lhdc, 14, tmph + 20, &H505050) : SetPixelV(lhdc, 15, tmph + 20, &H525252) : SetPixelV(lhdc, 16, tmph + 20, &H555555) : SetPixelV(lhdc, 17, tmph + 20, &H545454)
		SetPixelV(lhdc, 5, tmph + 21, &HFCFCFC) : SetPixelV(lhdc, 6, tmph + 21, &HF5F5F5) : SetPixelV(lhdc, 7, tmph + 21, &HEBEBEB) : SetPixelV(lhdc, 8, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, 9, tmph + 21, &HD6D6D6) : SetPixelV(lhdc, 10, tmph + 21, &HCECECE) : SetPixelV(lhdc, 11, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, 12, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 13, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 14, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 15, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 16, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, 17, tmph + 21, &HC5C5C5)
		SetPixelV(lhdc, 7, tmph + 22, &HFDFDFD) : SetPixelV(lhdc, 8, tmph + 22, &HF9F9F9) : SetPixelV(lhdc, 9, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, 10, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, 11, tmph + 22, &HEEEEEE) : SetPixelV(lhdc, 12, tmph + 22, &HEDEDED) : SetPixelV(lhdc, 13, tmph + 22, &HECECEC) : SetPixelV(lhdc, 14, tmph + 22, &HECECEC) : SetPixelV(lhdc, 15, tmph + 22, &HECECEC) : SetPixelV(lhdc, 16, tmph + 22, &HECECEC) : SetPixelV(lhdc, 17, tmph + 22, &HECECEC)
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, 0, &H67696A) : SetPixelV(lhdc, tmpw + 18, 0, &H666869) : SetPixelV(lhdc, tmpw + 19, 0, &H716F6F) : SetPixelV(lhdc, tmpw + 20, 0, &H6F6D6D) : SetPixelV(lhdc, tmpw + 21, 0, &H6F706E) : SetPixelV(lhdc, tmpw + 22, 0, &H727371) : SetPixelV(lhdc, tmpw + 23, 0, &H6E6E6E) : SetPixelV(lhdc, tmpw + 24, 0, &H707070) : SetPixelV(lhdc, tmpw + 25, 0, &HA6A6A6) : SetPixelV(lhdc, tmpw + 26, 0, &HEEEEEE) : SetPixelV(lhdc, tmpw + 34, 0, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 1, &HF5F4F6) : SetPixelV(lhdc, tmpw + 18, 1, &HF5F4F6) : SetPixelV(lhdc, tmpw + 19, 1, &HF5F4F6) : SetPixelV(lhdc, tmpw + 20, 1, &HF5F4F6) : SetPixelV(lhdc, tmpw + 21, 1, &HF4F3F5) : SetPixelV(lhdc, tmpw + 22, 1, &HF1F0F2) : SetPixelV(lhdc, tmpw + 23, 1, &HE0E0E0) : SetPixelV(lhdc, tmpw + 24, 1, &HC3C3C3) : SetPixelV(lhdc, tmpw + 25, 1, &H848484) : SetPixelV(lhdc, tmpw + 26, 1, &H6B6B6B) : SetPixelV(lhdc, tmpw + 27, 1, &HA0A0A0) : SetPixelV(lhdc, tmpw + 28, 1, &HF7F7F7) : SetPixelV(lhdc, tmpw + 34, 1, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 2, &HF3F2F4) : SetPixelV(lhdc, tmpw + 18, 2, &HF2F1F3) : SetPixelV(lhdc, tmpw + 19, 2, &HF3F2F4) : SetPixelV(lhdc, tmpw + 20, 2, &HF3F2F4) : SetPixelV(lhdc, tmpw + 21, 2, &HF0EFF1) : SetPixelV(lhdc, tmpw + 22, 2, &HF2F1F3) : SetPixelV(lhdc, tmpw + 23, 2, &HF6F6F6) : SetPixelV(lhdc, tmpw + 24, 2, &HE8E8E8) : SetPixelV(lhdc, tmpw + 25, 2, &HE0E0E0) : SetPixelV(lhdc, tmpw + 26, 2, &H999999) : SetPixelV(lhdc, tmpw + 27, 2, &H696969) : SetPixelV(lhdc, tmpw + 28, 2, &H717171) : SetPixelV(lhdc, tmpw + 29, 2, &HEBEBEB) : SetPixelV(lhdc, tmpw + 34, 2, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 3, &HEEEEEE) : SetPixelV(lhdc, tmpw + 18, 3, &HEDEDED) : SetPixelV(lhdc, tmpw + 19, 3, &HEEEEEE) : SetPixelV(lhdc, tmpw + 20, 3, &HEEEEEE) : SetPixelV(lhdc, tmpw + 21, 3, &HEEEEEE) : SetPixelV(lhdc, tmpw + 22, 3, &HEEEEEE) : SetPixelV(lhdc, tmpw + 23, 3, &HE9E9E9) : SetPixelV(lhdc, tmpw + 24, 3, &HEAEAEA) : SetPixelV(lhdc, tmpw + 25, 3, &HE7E7E7) : SetPixelV(lhdc, tmpw + 26, 3, &HD0D0D0) : SetPixelV(lhdc, tmpw + 27, 3, &H939393) : SetPixelV(lhdc, tmpw + 28, 3, &H727272) : SetPixelV(lhdc, tmpw + 29, 3, &H6F6F6F) : SetPixelV(lhdc, tmpw + 30, 3, &HEFEFEF) : SetPixelV(lhdc, tmpw + 34, 3, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 4, &HEBEBEB) : SetPixelV(lhdc, tmpw + 18, 4, &HEBEBEB) : SetPixelV(lhdc, tmpw + 19, 4, &HEBEBEB) : SetPixelV(lhdc, tmpw + 20, 4, &HEBEBEB) : SetPixelV(lhdc, tmpw + 21, 4, &HEDEDED) : SetPixelV(lhdc, tmpw + 22, 4, &HE6E6E6) : SetPixelV(lhdc, tmpw + 23, 4, &HE9E9E9) : SetPixelV(lhdc, tmpw + 24, 4, &HE6E6E6) : SetPixelV(lhdc, tmpw + 25, 4, &HDEDEDE) : SetPixelV(lhdc, tmpw + 26, 4, &HDCDCDC) : SetPixelV(lhdc, tmpw + 27, 4, &HB2B2B2) : SetPixelV(lhdc, tmpw + 28, 4, &H919191) : SetPixelV(lhdc, tmpw + 29, 4, &H6E6E6E) : SetPixelV(lhdc, tmpw + 30, 4, &H7F7F7F) : SetPixelV(lhdc, tmpw + 31, 4, &HFAFAFA) : SetPixelV(lhdc, tmpw + 34, 4, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 5, &HEBE8EA) : SetPixelV(lhdc, tmpw + 18, 5, &HEAE7E9) : SetPixelV(lhdc, tmpw + 19, 5, &HEBE8EA) : SetPixelV(lhdc, tmpw + 20, 5, &HEBE8EA) : SetPixelV(lhdc, tmpw + 21, 5, &HE5E8E6) : SetPixelV(lhdc, tmpw + 22, 5, &HE7EAE8) : SetPixelV(lhdc, tmpw + 23, 5, &HE5E5E5) : SetPixelV(lhdc, tmpw + 24, 5, &HE3E3E3) : SetPixelV(lhdc, tmpw + 25, 5, &HDFDFDF) : SetPixelV(lhdc, tmpw + 26, 5, &HDCDCDC) : SetPixelV(lhdc, tmpw + 27, 5, &HC3C3C3) : SetPixelV(lhdc, tmpw + 28, 5, &HA7A7A7) : SetPixelV(lhdc, tmpw + 29, 5, &H969696) : SetPixelV(lhdc, tmpw + 30, 5, &H717171) : SetPixelV(lhdc, tmpw + 31, 5, &HC5C5C5) : SetPixelV(lhdc, tmpw + 32, 5, &HFEFEFE) : SetPixelV(lhdc, tmpw + 34, 5, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 6, &HEBE8EA) : SetPixelV(lhdc, tmpw + 18, 6, &HEBE8EA) : SetPixelV(lhdc, tmpw + 19, 6, &HEBE8EA) : SetPixelV(lhdc, tmpw + 20, 6, &HEBE8EA) : SetPixelV(lhdc, tmpw + 21, 6, &HE8EBE9) : SetPixelV(lhdc, tmpw + 22, 6, &HE3E6E4) : SetPixelV(lhdc, tmpw + 23, 6, &HE5E5E5) : SetPixelV(lhdc, tmpw + 24, 6, &HE2E2E2) : SetPixelV(lhdc, tmpw + 25, 6, &HE0E0E0) : SetPixelV(lhdc, tmpw + 26, 6, &HDADADA) : SetPixelV(lhdc, tmpw + 27, 6, &HC7C7C7) : SetPixelV(lhdc, tmpw + 28, 6, &HB5B5B5) : SetPixelV(lhdc, tmpw + 29, 6, &HA6A6A6) : SetPixelV(lhdc, tmpw + 30, 6, &H8C8C8C) : SetPixelV(lhdc, tmpw + 31, 6, &H808080) : SetPixelV(lhdc, tmpw + 32, 6, &HF8F8F8) : SetPixelV(lhdc, tmpw + 34, 6, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 7, &HEAEAEA) : SetPixelV(lhdc, tmpw + 18, 7, &HEAEAEA) : SetPixelV(lhdc, tmpw + 19, 7, &HEAEAEA) : SetPixelV(lhdc, tmpw + 20, 7, &HEAEAEA) : SetPixelV(lhdc, tmpw + 21, 7, &HE9E6E8) : SetPixelV(lhdc, tmpw + 22, 7, &HE9E6E8) : SetPixelV(lhdc, tmpw + 23, 7, &HE4E4E4) : SetPixelV(lhdc, tmpw + 24, 7, &HE2E2E2) : SetPixelV(lhdc, tmpw + 25, 7, &HDFDFDF) : SetPixelV(lhdc, tmpw + 26, 7, &HD7D7D7) : SetPixelV(lhdc, tmpw + 27, 7, &HC4C4C4) : SetPixelV(lhdc, tmpw + 28, 7, &HB7B7B7) : SetPixelV(lhdc, tmpw + 29, 7, &HB4B5B3) : SetPixelV(lhdc, tmpw + 30, 7, &H9D9E9C) : SetPixelV(lhdc, tmpw + 31, 7, &H777777) : SetPixelV(lhdc, tmpw + 32, 7, &HE7E7E7) : SetPixelV(lhdc, tmpw + 34, 7, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 8, &HE1E1E1) : SetPixelV(lhdc, tmpw + 18, 8, &HE0E0E0) : SetPixelV(lhdc, tmpw + 19, 8, &HE1E1E1) : SetPixelV(lhdc, tmpw + 20, 8, &HE1E1E1) : SetPixelV(lhdc, tmpw + 21, 8, &HDFDCDE) : SetPixelV(lhdc, tmpw + 22, 8, &HDDDADC) : SetPixelV(lhdc, tmpw + 23, 8, &HDBDBDB) : SetPixelV(lhdc, tmpw + 24, 8, &HD6D6D6) : SetPixelV(lhdc, tmpw + 25, 8, &HD5D5D5) : SetPixelV(lhdc, tmpw + 26, 8, &HD1D1D1) : SetPixelV(lhdc, tmpw + 27, 8, &HC9C9C9) : SetPixelV(lhdc, tmpw + 28, 8, &HC4C4C4) : SetPixelV(lhdc, tmpw + 29, 8, &HC0C1BF) : SetPixelV(lhdc, tmpw + 30, 8, &HAFB0AE) : SetPixelV(lhdc, tmpw + 31, 8, &H818181) : SetPixelV(lhdc, tmpw + 32, 8, &HC3C3C3) : SetPixelV(lhdc, tmpw + 33, 8, &HFDFDFD) : SetPixelV(lhdc, tmpw + 34, 8, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 9, &HE5E2E3) : SetPixelV(lhdc, tmpw + 18, 9, &HE5E2E3) : SetPixelV(lhdc, tmpw + 19, 9, &HE5E2E3) : SetPixelV(lhdc, tmpw + 20, 9, &HE5E2E3) : SetPixelV(lhdc, tmpw + 21, 9, &HE1E1E1) : SetPixelV(lhdc, tmpw + 22, 9, &HE1E1E1) : SetPixelV(lhdc, tmpw + 23, 9, &HE1E1E1) : SetPixelV(lhdc, tmpw + 24, 9, &HDDDDDD) : SetPixelV(lhdc, tmpw + 25, 9, &HDBDBDB) : SetPixelV(lhdc, tmpw + 26, 9, &HD8D8D8) : SetPixelV(lhdc, tmpw + 27, 9, &HD2D2D2) : SetPixelV(lhdc, tmpw + 28, 9, &HCBCBCB) : SetPixelV(lhdc, tmpw + 29, 9, &HC4C4C4) : SetPixelV(lhdc, tmpw + 30, 9, &HBABABA) : SetPixelV(lhdc, tmpw + 31, 9, &H989898) : SetPixelV(lhdc, tmpw + 32, 9, &HA6A6A6) : SetPixelV(lhdc, tmpw + 33, 9, &HF9F9F9) : SetPixelV(lhdc, tmpw + 34, 9, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 10, &HEAE7E8) : SetPixelV(lhdc, tmpw + 18, 10, &HEAE7E8) : SetPixelV(lhdc, tmpw + 19, 10, &HEAE7E8) : SetPixelV(lhdc, tmpw + 20, 10, &HEAE7E8) : SetPixelV(lhdc, tmpw + 21, 10, &HE7E7E7) : SetPixelV(lhdc, tmpw + 22, 10, &HE6E6E6) : SetPixelV(lhdc, tmpw + 23, 10, &HE4E4E4) : SetPixelV(lhdc, tmpw + 24, 10, &HE0E0E0) : SetPixelV(lhdc, tmpw + 25, 10, &HE0E0E0) : SetPixelV(lhdc, tmpw + 26, 10, &HDEDEDE) : SetPixelV(lhdc, tmpw + 27, 10, &HD9D9D9) : SetPixelV(lhdc, tmpw + 28, 10, &HD3D3D3) : SetPixelV(lhdc, tmpw + 29, 10, &HCCCCCC) : SetPixelV(lhdc, tmpw + 30, 10, &HC3C3C3) : SetPixelV(lhdc, tmpw + 31, 10, &HA3A3A3) : SetPixelV(lhdc, tmpw + 32, 10, &H9C9C9C) : SetPixelV(lhdc, tmpw + 33, 10, &HF6F6F6) : SetPixelV(lhdc, tmpw + 34, 10, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 11, &HE8EBE9) : SetPixelV(lhdc, tmpw + 18, 11, &HE8EBE9) : SetPixelV(lhdc, tmpw + 19, 11, &HE8EBE9) : SetPixelV(lhdc, tmpw + 20, 11, &HE8EBE9) : SetPixelV(lhdc, tmpw + 21, 11, &HE9EAE8) : SetPixelV(lhdc, tmpw + 22, 11, &HE8E9E7) : SetPixelV(lhdc, tmpw + 23, 11, &HE9E9E9) : SetPixelV(lhdc, tmpw + 24, 11, &HE5E5E5) : SetPixelV(lhdc, tmpw + 25, 11, &HE4E4E4) : SetPixelV(lhdc, tmpw + 26, 11, &HE2E2E2) : SetPixelV(lhdc, tmpw + 27, 11, &HDBDBDB) : SetPixelV(lhdc, tmpw + 28, 11, &HD9D9D9) : SetPixelV(lhdc, tmpw + 29, 11, &HD1D1D1) : SetPixelV(lhdc, tmpw + 30, 11, &HC8C8C8) : SetPixelV(lhdc, tmpw + 31, 11, &HA4A4A4) : SetPixelV(lhdc, tmpw + 32, 11, &HA2A2A2) : SetPixelV(lhdc, tmpw + 33, 11, &HF4F4F4) : SetPixelV(lhdc, tmpw + 34, 11, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 18, 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 19, 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 20, 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 21, 12, &HEEEFED) : SetPixelV(lhdc, tmpw + 22, 12, &HEFF0EE) : SetPixelV(lhdc, tmpw + 23, 12, &HEEEEEE) : SetPixelV(lhdc, tmpw + 24, 12, &HECECEC) : SetPixelV(lhdc, tmpw + 25, 12, &HEAEAEA) : SetPixelV(lhdc, tmpw + 26, 12, &HE7E7E7) : SetPixelV(lhdc, tmpw + 27, 12, &HE2E2E2) : SetPixelV(lhdc, tmpw + 28, 12, &HDFDFDF) : SetPixelV(lhdc, tmpw + 29, 12, &HD8D8D8) : SetPixelV(lhdc, tmpw + 30, 12, &HD4D4D4) : SetPixelV(lhdc, tmpw + 31, 12, &H999999) : SetPixelV(lhdc, tmpw + 32, 12, &HAFAFAF) : SetPixelV(lhdc, tmpw + 33, 12, &HF5F5F5) : SetPixelV(lhdc, tmpw + 34, 12, &HFFFFFFFF)
		tmph = lh - 22
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 18, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 19, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 20, tmph + 12, &HEEF1EF) : SetPixelV(lhdc, tmpw + 21, tmph + 12, &HEEEFED) : SetPixelV(lhdc, tmpw + 22, tmph + 12, &HEFF0EE) : SetPixelV(lhdc, tmpw + 23, tmph + 12, &HEEEEEE) : SetPixelV(lhdc, tmpw + 24, tmph + 12, &HECECEC) : SetPixelV(lhdc, tmpw + 25, tmph + 12, &HEAEAEA) : SetPixelV(lhdc, tmpw + 26, tmph + 12, &HE7E7E7) : SetPixelV(lhdc, tmpw + 27, tmph + 12, &HE2E2E2) : SetPixelV(lhdc, tmpw + 28, tmph + 12, &HDFDFDF) : SetPixelV(lhdc, tmpw + 29, tmph + 12, &HD8D8D8) : SetPixelV(lhdc, tmpw + 30, tmph + 12, &HD4D4D4) : SetPixelV(lhdc, tmpw + 31, tmph + 12, &H999999) : SetPixelV(lhdc, tmpw + 32, tmph + 12, &HAFAFAF) : SetPixelV(lhdc, tmpw + 33, tmph + 12, &HF5F5F5)
		SetPixelV(lhdc, tmpw + 17, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 18, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 19, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 20, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 21, tmph + 13, &HF5F4F6) : SetPixelV(lhdc, tmpw + 22, tmph + 13, &HF0EFF1) : SetPixelV(lhdc, tmpw + 23, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 24, tmph + 13, &HF2F2F2) : SetPixelV(lhdc, tmpw + 25, tmph + 13, &HECECEC) : SetPixelV(lhdc, tmpw + 26, tmph + 13, &HEAEAEA) : SetPixelV(lhdc, tmpw + 27, tmph + 13, &HEBEBEB) : SetPixelV(lhdc, tmpw + 28, tmph + 13, &HE3E3E3) : SetPixelV(lhdc, tmpw + 29, tmph + 13, &HDEDEDE) : SetPixelV(lhdc, tmpw + 30, tmph + 13, &HD1D1D1) : SetPixelV(lhdc, tmpw + 31, tmph + 13, &H8A8A8A) : SetPixelV(lhdc, tmpw + 32, tmph + 13, &HD5D5D5) : SetPixelV(lhdc, tmpw + 33, tmph + 13, &HF8F8F8)
		SetPixelV(lhdc, tmpw + 17, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, tmpw + 18, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, tmpw + 19, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, tmpw + 20, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, tmpw + 21, tmph + 14, &HF8F7F9) : SetPixelV(lhdc, tmpw + 22, tmph + 14, &HF7F6F8) : SetPixelV(lhdc, tmpw + 23, tmph + 14, &HF7F7F7) : SetPixelV(lhdc, tmpw + 24, tmph + 14, &HF5F5F5) : SetPixelV(lhdc, tmpw + 25, tmph + 14, &HEFEFEF) : SetPixelV(lhdc, tmpw + 26, tmph + 14, &HEEEEEE) : SetPixelV(lhdc, tmpw + 27, tmph + 14, &HECECEC) : SetPixelV(lhdc, tmpw + 28, tmph + 14, &HE5E5E5) : SetPixelV(lhdc, tmpw + 29, tmph + 14, &HDEDEDE) : SetPixelV(lhdc, tmpw + 30, tmph + 14, &HB3B3B3) : SetPixelV(lhdc, tmpw + 31, tmph + 14, &H808080) : SetPixelV(lhdc, tmpw + 32, tmph + 14, &HE8E8E8) : SetPixelV(lhdc, tmpw + 33, tmph + 14, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, tmpw + 18, tmph + 15, &HFFFDFD) : SetPixelV(lhdc, tmpw + 19, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, tmpw + 20, tmph + 15, &HFFFEFE) : SetPixelV(lhdc, tmpw + 21, tmph + 15, &HFBFBFB) : SetPixelV(lhdc, tmpw + 22, tmph + 15, &HFCFCFC) : SetPixelV(lhdc, tmpw + 23, tmph + 15, &HFEFEFE) : SetPixelV(lhdc, tmpw + 24, tmph + 15, &HF8F8F8) : SetPixelV(lhdc, tmpw + 25, tmph + 15, &HF7F7F7) : SetPixelV(lhdc, tmpw + 26, tmph + 15, &HF5F5F5) : SetPixelV(lhdc, tmpw + 27, tmph + 15, &HEDEDED) : SetPixelV(lhdc, tmpw + 28, tmph + 15, &HEAEAEA) : SetPixelV(lhdc, tmpw + 29, tmph + 15, &HE0E0E0) : SetPixelV(lhdc, tmpw + 30, tmph + 15, &H8D8D8D) : SetPixelV(lhdc, tmpw + 31, tmph + 15, &HBABABA) : SetPixelV(lhdc, tmpw + 32, tmph + 15, &HF1F1F1)
		SetPixelV(lhdc, tmpw + 18, tmph + 16, &HFFFEFE) : SetPixelV(lhdc, tmpw + 22, tmph + 16, &HFEFEFE) : SetPixelV(lhdc, tmpw + 23, tmph + 16, &HFEFEFE) : SetPixelV(lhdc, tmpw + 25, tmph + 16, &HFCFCFC) : SetPixelV(lhdc, tmpw + 26, tmph + 16, &HF6F6F6) : SetPixelV(lhdc, tmpw + 27, tmph + 16, &HF2F2F2) : SetPixelV(lhdc, tmpw + 28, tmph + 16, &HE7E7E7) : SetPixelV(lhdc, tmpw + 29, tmph + 16, &H989898) : SetPixelV(lhdc, tmpw + 30, tmph + 16, &H828282) : SetPixelV(lhdc, tmpw + 31, tmph + 16, &HE2E2E2) : SetPixelV(lhdc, tmpw + 32, tmph + 16, &HF9F9F9)
		SetPixelV(lhdc, tmpw + 17, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, tmpw + 18, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, tmpw + 19, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, tmpw + 20, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, tmpw + 21, tmph + 17, &HFEFEFE) : SetPixelV(lhdc, tmpw + 23, tmph + 17, &HFEFEFE) : SetPixelV(lhdc, tmpw + 25, tmph + 17, &HFEFEFE) : SetPixelV(lhdc, tmpw + 26, tmph + 17, &HF6F6F6) : SetPixelV(lhdc, tmpw + 27, tmph + 17, &HF1F1F1) : SetPixelV(lhdc, tmpw + 28, tmph + 17, &H979797) : SetPixelV(lhdc, tmpw + 29, tmph + 17, &H6F6F6F) : SetPixelV(lhdc, tmpw + 30, tmph + 17, &HD2D2D2) : SetPixelV(lhdc, tmpw + 31, tmph + 17, &HF2F2F2) : SetPixelV(lhdc, tmpw + 32, tmph + 17, &HFEFEFE)
		SetPixelV(lhdc, tmpw + 17, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, tmpw + 18, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, tmpw + 19, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, tmpw + 20, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, tmpw + 22, tmph + 18, &HFDFDFD) : SetPixelV(lhdc, tmpw + 23, tmph + 18, &HFEFEFE) : SetPixelV(lhdc, tmpw + 24, tmph + 18, &HFDFDFD) : SetPixelV(lhdc, tmpw + 25, tmph + 18, &HFCFCFC) : SetPixelV(lhdc, tmpw + 26, tmph + 18, &HC5C5C5) : SetPixelV(lhdc, tmpw + 27, tmph + 18, &H838383) : SetPixelV(lhdc, tmpw + 28, tmph + 18, &H6F6F6F) : SetPixelV(lhdc, tmpw + 29, tmph + 18, &HC8C8C8) : SetPixelV(lhdc, tmpw + 30, tmph + 18, &HEBEBEB) : SetPixelV(lhdc, tmpw + 31, tmph + 18, &HFCFCFC)
		SetPixelV(lhdc, tmpw + 17, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, tmpw + 18, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, tmpw + 19, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, tmpw + 20, tmph + 19, &HFBFBFB) : SetPixelV(lhdc, tmpw + 21, tmph + 19, &HFAFAFA) : SetPixelV(lhdc, tmpw + 22, tmph + 19, &HEFEFEF) : SetPixelV(lhdc, tmpw + 23, tmph + 19, &HD0D0D0) : SetPixelV(lhdc, tmpw + 24, tmph + 19, &HA3A3A3) : SetPixelV(lhdc, tmpw + 25, tmph + 19, &H7E7E7E) : SetPixelV(lhdc, tmpw + 26, tmph + 19, &H6A6A6A) : SetPixelV(lhdc, tmpw + 27, tmph + 19, &H8F8F8F) : SetPixelV(lhdc, tmpw + 28, tmph + 19, &HCDCDCD) : SetPixelV(lhdc, tmpw + 29, tmph + 19, &HE8E8E8) : SetPixelV(lhdc, tmpw + 30, tmph + 19, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 20, &H545454) : SetPixelV(lhdc, tmpw + 18, tmph + 20, &H555555) : SetPixelV(lhdc, tmpw + 19, tmph + 20, &H525252) : SetPixelV(lhdc, tmpw + 20, tmph + 20, &H505050) : SetPixelV(lhdc, tmpw + 21, tmph + 20, &H535353) : SetPixelV(lhdc, tmpw + 22, tmph + 20, &H525252) : SetPixelV(lhdc, tmpw + 23, tmph + 20, &H616161) : SetPixelV(lhdc, tmpw + 24, tmph + 20, &H7A7A7A) : SetPixelV(lhdc, tmpw + 25, tmph + 20, &HA3A3A3) : SetPixelV(lhdc, tmpw + 26, tmph + 20, &HC5C5C5) : SetPixelV(lhdc, tmpw + 27, tmph + 20, &HDADADA) : SetPixelV(lhdc, tmpw + 28, tmph + 20, &HEDEDED) : SetPixelV(lhdc, tmpw + 29, tmph + 20, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 18, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 19, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 20, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 21, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 22, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, tmpw + 23, tmph + 21, &HCECECE) : SetPixelV(lhdc, tmpw + 24, tmph + 21, &HD7D7D7) : SetPixelV(lhdc, tmpw + 25, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, tmpw + 26, tmph + 21, &HECECEC) : SetPixelV(lhdc, tmpw + 27, tmph + 21, &HF6F6F6) : SetPixelV(lhdc, tmpw + 28, tmph + 21, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 18, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 19, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 20, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 21, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 22, tmph + 22, &HEDEDED) : SetPixelV(lhdc, tmpw + 23, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, tmpw + 24, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, tmpw + 25, tmph + 22, &HFAFAFA) : SetPixelV(lhdc, tmpw + 26, tmph + 22, &HFDFDFD)
		'Vlines
		tmph = 11 : tmph1 = lh - 10 : tmpw = lw - 34
		APILine(0, tmph, 0, tmph1, &HF7F7F7) : APILine(1, tmph, 1, tmph1, &HA0A0A0) : APILine(2, tmph, 2, tmph1, &H999999) : APILine(3, tmph, 3, tmph1, &HC3C3C3)
		APILine(4, tmph, 4, tmph1, &HC9C9C9) : APILine(5, tmph, 5, tmph1, &HD5D5D5) : APILine(6, tmph, 6, tmph1, &HD7D7D7) : APILine(7, tmph, 7, tmph1, &HDFDFDF)
		APILine(8, tmph, 8, tmph1, &HE0E0E0) : APILine(9, tmph, 9, tmph1, &HE0E0E0) : APILine(10, tmph, 10, tmph1, &HE4E4E4) : APILine(11, tmph, 11, tmph1, &HE6E8E6)
		APILine(12, tmph, 12, tmph1, &HE8E7E7) : APILine(13, tmph, 13, tmph1, &HEAE7E8) : APILine(14, tmph, 14, tmph1, &HEAE7E8) : APILine(15, tmph, 15, tmph1, &HEAE7E8)
		APILine(16, tmph, 16, tmph1, &HEAE7E8) : APILine(17, tmph, 17, tmph1, &HEAE7E8) : APILine(tmpw + 17, tmph, tmpw + 17, tmph1, &HEAE7E8) : APILine(tmpw + 18, tmph, tmpw + 18, tmph1, &HEAE7E8)
		APILine(tmpw + 19, tmph, tmpw + 19, tmph1, &HEAE7E8) : APILine(tmpw + 20, tmph, tmpw + 20, tmph1, &HEAE7E8) : APILine(tmpw + 21, tmph, tmpw + 21, tmph1, &HE7E7E7)
		APILine(tmpw + 22, tmph, tmpw + 22, tmph1, &HE6E6E6) : APILine(tmpw + 23, tmph, tmpw + 23, tmph1, &HE4E4E4) : APILine(tmpw + 24, tmph, tmpw + 24, tmph1, &HE0E0E0)
		APILine(tmpw + 25, tmph, tmpw + 25, tmph1, &HE0E0E0) : APILine(tmpw + 26, tmph, tmpw + 26, tmph1, &HDEDEDE) : APILine(tmpw + 27, tmph, tmpw + 27, tmph1, &HD9D9D9)
		APILine(tmpw + 28, tmph, tmpw + 28, tmph1, &HD3D3D3) : APILine(tmpw + 29, tmph, tmpw + 29, tmph1, &HCCCCCC) : APILine(tmpw + 30, tmph, tmpw + 30, tmph1, &HC3C3C3)
		APILine(tmpw + 31, tmph, tmpw + 31, tmph1, &HA3A3A3) : APILine(tmpw + 32, tmph, tmpw + 32, tmph1, &H9C9C9C) : APILine(tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6)
		'HLines
		APILine(17, 0, lw - 17, 0, &H67696A)
		APILine(17, 1, lw - 17, 1, &HF5F4F6)
		APILine(17, 2, lw - 17, 2, &HF3F2F4)
		APILine(17, 3, lw - 17, 3, &HEEEEEE)
		APILine(17, 4, lw - 17, 4, &HEBEBEB)
		APILine(17, 5, lw - 17, 5, &HEBE8EA)
		APILine(17, 6, lw - 17, 6, &HEBE8EA)
		APILine(17, 7, lw - 17, 7, &HEAEAEA)
		APILine(17, 8, lw - 17, 8, &HE1E1E1)
		APILine(17, 9, lw - 17, 9, &HE5E2E3)
		APILine(17, 10, lw - 17, 10, &HEAE7E8)
		APILine(17, 11, lw - 17, 11, &HE8EBE9)
		tmph = lh - 22
		APILine(17, tmph + 11, lw - 17, tmph + 11, &HE8EBE9)
		APILine(17, tmph + 12, lw - 17, tmph + 12, &HEEF1EF)
		APILine(17, tmph + 13, lw - 17, tmph + 13, &HF2F2F2)
		APILine(17, tmph + 14, lw - 17, tmph + 14, &HF5F5F5)
		APILine(17, tmph + 15, lw - 17, tmph + 15, &HFFFEFE)
		APILine(17, tmph + 16, lw - 17, tmph + 16, &HFFFFFF)
		APILine(17, tmph + 17, lw - 17, tmph + 17, &HFDFDFD)
		APILine(17, tmph + 18, lw - 17, tmph + 18, &HFEFEFE)
		APILine(17, tmph + 19, lw - 17, tmph + 19, &HFBFBFB)
		APILine(17, tmph + 20, lw - 17, tmph + 20, &H545454)
		APILine(17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5)
		APILine(17, tmph + 22, lw - 17, tmph + 22, &HECECEC)
		Exit Sub
		
DrawMacOSXButtonNormal_Error: 
	End Sub
	
	Private Sub DrawMacOSXButtonHot()
		'On Error GoTo DrawMacOSXButtonHot_Error
		
		Dim lhdc As Integer
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lhdc = MyBase.hdc
		'Variable vars (real into code)
		Dim lh, lw As Integer
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) : lw = MyBase.ClientRectangle.Width
		Dim tmph, tmpw As Integer
		Dim tmph1, tmpw1 As Integer
		'UPGRADE_ISSUE: UserControl property ctlButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		APIFillRectByCoords(hdc, 18, 11, lw - 34, lh - 19, &HE2A66A)
		SetPixelV(lhdc, 6, 0, &HFEFEFE) : SetPixelV(lhdc, 7, 0, &HE6E5E5) : SetPixelV(lhdc, 8, 0, &HA9A5A5) : SetPixelV(lhdc, 9, 0, &H6C5E5E) : SetPixelV(lhdc, 10, 0, &H482729) : SetPixelV(lhdc, 11, 0, &H370D0C) : SetPixelV(lhdc, 12, 0, &H370706) : SetPixelV(lhdc, 13, 0, &H360605) : SetPixelV(lhdc, 14, 0, &H3A0606) : SetPixelV(lhdc, 15, 0, &H410807) : SetPixelV(lhdc, 16, 0, &H450707) : SetPixelV(lhdc, 17, 0, &H450608)
		SetPixelV(lhdc, 5, 1, &HF0EFEF) : SetPixelV(lhdc, 6, 1, &HA38A8C) : SetPixelV(lhdc, 7, 1, &H6E342F) : SetPixelV(lhdc, 8, 1, &H661F1A) : SetPixelV(lhdc, 9, 1, &H9B6A63) : SetPixelV(lhdc, 10, 1, &HC9A29D) : SetPixelV(lhdc, 11, 1, &HE2BFBD) : SetPixelV(lhdc, 12, 1, &HE8C9C6) : SetPixelV(lhdc, 13, 1, &HEFD3CC) : SetPixelV(lhdc, 14, 1, &HEFD3CC) : SetPixelV(lhdc, 15, 1, &HF0D5C9) : SetPixelV(lhdc, 16, 1, &HF0D5C9) : SetPixelV(lhdc, 17, 1, &HF1D4C9)
		SetPixelV(lhdc, 3, 2, &HFEFEFE) : SetPixelV(lhdc, 4, 2, &HE5E5E5) : SetPixelV(lhdc, 5, 2, &H755E5E) : SetPixelV(lhdc, 6, 2, &H41070C) : SetPixelV(lhdc, 7, 2, &H7F2D28) : SetPixelV(lhdc, 8, 2, &HEC9892) : SetPixelV(lhdc, 9, 2, &HECB6AF) : SetPixelV(lhdc, 10, 2, &HE3BBB6) : SetPixelV(lhdc, 11, 2, &HE3C0BD) : SetPixelV(lhdc, 12, 2, &HE1C2BF) : SetPixelV(lhdc, 13, 2, &HDFC3BC) : SetPixelV(lhdc, 14, 2, &HDFC3BC) : SetPixelV(lhdc, 15, 2, &HE4C9BD) : SetPixelV(lhdc, 16, 2, &HE4C9BD) : SetPixelV(lhdc, 17, 2, &HE5C8BD)
		SetPixelV(lhdc, 3, 3, &HEEEEEE) : SetPixelV(lhdc, 4, 3, &H8A5A5A) : SetPixelV(lhdc, 5, 3, &H7A0702) : SetPixelV(lhdc, 6, 3, &H901501) : SetPixelV(lhdc, 7, 3, &HC38365) : SetPixelV(lhdc, 8, 3, &HE3B08F) : SetPixelV(lhdc, 9, 3, &HE1B394) : SetPixelV(lhdc, 10, 3, &HE5B798) : SetPixelV(lhdc, 11, 3, &HE6BC99) : SetPixelV(lhdc, 12, 3, &HE7BD9A) : SetPixelV(lhdc, 13, 3, &HE4BC99) : SetPixelV(lhdc, 14, 3, &HE7BF9C) : SetPixelV(lhdc, 15, 3, &HE9C1A1) : SetPixelV(lhdc, 16, 3, &HE8C0A1) : SetPixelV(lhdc, 17, 3, &HE8C0A1)
		SetPixelV(lhdc, 2, 4, &HFBFBFB) : SetPixelV(lhdc, 3, 4, &H897879) : SetPixelV(lhdc, 4, 4, &H4D0909) : SetPixelV(lhdc, 5, 4, &H951905) : SetPixelV(lhdc, 6, 4, &HBF422E) : SetPixelV(lhdc, 7, 4, &HD49475) : SetPixelV(lhdc, 8, 4, &HD7A483) : SetPixelV(lhdc, 9, 4, &HDAAC8D) : SetPixelV(lhdc, 10, 4, &HDBAD8E) : SetPixelV(lhdc, 11, 4, &HD9AF8C) : SetPixelV(lhdc, 12, 4, &HDCB28F) : SetPixelV(lhdc, 13, 4, &HDDB592) : SetPixelV(lhdc, 14, 4, &HDCB491) : SetPixelV(lhdc, 15, 4, &HDFB797) : SetPixelV(lhdc, 16, 4, &HE0B898) : SetPixelV(lhdc, 17, 4, &HE0B898)
		SetPixelV(lhdc, 1, 5, &HFEFEFE) : SetPixelV(lhdc, 2, 5, &HCDC9C9) : SetPixelV(lhdc, 3, 5, &H882517) : SetPixelV(lhdc, 4, 5, &H922100) : SetPixelV(lhdc, 5, 5, &HA13A00) : SetPixelV(lhdc, 6, 5, &HD57333) : SetPixelV(lhdc, 7, 5, &HDFA36F) : SetPixelV(lhdc, 8, 5, &HDDA876) : SetPixelV(lhdc, 9, 5, &HD8A573) : SetPixelV(lhdc, 10, 5, &HDFAE80) : SetPixelV(lhdc, 11, 5, &HDBAD7D) : SetPixelV(lhdc, 12, 5, &HDFB084) : SetPixelV(lhdc, 13, 5, &HDFB286) : SetPixelV(lhdc, 14, 5, &HDFB188) : SetPixelV(lhdc, 15, 5, &HE1B58D) : SetPixelV(lhdc, 16, 5, &HE3B58E) : SetPixelV(lhdc, 17, 5, &HE3B48E)
		SetPixelV(lhdc, 1, 6, &HF9F9F9) : SetPixelV(lhdc, 2, 6, &H7B706E) : SetPixelV(lhdc, 3, 6, &H871405) : SetPixelV(lhdc, 4, 6, &HA5330E) : SetPixelV(lhdc, 5, 6, &HB34C0D) : SetPixelV(lhdc, 6, 6, &HD27030) : SetPixelV(lhdc, 7, 6, &HD89C68) : SetPixelV(lhdc, 8, 6, &HDAA573) : SetPixelV(lhdc, 9, 6, &HD9A674) : SetPixelV(lhdc, 10, 6, &HD9A87A) : SetPixelV(lhdc, 11, 6, &HDBAD7D) : SetPixelV(lhdc, 12, 6, &HDBAC80) : SetPixelV(lhdc, 13, 6, &HDCAF83) : SetPixelV(lhdc, 14, 6, &HDFB188) : SetPixelV(lhdc, 15, 6, &HDEB28A) : SetPixelV(lhdc, 16, 6, &HDFB18A) : SetPixelV(lhdc, 17, 6, &HE0B18B)
		SetPixelV(lhdc, 1, 7, &HE8E8E7) : SetPixelV(lhdc, 2, 7, &H773F34) : SetPixelV(lhdc, 3, 7, &H9F2C00) : SetPixelV(lhdc, 4, 7, &HBA4B07) : SetPixelV(lhdc, 5, 7, &HC35E10) : SetPixelV(lhdc, 6, 7, &HCC7323) : SetPixelV(lhdc, 7, 7, &HDB8F46) : SetPixelV(lhdc, 8, 7, &HE8A763) : SetPixelV(lhdc, 9, 7, &HE3A76C) : SetPixelV(lhdc, 10, 7, &HE7AB70) : SetPixelV(lhdc, 11, 7, &HE8AE73) : SetPixelV(lhdc, 12, 7, &HE8AE73) : SetPixelV(lhdc, 13, 7, &HEDB17B) : SetPixelV(lhdc, 14, 7, &HEFB37D) : SetPixelV(lhdc, 15, 7, &HE9B57E) : SetPixelV(lhdc, 16, 7, &HE9B57E) : SetPixelV(lhdc, 17, 7, &HE9B47F)
		SetPixelV(lhdc, 0, 8, &HFDFDFD) : SetPixelV(lhdc, 1, 8, &HCAC5C5) : SetPixelV(lhdc, 2, 8, &H682A1F) : SetPixelV(lhdc, 3, 8, &HB23E0C) : SetPixelV(lhdc, 4, 8, &HCC5D19) : SetPixelV(lhdc, 5, 8, &HCE691B) : SetPixelV(lhdc, 6, 8, &HCE7525) : SetPixelV(lhdc, 7, 8, &HCD8138) : SetPixelV(lhdc, 8, 8, &HC58440) : SetPixelV(lhdc, 9, 8, &HC5894E) : SetPixelV(lhdc, 10, 8, &HC98D52) : SetPixelV(lhdc, 11, 8, &HC88E53) : SetPixelV(lhdc, 12, 8, &HCC9257) : SetPixelV(lhdc, 13, 8, &HCF935D) : SetPixelV(lhdc, 14, 8, &HD0945E) : SetPixelV(lhdc, 15, 8, &HCE9963) : SetPixelV(lhdc, 16, 8, &HCE9963) : SetPixelV(lhdc, 17, 8, &HCE9963)
		SetPixelV(lhdc, 0, 9, &HFAFAFA) : SetPixelV(lhdc, 1, 9, &HB9ADAB) : SetPixelV(lhdc, 2, 9, &H6E2B10) : SetPixelV(lhdc, 3, 9, &HB6580D) : SetPixelV(lhdc, 4, 9, &HCA6C20) : SetPixelV(lhdc, 5, 9, &HCE792B) : SetPixelV(lhdc, 6, 9, &HCE8132) : SetPixelV(lhdc, 7, 9, &HD08B42) : SetPixelV(lhdc, 8, 9, &HD3904B) : SetPixelV(lhdc, 9, 9, &HD3934C) : SetPixelV(lhdc, 10, 9, &HD89753) : SetPixelV(lhdc, 11, 9, &HDB9B5A) : SetPixelV(lhdc, 12, 9, &HDC9B5E) : SetPixelV(lhdc, 13, 9, &HDB9C60) : SetPixelV(lhdc, 14, 9, &HDB9C60) : SetPixelV(lhdc, 15, 9, &HDDA164) : SetPixelV(lhdc, 16, 9, &HDDA164) : SetPixelV(lhdc, 17, 9, &HDDA064)
		SetPixelV(lhdc, 0, 10, &HF7F7F7) : SetPixelV(lhdc, 1, 10, &HB0A09E) : SetPixelV(lhdc, 2, 10, &H712E13) : SetPixelV(lhdc, 3, 10, &HBD5F14) : SetPixelV(lhdc, 4, 10, &HD17327) : SetPixelV(lhdc, 5, 10, &HD47F31) : SetPixelV(lhdc, 6, 10, &HD98C3D) : SetPixelV(lhdc, 7, 10, &HD9944B) : SetPixelV(lhdc, 8, 10, &HD7944F) : SetPixelV(lhdc, 9, 10, &HDC9C55) : SetPixelV(lhdc, 10, 10, &HDC9B57) : SetPixelV(lhdc, 11, 10, &HE3A362) : SetPixelV(lhdc, 12, 10, &HE3A265) : SetPixelV(lhdc, 13, 10, &HE2A367) : SetPixelV(lhdc, 14, 10, &HE0A165) : SetPixelV(lhdc, 15, 10, &HE3A66A) : SetPixelV(lhdc, 16, 10, &HE3A66A) : SetPixelV(lhdc, 17, 10, &HE2A66A)
		tmph = lh - 22
		SetPixelV(lhdc, 0, tmph + 10, &HF7F7F7) : SetPixelV(lhdc, 1, tmph + 10, &HB0A09E) : SetPixelV(lhdc, 2, tmph + 10, &H712E13) : SetPixelV(lhdc, 3, tmph + 10, &HBD5F14) : SetPixelV(lhdc, 4, tmph + 10, &HD17327) : SetPixelV(lhdc, 5, tmph + 10, &HD47F31) : SetPixelV(lhdc, 6, tmph + 10, &HD98C3D) : SetPixelV(lhdc, 7, tmph + 10, &HD9944B) : SetPixelV(lhdc, 8, tmph + 10, &HD7944F) : SetPixelV(lhdc, 9, tmph + 10, &HDC9C55) : SetPixelV(lhdc, 10, tmph + 10, &HDC9B57) : SetPixelV(lhdc, 11, tmph + 10, &HE3A362) : SetPixelV(lhdc, 12, tmph + 10, &HE3A265) : SetPixelV(lhdc, 13, tmph + 10, &HE2A367) : SetPixelV(lhdc, 14, tmph + 10, &HE0A165) : SetPixelV(lhdc, 15, tmph + 10, &HE3A66A) : SetPixelV(lhdc, 16, tmph + 10, &HE3A66A) : SetPixelV(lhdc, 17, tmph + 10, &HE2A66A)
		SetPixelV(lhdc, 0, tmph + 11, &HF5F5F5) : SetPixelV(lhdc, 1, tmph + 11, &HACA39E) : SetPixelV(lhdc, 2, tmph + 11, &H744421) : SetPixelV(lhdc, 3, tmph + 11, &HC56F1F) : SetPixelV(lhdc, 4, tmph + 11, &HD17A2A) : SetPixelV(lhdc, 5, tmph + 11, &HD58C42) : SetPixelV(lhdc, 6, tmph + 11, &HD7914B) : SetPixelV(lhdc, 7, tmph + 11, &HDF9854) : SetPixelV(lhdc, 8, tmph + 11, &HE4A05F) : SetPixelV(lhdc, 9, tmph + 11, &HE29F66) : SetPixelV(lhdc, 10, tmph + 11, &HE4A56B) : SetPixelV(lhdc, 11, tmph + 11, &HDDA467) : SetPixelV(lhdc, 12, tmph + 11, &HE0A76A) : SetPixelV(lhdc, 13, tmph + 11, &HE2A96C) : SetPixelV(lhdc, 14, tmph + 11, &HE3A870) : SetPixelV(lhdc, 15, tmph + 11, &HE6AC76) : SetPixelV(lhdc, 16, tmph + 11, &HE6AC76) : SetPixelV(lhdc, 17, tmph + 11, &HE6AC76)
		SetPixelV(lhdc, 0, tmph + 12, &HF5F5F5) : SetPixelV(lhdc, 1, tmph + 12, &HB1AAA7) : SetPixelV(lhdc, 2, tmph + 12, &H825533) : SetPixelV(lhdc, 3, tmph + 12, &HCF792A) : SetPixelV(lhdc, 4, tmph + 12, &HE48D3D) : SetPixelV(lhdc, 5, tmph + 12, &HDD944A) : SetPixelV(lhdc, 6, tmph + 12, &HE49E58) : SetPixelV(lhdc, 7, tmph + 12, &HEBA460) : SetPixelV(lhdc, 8, tmph + 12, &HEEAA69) : SetPixelV(lhdc, 9, tmph + 12, &HF3B077) : SetPixelV(lhdc, 10, tmph + 12, &HEEAF75) : SetPixelV(lhdc, 11, tmph + 12, &HEBB275) : SetPixelV(lhdc, 12, tmph + 12, &HEFB679) : SetPixelV(lhdc, 13, tmph + 12, &HF1B87B) : SetPixelV(lhdc, 14, tmph + 12, &HF1B67E) : SetPixelV(lhdc, 15, tmph + 12, &HF2B781) : SetPixelV(lhdc, 16, tmph + 12, &HF1B681) : SetPixelV(lhdc, 17, tmph + 12, &HF1B681)
		SetPixelV(lhdc, 0, tmph + 13, &HF7F7F7) : SetPixelV(lhdc, 1, tmph + 13, &HC2C2C1) : SetPixelV(lhdc, 2, tmph + 13, &H6B5D4E) : SetPixelV(lhdc, 3, tmph + 13, &HC27831) : SetPixelV(lhdc, 4, tmph + 13, &HDA8E46) : SetPixelV(lhdc, 5, tmph + 13, &HE7A05C) : SetPixelV(lhdc, 6, tmph + 13, &HEAA665) : SetPixelV(lhdc, 7, tmph + 13, &HE9AF6E) : SetPixelV(lhdc, 8, tmph + 13, &HEFB377) : SetPixelV(lhdc, 9, tmph + 13, &HF3B579) : SetPixelV(lhdc, 10, tmph + 13, &HF7B97D) : SetPixelV(lhdc, 11, tmph + 13, &HF2BB7E) : SetPixelV(lhdc, 12, tmph + 13, &HF4BB83) : SetPixelV(lhdc, 13, tmph + 13, &HF5BE85) : SetPixelV(lhdc, 14, tmph + 13, &HF4BB87) : SetPixelV(lhdc, 15, tmph + 13, &HF5BE8A) : SetPixelV(lhdc, 16, tmph + 13, &HF5BD8A) : SetPixelV(lhdc, 17, tmph + 13, &HF3BD8A)
		SetPixelV(lhdc, 0, tmph + 14, &HFBFBFB) : SetPixelV(lhdc, 1, tmph + 14, &HE1E1E1) : SetPixelV(lhdc, 2, tmph + 14, &H85796E) : SetPixelV(lhdc, 3, tmph + 14, &HB76F2B) : SetPixelV(lhdc, 4, tmph + 14, &HDE924A) : SetPixelV(lhdc, 5, tmph + 14, &HE8A15D) : SetPixelV(lhdc, 6, tmph + 14, &HF2AE6D) : SetPixelV(lhdc, 7, tmph + 14, &HF1B776) : SetPixelV(lhdc, 8, tmph + 14, &HF2B67A) : SetPixelV(lhdc, 9, tmph + 14, &HFBBD81) : SetPixelV(lhdc, 10, tmph + 14, &HFFC286) : SetPixelV(lhdc, 11, tmph + 14, &HFAC386) : SetPixelV(lhdc, 12, tmph + 14, &HFBC28A) : SetPixelV(lhdc, 13, tmph + 14, &HFAC38A) : SetPixelV(lhdc, 14, tmph + 14, &HFAC18D) : SetPixelV(lhdc, 15, tmph + 14, &HFDC592) : SetPixelV(lhdc, 16, tmph + 14, &HFDC592) : SetPixelV(lhdc, 17, tmph + 14, &HFCC592)
		SetPixelV(lhdc, 0, tmph + 15, &HFEFEFE) : SetPixelV(lhdc, 1, tmph + 15, &HEDEDED) : SetPixelV(lhdc, 2, tmph + 15, &HA2A0A0) : SetPixelV(lhdc, 3, tmph + 15, &H816753) : SetPixelV(lhdc, 4, tmph + 15, &HC09068) : SetPixelV(lhdc, 5, tmph + 15, &HEDA55F) : SetPixelV(lhdc, 6, tmph + 15, &HFAB26C) : SetPixelV(lhdc, 7, tmph + 15, &HFCBF7D) : SetPixelV(lhdc, 8, tmph + 15, &HF7C182) : SetPixelV(lhdc, 9, tmph + 15, &HF8C38A) : SetPixelV(lhdc, 10, tmph + 15, &HFACA90) : SetPixelV(lhdc, 11, tmph + 15, &HF7CB8E) : SetPixelV(lhdc, 12, tmph + 15, &HF8CC8F) : SetPixelV(lhdc, 13, tmph + 15, &HFACC96) : SetPixelV(lhdc, 14, tmph + 15, &HF9CB95) : SetPixelV(lhdc, 15, tmph + 15, &HF9CE97) : SetPixelV(lhdc, 16, tmph + 15, &HF8CD97) : SetPixelV(lhdc, 17, tmph + 15, &HF8CE97)
		SetPixelV(lhdc, 1, tmph + 16, &HF6F6F6) : SetPixelV(lhdc, 2, tmph + 16, &HD6D6D6) : SetPixelV(lhdc, 3, tmph + 16, &H8E7C6F) : SetPixelV(lhdc, 4, tmph + 16, &H946843) : SetPixelV(lhdc, 5, tmph + 16, &HEEA762) : SetPixelV(lhdc, 6, tmph + 16, &HFFB771) : SetPixelV(lhdc, 7, tmph + 16, &HFEC17F) : SetPixelV(lhdc, 8, tmph + 16, &HFFC98A) : SetPixelV(lhdc, 9, tmph + 16, &HFFCE95) : SetPixelV(lhdc, 10, tmph + 16, &HFBCB91) : SetPixelV(lhdc, 11, tmph + 16, &HFFD396) : SetPixelV(lhdc, 12, tmph + 16, &HFFD396) : SetPixelV(lhdc, 13, tmph + 16, &HFFD29C) : SetPixelV(lhdc, 14, tmph + 16, &HFFD39D) : SetPixelV(lhdc, 15, tmph + 16, &HFFD49E) : SetPixelV(lhdc, 16, tmph + 16, &HFFD49E) : SetPixelV(lhdc, 17, tmph + 16, &HFED59E)
		SetPixelV(lhdc, 1, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 2, tmph + 17, &HEDEDED) : SetPixelV(lhdc, 3, tmph + 17, &HBEBEBE) : SetPixelV(lhdc, 4, tmph + 17, &H6C6C6C) : SetPixelV(lhdc, 5, tmph + 17, &H7C684F) : SetPixelV(lhdc, 6, tmph + 17, &HD1AE81) : SetPixelV(lhdc, 7, tmph + 17, &HF1C284) : SetPixelV(lhdc, 8, tmph + 17, &HFDCE90) : SetPixelV(lhdc, 9, tmph + 17, &HF8D193) : SetPixelV(lhdc, 10, tmph + 17, &HFBD899) : SetPixelV(lhdc, 11, tmph + 17, &HF5DC9E) : SetPixelV(lhdc, 12, tmph + 17, &HF8DFA1) : SetPixelV(lhdc, 13, tmph + 17, &HF8DFA1) : SetPixelV(lhdc, 14, tmph + 17, &HF8DFA1) : SetPixelV(lhdc, 15, tmph + 17, &HF8DEA3) : SetPixelV(lhdc, 16, tmph + 17, &HF7DDA3) : SetPixelV(lhdc, 17, tmph + 17, &HF7DDA3)
		SetPixelV(lhdc, 2, tmph + 18, &HF9F9F9) : SetPixelV(lhdc, 3, tmph + 18, &HE6E6E6) : SetPixelV(lhdc, 4, tmph + 18, &HBABABA) : SetPixelV(lhdc, 5, tmph + 18, &H827666) : SetPixelV(lhdc, 6, tmph + 18, &H836743) : SetPixelV(lhdc, 7, tmph + 18, &HBE935B) : SetPixelV(lhdc, 8, tmph + 18, &HF4C78B) : SetPixelV(lhdc, 9, tmph + 18, &HFDD79A) : SetPixelV(lhdc, 10, tmph + 18, &HFFDFA0) : SetPixelV(lhdc, 11, tmph + 18, &HFBE2A4) : SetPixelV(lhdc, 12, tmph + 18, &HFFE7A9) : SetPixelV(lhdc, 13, tmph + 18, &HFFE9AB) : SetPixelV(lhdc, 14, tmph + 18, &HFFE7A9) : SetPixelV(lhdc, 15, tmph + 18, &HFFE6AC) : SetPixelV(lhdc, 16, tmph + 18, &HFFE6AD) : SetPixelV(lhdc, 17, tmph + 18, &HFFE6AD)
		SetPixelV(lhdc, 2, tmph + 19, &HFEFEFE) : SetPixelV(lhdc, 3, tmph + 19, &HF8F8F8) : SetPixelV(lhdc, 4, tmph + 19, &HE6E6E6) : SetPixelV(lhdc, 5, tmph + 19, &HC8C8C8) : SetPixelV(lhdc, 6, tmph + 19, &H8F8F8F) : SetPixelV(lhdc, 7, tmph + 19, &H686462) : SetPixelV(lhdc, 8, tmph + 19, &H6D655E) : SetPixelV(lhdc, 9, tmph + 19, &H918472) : SetPixelV(lhdc, 10, tmph + 19, &HB3A88E) : SetPixelV(lhdc, 11, tmph + 19, &HDAD1B2) : SetPixelV(lhdc, 12, tmph + 19, &HE3DBBA) : SetPixelV(lhdc, 13, tmph + 19, &HE7E0C0) : SetPixelV(lhdc, 14, tmph + 19, &HE9E2C1) : SetPixelV(lhdc, 15, tmph + 19, &HE9E2C5) : SetPixelV(lhdc, 16, tmph + 19, &HE9E1C5) : SetPixelV(lhdc, 17, tmph + 19, &HE9E2C5)
		SetPixelV(lhdc, 3, tmph + 20, &HFEFEFE) : SetPixelV(lhdc, 4, tmph + 20, &HF9F9F9) : SetPixelV(lhdc, 5, tmph + 20, &HECECEC) : SetPixelV(lhdc, 6, tmph + 20, &HDADADA) : SetPixelV(lhdc, 7, tmph + 20, &HC2C2C1) : SetPixelV(lhdc, 8, tmph + 20, &H9F9D9B) : SetPixelV(lhdc, 9, tmph + 20, &H827D75) : SetPixelV(lhdc, 10, tmph + 20, &H6A6353) : SetPixelV(lhdc, 11, tmph + 20, &H5F5941) : SetPixelV(lhdc, 12, tmph + 20, &H5D553B) : SetPixelV(lhdc, 13, tmph + 20, &H595338) : SetPixelV(lhdc, 14, tmph + 20, &H5E5739) : SetPixelV(lhdc, 15, tmph + 20, &H5F5A3C) : SetPixelV(lhdc, 16, tmph + 20, &H635E3F) : SetPixelV(lhdc, 17, tmph + 20, &H635D40)
		SetPixelV(lhdc, 5, tmph + 21, &HFCFCFC) : SetPixelV(lhdc, 6, tmph + 21, &HF5F5F5) : SetPixelV(lhdc, 7, tmph + 21, &HEBEBEB) : SetPixelV(lhdc, 8, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, 9, tmph + 21, &HD6D6D6) : SetPixelV(lhdc, 10, tmph + 21, &HCECECE) : SetPixelV(lhdc, 11, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, 12, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 13, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 14, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 15, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 16, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, 17, tmph + 21, &HC5C5C5)
		SetPixelV(lhdc, 7, tmph + 22, &HFDFDFD) : SetPixelV(lhdc, 8, tmph + 22, &HF9F9F9) : SetPixelV(lhdc, 9, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, 10, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, 11, tmph + 22, &HEEEEEE) : SetPixelV(lhdc, 12, tmph + 22, &HEDEDED) : SetPixelV(lhdc, 13, tmph + 22, &HECECEC) : SetPixelV(lhdc, 14, tmph + 22, &HECECEC) : SetPixelV(lhdc, 15, tmph + 22, &HECECEC) : SetPixelV(lhdc, 16, tmph + 22, &HECECEC) : SetPixelV(lhdc, 17, tmph + 22, &HECECEC)
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, 0, &H450608) : SetPixelV(lhdc, tmpw + 18, 0, &H450608) : SetPixelV(lhdc, tmpw + 19, 0, &H3B0707) : SetPixelV(lhdc, tmpw + 20, 0, &H370706) : SetPixelV(lhdc, tmpw + 21, 0, &H360507) : SetPixelV(lhdc, tmpw + 22, 0, &H3B0F10) : SetPixelV(lhdc, tmpw + 23, 0, &H442526) : SetPixelV(lhdc, tmpw + 24, 0, &H604E4E) : SetPixelV(lhdc, tmpw + 25, 0, &HA29D9E) : SetPixelV(lhdc, tmpw + 26, 0, &HEEEEEE) : SetPixelV(lhdc, tmpw + 34, 0, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 1, &HF1D4C9) : SetPixelV(lhdc, tmpw + 18, 1, &HF1D4C9) : SetPixelV(lhdc, tmpw + 19, 1, &HEDD3CD) : SetPixelV(lhdc, tmpw + 20, 1, &HEBD1CB) : SetPixelV(lhdc, tmpw + 21, 1, &HE9CEC4) : SetPixelV(lhdc, tmpw + 22, 1, &HE5C1B9) : SetPixelV(lhdc, tmpw + 23, 1, &HCFA89F) : SetPixelV(lhdc, tmpw + 24, 1, &HAA6E68) : SetPixelV(lhdc, tmpw + 25, 1, &H73211B) : SetPixelV(lhdc, tmpw + 26, 1, &H702924) : SetPixelV(lhdc, tmpw + 27, 1, &HAA9897) : SetPixelV(lhdc, tmpw + 28, 1, &HF7F7F7) : SetPixelV(lhdc, tmpw + 34, 1, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 2, &HE5C8BD) : SetPixelV(lhdc, tmpw + 18, 2, &HE5C8BD) : SetPixelV(lhdc, tmpw + 19, 2, &HDEC4BE) : SetPixelV(lhdc, tmpw + 20, 2, &HDCC2BC) : SetPixelV(lhdc, tmpw + 21, 2, &HE2C7BD) : SetPixelV(lhdc, tmpw + 22, 2, &HE2BEB6) : SetPixelV(lhdc, tmpw + 23, 2, &HE8C1B8) : SetPixelV(lhdc, tmpw + 24, 2, &HF0B4AE) : SetPixelV(lhdc, tmpw + 25, 2, &HF29C96) : SetPixelV(lhdc, tmpw + 26, 2, &H822D27) : SetPixelV(lhdc, tmpw + 27, 2, &H400807) : SetPixelV(lhdc, tmpw + 28, 2, &H71585A) : SetPixelV(lhdc, tmpw + 29, 2, &HEBEBEB) : SetPixelV(lhdc, tmpw + 34, 2, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 3, &HE8C0A1) : SetPixelV(lhdc, tmpw + 18, 3, &HE8C0A1) : SetPixelV(lhdc, tmpw + 19, 3, &HE5C09A) : SetPixelV(lhdc, tmpw + 20, 3, &HE4BF99) : SetPixelV(lhdc, tmpw + 21, 3, &HE4BA97) : SetPixelV(lhdc, tmpw + 22, 3, &HE9BF9C) : SetPixelV(lhdc, tmpw + 23, 3, &HDFB695) : SetPixelV(lhdc, tmpw + 24, 3, &HDFB695) : SetPixelV(lhdc, tmpw + 25, 3, &HE0AE90) : SetPixelV(lhdc, tmpw + 26, 3, &HCB8469) : SetPixelV(lhdc, tmpw + 27, 3, &H941600) : SetPixelV(lhdc, tmpw + 28, 3, &H830800) : SetPixelV(lhdc, tmpw + 29, 3, &H895253) : SetPixelV(lhdc, tmpw + 30, 3, &HF0EFEF) : SetPixelV(lhdc, tmpw + 34, 3, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 4, &HE0B898) : SetPixelV(lhdc, tmpw + 18, 4, &HE0B897) : SetPixelV(lhdc, tmpw + 19, 4, &HDAB58F) : SetPixelV(lhdc, tmpw + 20, 4, &HDBB690) : SetPixelV(lhdc, tmpw + 21, 4, &HDBB18E) : SetPixelV(lhdc, tmpw + 22, 4, &HD7AD8A) : SetPixelV(lhdc, tmpw + 23, 4, &HDAB190) : SetPixelV(lhdc, tmpw + 24, 4, &HD2A988) : SetPixelV(lhdc, tmpw + 25, 4, &HD6A486) : SetPixelV(lhdc, tmpw + 26, 4, &HDA9378) : SetPixelV(lhdc, tmpw + 27, 4, &HBF4129) : SetPixelV(lhdc, tmpw + 28, 4, &H991B03) : SetPixelV(lhdc, tmpw + 29, 4, &H500709) : SetPixelV(lhdc, tmpw + 30, 4, &H826F70) : SetPixelV(lhdc, tmpw + 31, 4, &HFAFAFA) : SetPixelV(lhdc, tmpw + 34, 4, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 5, &HE3B48E) : SetPixelV(lhdc, tmpw + 18, 5, &HE3B48D) : SetPixelV(lhdc, tmpw + 19, 5, &HE0B387) : SetPixelV(lhdc, tmpw + 20, 5, &HDEB185) : SetPixelV(lhdc, tmpw + 21, 5, &HE1B084) : SetPixelV(lhdc, tmpw + 22, 5, &HE3AE83) : SetPixelV(lhdc, tmpw + 23, 5, &HE1AF7B) : SetPixelV(lhdc, tmpw + 24, 5, &HE0A976) : SetPixelV(lhdc, tmpw + 25, 5, &HDCA473) : SetPixelV(lhdc, tmpw + 26, 5, &HDEA372) : SetPixelV(lhdc, tmpw + 27, 5, &HCC712E) : SetPixelV(lhdc, tmpw + 28, 5, &HA53900) : SetPixelV(lhdc, tmpw + 29, 5, &H9D2200) : SetPixelV(lhdc, tmpw + 30, 5, &H9E2114) : SetPixelV(lhdc, tmpw + 31, 5, &HC7C5C4) : SetPixelV(lhdc, tmpw + 32, 5, &HFEFEFE) : SetPixelV(lhdc, tmpw + 34, 5, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 6, &HE0B18B) : SetPixelV(lhdc, tmpw + 18, 6, &HE0B18A) : SetPixelV(lhdc, tmpw + 19, 6, &HDEB185) : SetPixelV(lhdc, tmpw + 20, 6, &HDEB185) : SetPixelV(lhdc, tmpw + 21, 6, &HDCAB7F) : SetPixelV(lhdc, tmpw + 22, 6, &HE1AC81) : SetPixelV(lhdc, tmpw + 23, 6, &HDCAA76) : SetPixelV(lhdc, tmpw + 24, 6, &HDCA572) : SetPixelV(lhdc, tmpw + 25, 6, &HDBA372) : SetPixelV(lhdc, tmpw + 26, 6, &HD79C6B) : SetPixelV(lhdc, tmpw + 27, 6, &HD17633) : SetPixelV(lhdc, tmpw + 28, 6, &HB74B0B) : SetPixelV(lhdc, tmpw + 29, 6, &HAC310D) : SetPixelV(lhdc, tmpw + 30, 6, &H961507) : SetPixelV(lhdc, tmpw + 31, 6, &H736D6A) : SetPixelV(lhdc, tmpw + 32, 6, &HF8F8F8) : SetPixelV(lhdc, tmpw + 34, 6, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 7, &HE9B47F) : SetPixelV(lhdc, tmpw + 18, 7, &HEAB47E) : SetPixelV(lhdc, tmpw + 19, 7, &HEFB67E) : SetPixelV(lhdc, tmpw + 20, 7, &HE8AF77) : SetPixelV(lhdc, tmpw + 21, 7, &HE7AF74) : SetPixelV(lhdc, tmpw + 22, 7, &HE4AC71) : SetPixelV(lhdc, tmpw + 23, 7, &HEAAD6F) : SetPixelV(lhdc, tmpw + 24, 7, &HE9A968) : SetPixelV(lhdc, tmpw + 25, 7, &HE7A564) : SetPixelV(lhdc, tmpw + 26, 7, &HD9904C) : SetPixelV(lhdc, tmpw + 27, 7, &HC5711F) : SetPixelV(lhdc, tmpw + 28, 7, &HC16010) : SetPixelV(lhdc, tmpw + 29, 7, &HBB4D05) : SetPixelV(lhdc, tmpw + 30, 7, &HA02D00) : SetPixelV(lhdc, tmpw + 31, 7, &H774033) : SetPixelV(lhdc, tmpw + 32, 7, &HE7E6E6) : SetPixelV(lhdc, tmpw + 34, 7, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 8, &HCE9963) : SetPixelV(lhdc, tmpw + 18, 8, &HCF9963) : SetPixelV(lhdc, tmpw + 19, 8, &HCE955D) : SetPixelV(lhdc, tmpw + 20, 8, &HCE955D) : SetPixelV(lhdc, tmpw + 21, 8, &HCA9257) : SetPixelV(lhdc, tmpw + 22, 8, &HC89055) : SetPixelV(lhdc, tmpw + 23, 8, &HCB8E50) : SetPixelV(lhdc, tmpw + 24, 8, &HCB8B4A) : SetPixelV(lhdc, tmpw + 25, 8, &HC58342) : SetPixelV(lhdc, tmpw + 26, 8, &HC87F3B) : SetPixelV(lhdc, tmpw + 27, 8, &HCA7624) : SetPixelV(lhdc, tmpw + 28, 8, &HCA6919) : SetPixelV(lhdc, tmpw + 29, 8, &HCC5E16) : SetPixelV(lhdc, tmpw + 30, 8, &HB23E07) : SetPixelV(lhdc, tmpw + 31, 8, &H682B1D) : SetPixelV(lhdc, tmpw + 32, 8, &HC7C2C2) : SetPixelV(lhdc, tmpw + 33, 8, &HFDFDFD) : SetPixelV(lhdc, tmpw + 34, 8, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 9, &HDDA064) : SetPixelV(lhdc, tmpw + 18, 9, &HDCA064) : SetPixelV(lhdc, tmpw + 19, 9, &HDA9D5D) : SetPixelV(lhdc, tmpw + 20, 9, &HD99C5C) : SetPixelV(lhdc, tmpw + 21, 9, &HDA9D5D) : SetPixelV(lhdc, tmpw + 22, 9, &HDA9A5A) : SetPixelV(lhdc, tmpw + 23, 9, &HD89753) : SetPixelV(lhdc, tmpw + 24, 9, &HD7914E) : SetPixelV(lhdc, tmpw + 25, 9, &HD38E49) : SetPixelV(lhdc, tmpw + 26, 9, &HD38B43) : SetPixelV(lhdc, tmpw + 27, 9, &HCD8430) : SetPixelV(lhdc, tmpw + 28, 9, &HCA7826) : SetPixelV(lhdc, tmpw + 29, 9, &HCE6C1E) : SetPixelV(lhdc, tmpw + 30, 9, &HB9560C) : SetPixelV(lhdc, tmpw + 31, 9, &H742E0D) : SetPixelV(lhdc, tmpw + 32, 9, &HB3A6A4) : SetPixelV(lhdc, tmpw + 33, 9, &HF9F9F9) : SetPixelV(lhdc, tmpw + 34, 9, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 10, &HE2A66A) : SetPixelV(lhdc, tmpw + 18, 10, &HE2A66A) : SetPixelV(lhdc, tmpw + 19, 10, &HE1A464) : SetPixelV(lhdc, tmpw + 20, 10, &HE0A363) : SetPixelV(lhdc, tmpw + 21, 10, &HE0A363) : SetPixelV(lhdc, tmpw + 22, 10, &HE1A161) : SetPixelV(lhdc, tmpw + 23, 10, &HE09F5B) : SetPixelV(lhdc, tmpw + 24, 10, &HDE9855) : SetPixelV(lhdc, tmpw + 25, 10, &HDC9752) : SetPixelV(lhdc, tmpw + 26, 10, &HDB934B) : SetPixelV(lhdc, tmpw + 27, 10, &HD68D39) : SetPixelV(lhdc, tmpw + 28, 10, &HD17F2D) : SetPixelV(lhdc, tmpw + 29, 10, &HD67426) : SetPixelV(lhdc, tmpw + 30, 10, &HC05D13) : SetPixelV(lhdc, tmpw + 31, 10, &H7C3514) : SetPixelV(lhdc, tmpw + 32, 10, &HAB9B98) : SetPixelV(lhdc, tmpw + 33, 10, &HF6F6F6) : SetPixelV(lhdc, tmpw + 34, 10, &HFFFFFFFF)
		tmph = lh - 22
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, tmph + 10, &HE2A66A) : SetPixelV(lhdc, tmpw + 18, tmph + 10, &HE2A66A) : SetPixelV(lhdc, tmpw + 19, tmph + 10, &HE1A464) : SetPixelV(lhdc, tmpw + 20, tmph + 10, &HE0A363) : SetPixelV(lhdc, tmpw + 21, tmph + 10, &HE0A363) : SetPixelV(lhdc, tmpw + 22, tmph + 10, &HE1A161) : SetPixelV(lhdc, tmpw + 23, tmph + 10, &HE09F5B) : SetPixelV(lhdc, tmpw + 24, tmph + 10, &HDE9855) : SetPixelV(lhdc, tmpw + 25, tmph + 10, &HDC9752) : SetPixelV(lhdc, tmpw + 26, tmph + 10, &HDB934B) : SetPixelV(lhdc, tmpw + 27, tmph + 10, &HD68D39) : SetPixelV(lhdc, tmpw + 28, tmph + 10, &HD17F2D) : SetPixelV(lhdc, tmpw + 29, tmph + 10, &HD67426) : SetPixelV(lhdc, tmpw + 30, tmph + 10, &HC05D13) : SetPixelV(lhdc, tmpw + 31, tmph + 10, &H7C3514) : SetPixelV(lhdc, tmpw + 32, tmph + 10, &HAB9B98) : SetPixelV(lhdc, tmpw + 33, tmph + 10, &HF6F6F6)
		SetPixelV(lhdc, tmpw + 17, tmph + 11, &HE6AC76) : SetPixelV(lhdc, tmpw + 18, tmph + 11, &HE6AC76) : SetPixelV(lhdc, tmpw + 19, tmph + 11, &HE2A86D) : SetPixelV(lhdc, tmpw + 20, tmph + 11, &HE5A66C) : SetPixelV(lhdc, tmpw + 21, tmph + 11, &HE1A56A) : SetPixelV(lhdc, tmpw + 22, tmph + 11, &HE4A46A) : SetPixelV(lhdc, tmpw + 23, tmph + 11, &HE1A266) : SetPixelV(lhdc, tmpw + 24, tmph + 11, &HE6A364) : SetPixelV(lhdc, tmpw + 25, tmph + 11, &HE19F5E) : SetPixelV(lhdc, tmpw + 26, tmph + 11, &HDF9A55) : SetPixelV(lhdc, tmpw + 27, tmph + 11, &HD89048) : SetPixelV(lhdc, tmpw + 28, tmph + 11, &HD88A3E) : SetPixelV(lhdc, tmpw + 29, tmph + 11, &HCF7927) : SetPixelV(lhdc, tmpw + 30, tmph + 11, &HC87220) : SetPixelV(lhdc, tmpw + 31, tmph + 11, &H77481E) : SetPixelV(lhdc, tmpw + 32, tmph + 11, &HABA39E) : SetPixelV(lhdc, tmpw + 33, tmph + 11, &HF4F4F4)
		SetPixelV(lhdc, tmpw + 17, tmph + 12, &HF1B681) : SetPixelV(lhdc, tmpw + 18, tmph + 12, &HF0B780) : SetPixelV(lhdc, tmpw + 19, tmph + 12, &HF2B87D) : SetPixelV(lhdc, tmpw + 20, tmph + 12, &HF5B67C) : SetPixelV(lhdc, tmpw + 21, tmph + 12, &HF1B57A) : SetPixelV(lhdc, tmpw + 22, tmph + 12, &HF2B278) : SetPixelV(lhdc, tmpw + 23, tmph + 12, &HF0B175) : SetPixelV(lhdc, tmpw + 24, tmph + 12, &HF3B071) : SetPixelV(lhdc, tmpw + 25, tmph + 12, &HECAA69) : SetPixelV(lhdc, tmpw + 26, tmph + 12, &HE9A45F) : SetPixelV(lhdc, tmpw + 27, tmph + 12, &HE8A058) : SetPixelV(lhdc, tmpw + 28, tmph + 12, &HE5974B) : SetPixelV(lhdc, tmpw + 29, tmph + 12, &HE38D3B) : SetPixelV(lhdc, tmpw + 30, tmph + 12, &HD37D2B) : SetPixelV(lhdc, tmpw + 31, tmph + 12, &H895A32) : SetPixelV(lhdc, tmpw + 32, tmph + 12, &HB4AFAC) : SetPixelV(lhdc, tmpw + 33, tmph + 12, &HF5F5F5)
		SetPixelV(lhdc, tmpw + 17, tmph + 13, &HF3BD8A) : SetPixelV(lhdc, tmpw + 18, tmph + 13, &HF3BD8A) : SetPixelV(lhdc, tmpw + 19, tmph + 13, &HF2BD84) : SetPixelV(lhdc, tmpw + 20, tmph + 13, &HF5BC84) : SetPixelV(lhdc, tmpw + 21, tmph + 13, &HF3BC83) : SetPixelV(lhdc, tmpw + 22, tmph + 13, &HF4B981) : SetPixelV(lhdc, tmpw + 23, tmph + 13, &HF2B97C) : SetPixelV(lhdc, tmpw + 24, tmph + 13, &HF5B77B) : SetPixelV(lhdc, tmpw + 25, tmph + 13, &HF1B476) : SetPixelV(lhdc, tmpw + 26, tmph + 13, &HEFAF6E) : SetPixelV(lhdc, tmpw + 27, tmph + 13, &HE5A45F) : SetPixelV(lhdc, tmpw + 28, tmph + 13, &HE49F5A) : SetPixelV(lhdc, tmpw + 29, tmph + 13, &HDA8F4A) : SetPixelV(lhdc, tmpw + 30, tmph + 13, &HC57A35) : SetPixelV(lhdc, tmpw + 31, tmph + 13, &H736353) : SetPixelV(lhdc, tmpw + 32, tmph + 13, &HD6D5D5) : SetPixelV(lhdc, tmpw + 33, tmph + 13, &HF8F8F8)
		SetPixelV(lhdc, tmpw + 17, tmph + 14, &HFCC592) : SetPixelV(lhdc, tmpw + 18, tmph + 14, &HFBC592) : SetPixelV(lhdc, tmpw + 19, tmph + 14, &HF7C289) : SetPixelV(lhdc, tmpw + 20, tmph + 14, &HFCC38B) : SetPixelV(lhdc, tmpw + 21, tmph + 14, &HFAC38A) : SetPixelV(lhdc, tmpw + 22, tmph + 14, &HFDC28A) : SetPixelV(lhdc, tmpw + 23, tmph + 14, &HFBC285) : SetPixelV(lhdc, tmpw + 24, tmph + 14, &HFBBD81) : SetPixelV(lhdc, tmpw + 25, tmph + 14, &HF6B97B) : SetPixelV(lhdc, tmpw + 26, tmph + 14, &HF6B675) : SetPixelV(lhdc, tmpw + 27, tmph + 14, &HF0AF6A) : SetPixelV(lhdc, tmpw + 28, tmph + 14, &HE8A35E) : SetPixelV(lhdc, tmpw + 29, tmph + 14, &HDD924D) : SetPixelV(lhdc, tmpw + 30, tmph + 14, &HBA702B) : SetPixelV(lhdc, tmpw + 31, tmph + 14, &H847A70) : SetPixelV(lhdc, tmpw + 32, tmph + 14, &HE8E8E8) : SetPixelV(lhdc, tmpw + 33, tmph + 14, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 15, &HF8CE97) : SetPixelV(lhdc, tmpw + 18, tmph + 15, &HF9CD97) : SetPixelV(lhdc, tmpw + 19, tmph + 15, &HF9CE95) : SetPixelV(lhdc, tmpw + 20, tmph + 15, &HF7CC93) : SetPixelV(lhdc, tmpw + 21, tmph + 15, &HF6CB92) : SetPixelV(lhdc, tmpw + 22, tmph + 15, &HF9CA92) : SetPixelV(lhdc, tmpw + 23, tmph + 15, &HFCCD90) : SetPixelV(lhdc, tmpw + 24, tmph + 15, &HF8C488) : SetPixelV(lhdc, tmpw + 25, tmph + 15, &HF3BD80) : SetPixelV(lhdc, tmpw + 26, tmph + 15, &HFABD7D) : SetPixelV(lhdc, tmpw + 27, tmph + 15, &HF7B26D) : SetPixelV(lhdc, tmpw + 28, tmph + 15, &HEAA560) : SetPixelV(lhdc, tmpw + 29, tmph + 15, &HC0925D) : SetPixelV(lhdc, tmpw + 30, tmph + 15, &H896F54) : SetPixelV(lhdc, tmpw + 31, tmph + 15, &HBABABB) : SetPixelV(lhdc, tmpw + 32, tmph + 15, &HF1F1F1)
		SetPixelV(lhdc, tmpw + 17, tmph + 16, &HFED59E) : SetPixelV(lhdc, tmpw + 18, tmph + 16, &HFFD59F) : SetPixelV(lhdc, tmpw + 19, tmph + 16, &HFED39A) : SetPixelV(lhdc, tmpw + 20, tmph + 16, &HFFD49B) : SetPixelV(lhdc, tmpw + 21, tmph + 16, &HFCD198) : SetPixelV(lhdc, tmpw + 22, tmph + 16, &HFFD098) : SetPixelV(lhdc, tmpw + 23, tmph + 16, &HFECF92) : SetPixelV(lhdc, tmpw + 24, tmph + 16, &HFFCB8F) : SetPixelV(lhdc, tmpw + 25, tmph + 16, &HFFC98C) : SetPixelV(lhdc, tmpw + 26, tmph + 16, &HFEC181) : SetPixelV(lhdc, tmpw + 27, tmph + 16, &HFBB671) : SetPixelV(lhdc, tmpw + 28, tmph + 16, &HF0AB66) : SetPixelV(lhdc, tmpw + 29, tmph + 16, &H9F733E) : SetPixelV(lhdc, tmpw + 30, tmph + 16, &H918478) : SetPixelV(lhdc, tmpw + 31, tmph + 16, &HE2E2E2) : SetPixelV(lhdc, tmpw + 32, tmph + 16, &HF9F9F9)
		SetPixelV(lhdc, tmpw + 17, tmph + 17, &HF7DDA3) : SetPixelV(lhdc, tmpw + 18, tmph + 17, &HF8DDA4) : SetPixelV(lhdc, tmpw + 19, tmph + 17, &HF9E0A2) : SetPixelV(lhdc, tmpw + 20, tmph + 17, &HF5DC9E) : SetPixelV(lhdc, tmpw + 21, tmph + 17, &HF8DEA2) : SetPixelV(lhdc, tmpw + 22, tmph + 17, &HFBDDA2) : SetPixelV(lhdc, tmpw + 23, tmph + 17, &HF7D495) : SetPixelV(lhdc, tmpw + 24, tmph + 17, &HF8D193) : SetPixelV(lhdc, tmpw + 25, tmph + 17, &HFCCD90) : SetPixelV(lhdc, tmpw + 26, tmph + 17, &HF1C088) : SetPixelV(lhdc, tmpw + 27, tmph + 17, &HDBB186) : SetPixelV(lhdc, tmpw + 28, tmph + 17, &H8C7259) : SetPixelV(lhdc, tmpw + 29, tmph + 17, &H6D6B6B) : SetPixelV(lhdc, tmpw + 30, tmph + 17, &HD2D2D2) : SetPixelV(lhdc, tmpw + 31, tmph + 17, &HF2F2F2) : SetPixelV(lhdc, tmpw + 32, tmph + 17, &HFEFEFE)
		SetPixelV(lhdc, tmpw + 17, tmph + 18, &HFFE6AD) : SetPixelV(lhdc, tmpw + 18, tmph + 18, &HFFE6AD) : SetPixelV(lhdc, tmpw + 19, tmph + 18, &HFFE7A9) : SetPixelV(lhdc, tmpw + 20, tmph + 18, &HFFEAAC) : SetPixelV(lhdc, tmpw + 21, tmph + 18, &HF7DDA1) : SetPixelV(lhdc, tmpw + 22, tmph + 18, &HFFE1A6) : SetPixelV(lhdc, tmpw + 23, tmph + 18, &HFFE1A2) : SetPixelV(lhdc, tmpw + 24, tmph + 18, &HFED799) : SetPixelV(lhdc, tmpw + 25, tmph + 18, &HFACC8F) : SetPixelV(lhdc, tmpw + 26, tmph + 18, &HC99A64) : SetPixelV(lhdc, tmpw + 27, tmph + 18, &H977048) : SetPixelV(lhdc, tmpw + 28, tmph + 18, &H817060) : SetPixelV(lhdc, tmpw + 29, tmph + 18, &HC8C8C8) : SetPixelV(lhdc, tmpw + 30, tmph + 18, &HEBEBEB) : SetPixelV(lhdc, tmpw + 31, tmph + 18, &HFCFCFC)
		SetPixelV(lhdc, tmpw + 17, tmph + 19, &HE9E2C5) : SetPixelV(lhdc, tmpw + 18, tmph + 19, &HE9E2C5) : SetPixelV(lhdc, tmpw + 19, tmph + 19, &HEAE2C4) : SetPixelV(lhdc, tmpw + 20, tmph + 19, &HE7DFC1) : SetPixelV(lhdc, tmpw + 21, tmph + 19, &HEEE4C6) : SetPixelV(lhdc, tmpw + 22, tmph + 19, &HDBD1B4) : SetPixelV(lhdc, tmpw + 23, tmph + 19, &HB7AF93) : SetPixelV(lhdc, tmpw + 24, tmph + 19, &H8D8973) : SetPixelV(lhdc, tmpw + 25, tmph + 19, &H736D60) : SetPixelV(lhdc, tmpw + 26, tmph + 19, &H6A6660) : SetPixelV(lhdc, tmpw + 27, tmph + 19, &H8E9090) : SetPixelV(lhdc, tmpw + 28, tmph + 19, &HCDCDCD) : SetPixelV(lhdc, tmpw + 29, tmph + 19, &HE8E8E8) : SetPixelV(lhdc, tmpw + 30, tmph + 19, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 20, &H635D40) : SetPixelV(lhdc, tmpw + 18, tmph + 20, &H615B3F) : SetPixelV(lhdc, tmpw + 19, tmph + 20, &H60583C) : SetPixelV(lhdc, tmpw + 20, tmph + 20, &H5D563A) : SetPixelV(lhdc, tmpw + 21, tmph + 20, &H61583D) : SetPixelV(lhdc, tmpw + 22, tmph + 20, &H605840) : SetPixelV(lhdc, tmpw + 23, tmph + 20, &H6A6556) : SetPixelV(lhdc, tmpw + 24, tmph + 20, &H7F7D75) : SetPixelV(lhdc, tmpw + 25, tmph + 20, &HA4A3A1) : SetPixelV(lhdc, tmpw + 26, tmph + 20, &HC5C5C5) : SetPixelV(lhdc, tmpw + 27, tmph + 20, &HDADADA) : SetPixelV(lhdc, tmpw + 28, tmph + 20, &HEDEDED) : SetPixelV(lhdc, tmpw + 29, tmph + 20, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 18, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 19, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 20, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 21, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 22, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, tmpw + 23, tmph + 21, &HCECECE) : SetPixelV(lhdc, tmpw + 24, tmph + 21, &HD7D7D7) : SetPixelV(lhdc, tmpw + 25, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, tmpw + 26, tmph + 21, &HECECEC) : SetPixelV(lhdc, tmpw + 27, tmph + 21, &HF6F6F6) : SetPixelV(lhdc, tmpw + 28, tmph + 21, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 18, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 19, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 20, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 21, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 22, tmph + 22, &HEDEDED) : SetPixelV(lhdc, tmpw + 23, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, tmpw + 24, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, tmpw + 25, tmph + 22, &HFAFAFA) : SetPixelV(lhdc, tmpw + 26, tmph + 22, &HFDFDFD)
		tmph = 11 : tmph1 = lh - 10 : tmpw = lw - 34
		'Generar lineas intermedias
		APILine(0, tmph, 0, tmph1, &HF7F7F7) : APILine(1, tmph, 1, tmph1, &HB0A09E) : APILine(2, tmph, 2, tmph1, &H712E13) : APILine(3, tmph, 3, tmph1, &HBD5F14)
		APILine(4, tmph, 4, tmph1, &HD17327) : APILine(5, tmph, 5, tmph1, &HD47F31) : APILine(6, tmph, 6, tmph1, &HD98C3D) : APILine(7, tmph, 7, tmph1, &HD9944B)
		APILine(8, tmph, 8, tmph1, &HD7944F) : APILine(9, tmph, 9, tmph1, &HDC9C55) : APILine(10, tmph, 10, tmph1, &HDC9B57) : APILine(11, tmph, 11, tmph1, &HE3A362)
		APILine(12, tmph, 12, tmph1, &HE3A265) : APILine(13, tmph, 13, tmph1, &HE2A367) : APILine(14, tmph, 14, tmph1, &HE0A165) : APILine(15, tmph, 15, tmph1, &HE3A66A)
		APILine(16, tmph, 16, tmph1, &HE3A66A) : APILine(17, tmph, 17, tmph1, &HE2A66A)
		APILine(tmpw + 17, tmph, tmpw + 17, tmph1, &HE2A66A) : APILine(tmpw + 18, tmph, tmpw + 18, tmph1, &HE2A66A) : APILine(tmpw + 19, tmph, tmpw + 19, tmph1, &HE1A464)
		APILine(tmpw + 20, tmph, tmpw + 20, tmph1, &HE0A363) : APILine(tmpw + 21, tmph, tmpw + 21, tmph1, &HE0A363) : APILine(tmpw + 22, tmph, tmpw + 22, tmph1, &HE1A161)
		APILine(tmpw + 23, tmph, tmpw + 23, tmph1, &HE09F5B) : APILine(tmpw + 24, tmph, tmpw + 24, tmph1, &HDE9855) : APILine(tmpw + 25, tmph, tmpw + 25, tmph1, &HDC9752)
		APILine(tmpw + 26, tmph, tmpw + 26, tmph1, &HDB934B) : APILine(tmpw + 27, tmph, tmpw + 27, tmph1, &HD68D39) : APILine(tmpw + 28, tmph, tmpw + 28, tmph1, &HD17F2D)
		APILine(tmpw + 29, tmph, tmpw + 29, tmph1, &HD67426) : APILine(tmpw + 30, tmph, tmpw + 30, tmph1, &HC05D13) : APILine(tmpw + 31, tmph, tmpw + 31, tmph1, &H7C3514)
		APILine(tmpw + 32, tmph, tmpw + 32, tmph1, &HAB9B98) : APILine(tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6)
		'Lineas verticales
		APILine(17, 0, lw - 17, 0, &H450608)
		APILine(17, 1, lw - 17, 1, &HF1D4C9)
		APILine(17, 2, lw - 17, 2, &HE5C8BD)
		APILine(17, 3, lw - 17, 3, &HE8C0A1)
		APILine(17, 4, lw - 17, 4, &HE0B898)
		APILine(17, 5, lw - 17, 5, &HE3B48E)
		APILine(17, 6, lw - 17, 6, &HE0B18B)
		APILine(17, 7, lw - 17, 7, &HE9B47F)
		APILine(17, 8, lw - 17, 8, &HCE9963)
		APILine(17, 9, lw - 17, 9, &HDDA064)
		APILine(17, 10, lw - 17, 10, &HE2A66A)
		APILine(17, 11, lw - 17, 11, &HE6AC76)
		tmph = lh - 22
		APILine(17, tmph + 11, lw - 17, tmph + 11, &HE6AC76)
		APILine(17, tmph + 12, lw - 17, tmph + 12, &HF1B681)
		APILine(17, tmph + 13, lw - 17, tmph + 13, &HF3BD8A)
		APILine(17, tmph + 14, lw - 17, tmph + 14, &HFCC592)
		APILine(17, tmph + 15, lw - 17, tmph + 15, &HF8CE97)
		APILine(17, tmph + 16, lw - 17, tmph + 16, &HFED59E)
		APILine(17, tmph + 17, lw - 17, tmph + 17, &HF7DDA3)
		APILine(17, tmph + 18, lw - 17, tmph + 18, &HFFE6AD)
		APILine(17, tmph + 19, lw - 17, tmph + 19, &HE9E2C5)
		APILine(17, tmph + 20, lw - 17, tmph + 20, &H635D40)
		APILine(17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5)
		APILine(17, tmph + 22, lw - 17, tmph + 22, &HECECEC)
		Exit Sub
		
DrawMacOSXButtonHot_Error: 
	End Sub
	
	Private Sub DrawMacOSXButtonPressed()
		'On Error GoTo DrawMacOSXButtonPressed_Error
		
		Dim lhdc As Integer
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lhdc = MyBase.hdc
		'Variable vars (real into code)
		Dim lh, lw As Integer
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) : lw = MyBase.ClientRectangle.Width
		Dim tmph, tmpw As Integer
		Dim tmph1, tmpw1 As Integer
		'UPGRADE_ISSUE: UserControl property ctlButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		APIFillRectByCoords(hdc, 18, 11, lw - 34, lh - 19, &HCC9B6A)
		SetPixelV(lhdc, 6, 0, &HFEFEFE) : SetPixelV(lhdc, 7, 0, &HE5E4E4) : SetPixelV(lhdc, 8, 0, &HA5A2A2) : SetPixelV(lhdc, 9, 0, &H675C5C) : SetPixelV(lhdc, 10, 0, &H422729) : SetPixelV(lhdc, 11, 0, &H300E0D) : SetPixelV(lhdc, 12, 0, &H300A09) : SetPixelV(lhdc, 13, 0, &H2F0908) : SetPixelV(lhdc, 14, 0, &H330909) : SetPixelV(lhdc, 15, 0, &H390A0A) : SetPixelV(lhdc, 16, 0, &H3C0A0A) : SetPixelV(lhdc, 17, 0, &H3C090A)
		SetPixelV(lhdc, 5, 1, &HF0EEEE) : SetPixelV(lhdc, 6, 1, &H9D888A) : SetPixelV(lhdc, 7, 1, &H653531) : SetPixelV(lhdc, 8, 1, &H5A201D) : SetPixelV(lhdc, 9, 1, &H8D655F) : SetPixelV(lhdc, 10, 1, &HB99995) : SetPixelV(lhdc, 11, 1, &HD0B4B2) : SetPixelV(lhdc, 12, 1, &HD7BEBB) : SetPixelV(lhdc, 13, 1, &HDDC6C0) : SetPixelV(lhdc, 14, 1, &HDDC6C0) : SetPixelV(lhdc, 15, 1, &HDDC7BE) : SetPixelV(lhdc, 16, 1, &HDDC7BE) : SetPixelV(lhdc, 17, 1, &HDEC7BE)
		SetPixelV(lhdc, 3, 2, &HFEFEFE) : SetPixelV(lhdc, 4, 2, &HE4E4E4) : SetPixelV(lhdc, 5, 2, &H6F5C5C) : SetPixelV(lhdc, 6, 2, &H390A0E) : SetPixelV(lhdc, 7, 2, &H712E2A) : SetPixelV(lhdc, 8, 2, &HD6928D) : SetPixelV(lhdc, 9, 2, &HD8ACA6) : SetPixelV(lhdc, 10, 2, &HD1B0AC) : SetPixelV(lhdc, 11, 2, &HD1B5B2) : SetPixelV(lhdc, 12, 2, &HD0B7B4) : SetPixelV(lhdc, 13, 2, &HCEB7B1) : SetPixelV(lhdc, 14, 2, &HCEB7B1) : SetPixelV(lhdc, 15, 2, &HD2BCB2) : SetPixelV(lhdc, 16, 2, &HD2BCB2) : SetPixelV(lhdc, 17, 2, &HD3BCB2)
		SetPixelV(lhdc, 3, 3, &HEEEDED) : SetPixelV(lhdc, 4, 3, &H805858) : SetPixelV(lhdc, 5, 3, &H6A0D08) : SetPixelV(lhdc, 6, 3, &H7D1909) : SetPixelV(lhdc, 7, 3, &HB07B63) : SetPixelV(lhdc, 8, 3, &HCFA58A) : SetPixelV(lhdc, 9, 3, &HCDA78E) : SetPixelV(lhdc, 10, 3, &HD1AB92) : SetPixelV(lhdc, 11, 3, &HD2AF93) : SetPixelV(lhdc, 12, 3, &HD3B094) : SetPixelV(lhdc, 13, 3, &HD0AF93) : SetPixelV(lhdc, 14, 3, &HD3B296) : SetPixelV(lhdc, 15, 3, &HD4B49A) : SetPixelV(lhdc, 16, 3, &HD4B39A) : SetPixelV(lhdc, 17, 3, &HD4B39A)
		SetPixelV(lhdc, 2, 4, &HFBFBFB) : SetPixelV(lhdc, 3, 4, &H837576) : SetPixelV(lhdc, 4, 4, &H440C0C) : SetPixelV(lhdc, 5, 4, &H821D0D) : SetPixelV(lhdc, 6, 4, &HA94433) : SetPixelV(lhdc, 7, 4, &HC08B72) : SetPixelV(lhdc, 8, 4, &HC49A7F) : SetPixelV(lhdc, 9, 4, &HC6A188) : SetPixelV(lhdc, 10, 4, &HC7A189) : SetPixelV(lhdc, 11, 4, &HC6A387) : SetPixelV(lhdc, 12, 4, &HC8A689) : SetPixelV(lhdc, 13, 4, &HC9A98C) : SetPixelV(lhdc, 14, 4, &HC8A88B) : SetPixelV(lhdc, 15, 4, &HCBAB91) : SetPixelV(lhdc, 16, 4, &HCCAC92) : SetPixelV(lhdc, 17, 4, &HCCAC92)
		SetPixelV(lhdc, 1, 5, &HFEFEFE) : SetPixelV(lhdc, 2, 5, &HCAC8C7) : SetPixelV(lhdc, 3, 5, &H79281D) : SetPixelV(lhdc, 4, 5, &H7F2409) : SetPixelV(lhdc, 5, 5, &H8C3809) : SetPixelV(lhdc, 6, 5, &HBD6D39) : SetPixelV(lhdc, 7, 5, &HC9986E) : SetPixelV(lhdc, 8, 5, &HC89D74) : SetPixelV(lhdc, 9, 5, &HC49A71) : SetPixelV(lhdc, 10, 5, &HCAA17C) : SetPixelV(lhdc, 11, 5, &HC6A07A) : SetPixelV(lhdc, 12, 5, &HCAA480) : SetPixelV(lhdc, 13, 5, &HCAA582) : SetPixelV(lhdc, 14, 5, &HCBA584) : SetPixelV(lhdc, 15, 5, &HCDA989) : SetPixelV(lhdc, 16, 5, &HCFA98A) : SetPixelV(lhdc, 17, 5, &HCFA88A)
		SetPixelV(lhdc, 1, 6, &HF9F9F9) : SetPixelV(lhdc, 2, 6, &H756C6B) : SetPixelV(lhdc, 3, 6, &H76190D) : SetPixelV(lhdc, 4, 6, &H913416) : SetPixelV(lhdc, 5, 6, &H9D4916) : SetPixelV(lhdc, 6, 6, &HBA6A36) : SetPixelV(lhdc, 7, 6, &HC39268) : SetPixelV(lhdc, 8, 6, &HC59A71) : SetPixelV(lhdc, 9, 6, &HC59B72) : SetPixelV(lhdc, 10, 6, &HC59C77) : SetPixelV(lhdc, 11, 6, &HC6A07A) : SetPixelV(lhdc, 12, 6, &HC6A07C) : SetPixelV(lhdc, 13, 6, &HC7A27F) : SetPixelV(lhdc, 14, 6, &HCBA584) : SetPixelV(lhdc, 15, 6, &HCAA686) : SetPixelV(lhdc, 16, 6, &HCBA586) : SetPixelV(lhdc, 17, 6, &HCCA587)
		SetPixelV(lhdc, 1, 7, &HE8E7E7) : SetPixelV(lhdc, 2, 7, &H6C3E35) : SetPixelV(lhdc, 3, 7, &H8A2D09) : SetPixelV(lhdc, 4, 7, &HA34812) : SetPixelV(lhdc, 5, 7, &HAB591A) : SetPixelV(lhdc, 6, 7, &HB46B2B) : SetPixelV(lhdc, 7, 7, &HC3854A) : SetPixelV(lhdc, 8, 7, &HD19C64) : SetPixelV(lhdc, 9, 7, &HCD9C6C) : SetPixelV(lhdc, 10, 7, &HD1A070) : SetPixelV(lhdc, 11, 7, &HD2A272) : SetPixelV(lhdc, 12, 7, &HD2A272) : SetPixelV(lhdc, 13, 7, &HD6A57A) : SetPixelV(lhdc, 14, 7, &HD8A77C) : SetPixelV(lhdc, 15, 7, &HD2A87C) : SetPixelV(lhdc, 16, 7, &HD2A87C) : SetPixelV(lhdc, 17, 7, &HD2A77D)
		SetPixelV(lhdc, 0, 8, &HFDFDFD) : SetPixelV(lhdc, 1, 8, &HC7C3C3) : SetPixelV(lhdc, 2, 8, &H5C2A21) : SetPixelV(lhdc, 3, 8, &H9C3E15) : SetPixelV(lhdc, 4, 8, &HB35A22) : SetPixelV(lhdc, 5, 8, &HB56324) : SetPixelV(lhdc, 6, 8, &HB66D2D) : SetPixelV(lhdc, 7, 8, &HB6783D) : SetPixelV(lhdc, 8, 8, &HB07B44) : SetPixelV(lhdc, 9, 8, &HB18050) : SetPixelV(lhdc, 10, 8, &HB58454) : SetPixelV(lhdc, 11, 8, &HB48554) : SetPixelV(lhdc, 12, 8, &HB78858) : SetPixelV(lhdc, 13, 8, &HBA895E) : SetPixelV(lhdc, 14, 8, &HBB8A5F) : SetPixelV(lhdc, 15, 8, &HB98E62) : SetPixelV(lhdc, 16, 8, &HB98E62) : SetPixelV(lhdc, 17, 8, &HB98E62)
		SetPixelV(lhdc, 0, 9, &HFAFAFA) : SetPixelV(lhdc, 1, 9, &HB4ABA9) : SetPixelV(lhdc, 2, 9, &H612A14) : SetPixelV(lhdc, 3, 9, &HA05316) : SetPixelV(lhdc, 4, 9, &HB36628) : SetPixelV(lhdc, 5, 9, &HB67132) : SetPixelV(lhdc, 6, 9, &HB67738) : SetPixelV(lhdc, 7, 9, &HB98146) : SetPixelV(lhdc, 8, 9, &HBD864E) : SetPixelV(lhdc, 9, 9, &HBD894F) : SetPixelV(lhdc, 10, 9, &HC28D55) : SetPixelV(lhdc, 11, 9, &HC4905B) : SetPixelV(lhdc, 12, 9, &HC5905F) : SetPixelV(lhdc, 13, 9, &HC49161) : SetPixelV(lhdc, 14, 9, &HC49161) : SetPixelV(lhdc, 15, 9, &HC69564) : SetPixelV(lhdc, 16, 9, &HC69564) : SetPixelV(lhdc, 17, 9, &HC69464)
		SetPixelV(lhdc, 0, 10, &HF7F7F7) : SetPixelV(lhdc, 1, 10, &HA99D9B) : SetPixelV(lhdc, 2, 10, &H632D17) : SetPixelV(lhdc, 3, 10, &HA65A1D) : SetPixelV(lhdc, 4, 10, &HB96C2E) : SetPixelV(lhdc, 5, 10, &HBC7738) : SetPixelV(lhdc, 6, 10, &HC18242) : SetPixelV(lhdc, 7, 10, &HC2894E) : SetPixelV(lhdc, 8, 10, &HC18A52) : SetPixelV(lhdc, 9, 10, &HC59157) : SetPixelV(lhdc, 10, 10, &HC59159) : SetPixelV(lhdc, 11, 10, &HCC9863) : SetPixelV(lhdc, 12, 10, &HCC9665) : SetPixelV(lhdc, 13, 10, &HCB9767) : SetPixelV(lhdc, 14, 10, &HC99565) : SetPixelV(lhdc, 15, 10, &HCC9A6A) : SetPixelV(lhdc, 16, 10, &HCC9A6A) : SetPixelV(lhdc, 17, 10, &HCC9B6A)
		tmph = lh - 22
		SetPixelV(lhdc, 0, tmph + 10, &HF7F7F7) : SetPixelV(lhdc, 1, tmph + 10, &HA99D9B) : SetPixelV(lhdc, 2, tmph + 10, &H632D17) : SetPixelV(lhdc, 3, tmph + 10, &HA65A1D) : SetPixelV(lhdc, 4, tmph + 10, &HB96C2E) : SetPixelV(lhdc, 5, tmph + 10, &HBC7738) : SetPixelV(lhdc, 6, tmph + 10, &HC18242) : SetPixelV(lhdc, 7, tmph + 10, &HC2894E) : SetPixelV(lhdc, 8, tmph + 10, &HC18A52) : SetPixelV(lhdc, 9, tmph + 10, &HC59157) : SetPixelV(lhdc, 10, tmph + 10, &HC59159) : SetPixelV(lhdc, 11, tmph + 10, &HCC9863) : SetPixelV(lhdc, 12, tmph + 10, &HCC9665) : SetPixelV(lhdc, 13, tmph + 10, &HCB9767) : SetPixelV(lhdc, 14, tmph + 10, &HC99565) : SetPixelV(lhdc, 15, tmph + 10, &HCC9A6A) : SetPixelV(lhdc, 16, tmph + 10, &HCC9A6A) : SetPixelV(lhdc, 17, tmph + 10, &HCC9B6A)
		SetPixelV(lhdc, 0, tmph + 11, &HF5F5F5) : SetPixelV(lhdc, 1, tmph + 11, &HA59F9A) : SetPixelV(lhdc, 2, tmph + 11, &H674024) : SetPixelV(lhdc, 3, tmph + 11, &HAE6827) : SetPixelV(lhdc, 4, tmph + 11, &HB97231) : SetPixelV(lhdc, 5, tmph + 11, &HBE8247) : SetPixelV(lhdc, 6, tmph + 11, &HC0874E) : SetPixelV(lhdc, 7, tmph + 11, &HC78E56) : SetPixelV(lhdc, 8, tmph + 11, &HCD9561) : SetPixelV(lhdc, 9, tmph + 11, &HCB9466) : SetPixelV(lhdc, 10, tmph + 11, &HCD9A6B) : SetPixelV(lhdc, 11, tmph + 11, &HC79867) : SetPixelV(lhdc, 12, tmph + 11, &HCA9B6A) : SetPixelV(lhdc, 13, tmph + 11, &HCC9D6C) : SetPixelV(lhdc, 14, tmph + 11, &HCD9D70) : SetPixelV(lhdc, 15, tmph + 11, &HD0A175) : SetPixelV(lhdc, 16, tmph + 11, &HD0A175) : SetPixelV(lhdc, 17, tmph + 11, &HD0A175)
		SetPixelV(lhdc, 0, tmph + 12, &HF5F5F5) : SetPixelV(lhdc, 1, tmph + 12, &HACA7A4) : SetPixelV(lhdc, 2, tmph + 12, &H755035) : SetPixelV(lhdc, 3, tmph + 12, &HB77131) : SetPixelV(lhdc, 4, tmph + 12, &HCB8443) : SetPixelV(lhdc, 5, tmph + 12, &HC5894E) : SetPixelV(lhdc, 6, tmph + 12, &HCC935A) : SetPixelV(lhdc, 7, tmph + 12, &HD29962) : SetPixelV(lhdc, 8, tmph + 12, &HD69F6A) : SetPixelV(lhdc, 9, tmph + 12, &HDBA476) : SetPixelV(lhdc, 10, tmph + 12, &HD6A374) : SetPixelV(lhdc, 11, tmph + 12, &HD4A574) : SetPixelV(lhdc, 12, tmph + 12, &HD8A978) : SetPixelV(lhdc, 13, tmph + 12, &HDAAB7A) : SetPixelV(lhdc, 14, tmph + 12, &HDAAA7D) : SetPixelV(lhdc, 15, tmph + 12, &HDBAB7F) : SetPixelV(lhdc, 16, tmph + 12, &HDAAA7F) : SetPixelV(lhdc, 17, tmph + 12, &HDAAA7F)
		SetPixelV(lhdc, 0, tmph + 13, &HF7F7F7) : SetPixelV(lhdc, 1, tmph + 13, &HC0C0BF) : SetPixelV(lhdc, 2, tmph + 13, &H63574B) : SetPixelV(lhdc, 3, tmph + 13, &HAC7036) : SetPixelV(lhdc, 4, tmph + 13, &HC2854A) : SetPixelV(lhdc, 5, tmph + 13, &HCF955E) : SetPixelV(lhdc, 6, tmph + 13, &HD29B66) : SetPixelV(lhdc, 7, tmph + 13, &HD1A26E) : SetPixelV(lhdc, 8, tmph + 13, &HD8A776) : SetPixelV(lhdc, 9, tmph + 13, &HDBA878) : SetPixelV(lhdc, 10, tmph + 13, &HDFAC7C) : SetPixelV(lhdc, 11, tmph + 13, &HDBAF7D) : SetPixelV(lhdc, 12, tmph + 13, &HDDAF81) : SetPixelV(lhdc, 13, tmph + 13, &HDEB183) : SetPixelV(lhdc, 14, tmph + 13, &HDDAF84) : SetPixelV(lhdc, 15, tmph + 13, &HDEB087) : SetPixelV(lhdc, 16, tmph + 13, &HDEB087) : SetPixelV(lhdc, 17, tmph + 13, &HDCB087)
		SetPixelV(lhdc, 0, tmph + 14, &HFBFBFB) : SetPixelV(lhdc, 1, tmph + 14, &HE1E1E1) : SetPixelV(lhdc, 2, tmph + 14, &H7C7269) : SetPixelV(lhdc, 3, tmph + 14, &HA26830) : SetPixelV(lhdc, 4, tmph + 14, &HC6884E) : SetPixelV(lhdc, 5, tmph + 14, &HD0965F) : SetPixelV(lhdc, 6, tmph + 14, &HDAA26E) : SetPixelV(lhdc, 7, tmph + 14, &HD9AA75) : SetPixelV(lhdc, 8, tmph + 14, &HDBAA79) : SetPixelV(lhdc, 9, tmph + 14, &HE2AF7F) : SetPixelV(lhdc, 10, tmph + 14, &HE6B484) : SetPixelV(lhdc, 11, tmph + 14, &HE2B684) : SetPixelV(lhdc, 12, tmph + 14, &HE3B588) : SetPixelV(lhdc, 13, tmph + 14, &HE2B587) : SetPixelV(lhdc, 14, tmph + 14, &HE2B48A) : SetPixelV(lhdc, 15, tmph + 14, &HE5B78E) : SetPixelV(lhdc, 16, tmph + 14, &HE5B78E) : SetPixelV(lhdc, 17, tmph + 14, &HE4B88E)
		SetPixelV(lhdc, 0, tmph + 15, &HFEFEFE) : SetPixelV(lhdc, 1, tmph + 15, &HEDEDED) : SetPixelV(lhdc, 2, tmph + 15, &H9E9C9C) : SetPixelV(lhdc, 3, tmph + 15, &H766051) : SetPixelV(lhdc, 4, tmph + 15, &HAD8666) : SetPixelV(lhdc, 5, tmph + 15, &HD49A61) : SetPixelV(lhdc, 6, tmph + 15, &HE0A66D) : SetPixelV(lhdc, 7, tmph + 15, &HE3B17C) : SetPixelV(lhdc, 8, tmph + 15, &HE0B380) : SetPixelV(lhdc, 9, tmph + 15, &HE0B587) : SetPixelV(lhdc, 10, tmph + 15, &HE2BC8C) : SetPixelV(lhdc, 11, tmph + 15, &HE0BB8B) : SetPixelV(lhdc, 12, tmph + 15, &HE0BC8B) : SetPixelV(lhdc, 13, tmph + 15, &HE3BD92) : SetPixelV(lhdc, 14, tmph + 15, &HE2BC91) : SetPixelV(lhdc, 15, tmph + 15, &HE2BF93) : SetPixelV(lhdc, 16, tmph + 15, &HE1BE93) : SetPixelV(lhdc, 17, tmph + 15, &HE1BF93)
		SetPixelV(lhdc, 1, tmph + 16, &HF6F6F6) : SetPixelV(lhdc, 2, tmph + 16, &HD5D5D5) : SetPixelV(lhdc, 3, tmph + 16, &H86766C) : SetPixelV(lhdc, 4, tmph + 16, &H856144) : SetPixelV(lhdc, 5, tmph + 16, &HD59C63) : SetPixelV(lhdc, 6, tmph + 16, &HE5AB71) : SetPixelV(lhdc, 7, tmph + 16, &HE5B37E) : SetPixelV(lhdc, 8, tmph + 16, &HE7BB88) : SetPixelV(lhdc, 9, tmph + 16, &HE7BF91) : SetPixelV(lhdc, 10, tmph + 16, &HE3BC8D) : SetPixelV(lhdc, 11, tmph + 16, &HE7C392) : SetPixelV(lhdc, 12, tmph + 16, &HE7C392) : SetPixelV(lhdc, 13, tmph + 16, &HE8C398) : SetPixelV(lhdc, 14, tmph + 16, &HE8C499) : SetPixelV(lhdc, 15, tmph + 16, &HE8C599) : SetPixelV(lhdc, 16, tmph + 16, &HE8C599) : SetPixelV(lhdc, 17, tmph + 16, &HE7C699)
		SetPixelV(lhdc, 1, tmph + 17, &HFDFDFD) : SetPixelV(lhdc, 2, tmph + 17, &HEDEDED) : SetPixelV(lhdc, 3, tmph + 17, &HBDBDBD) : SetPixelV(lhdc, 4, tmph + 17, &H676767) : SetPixelV(lhdc, 5, tmph + 17, &H71604C) : SetPixelV(lhdc, 6, tmph + 17, &HBEA17D) : SetPixelV(lhdc, 7, tmph + 17, &HDAB381) : SetPixelV(lhdc, 8, tmph + 17, &HE5BE8C) : SetPixelV(lhdc, 9, tmph + 17, &HE1C18F) : SetPixelV(lhdc, 10, tmph + 17, &HE4C895) : SetPixelV(lhdc, 11, tmph + 17, &HDFCA98) : SetPixelV(lhdc, 12, tmph + 17, &HE2CE9B) : SetPixelV(lhdc, 13, tmph + 17, &HE2CE9B) : SetPixelV(lhdc, 14, tmph + 17, &HE2CE9B) : SetPixelV(lhdc, 15, tmph + 17, &HE2CD9D) : SetPixelV(lhdc, 16, tmph + 17, &HE2CC9D) : SetPixelV(lhdc, 17, tmph + 17, &HE2CC9D)
		SetPixelV(lhdc, 2, tmph + 18, &HF9F9F9) : SetPixelV(lhdc, 3, tmph + 18, &HE6E6E6) : SetPixelV(lhdc, 4, tmph + 18, &HB9B9B9) : SetPixelV(lhdc, 5, tmph + 18, &H7A7163) : SetPixelV(lhdc, 6, tmph + 18, &H776043) : SetPixelV(lhdc, 7, tmph + 18, &HAB885B) : SetPixelV(lhdc, 8, tmph + 18, &HDDB888) : SetPixelV(lhdc, 9, tmph + 18, &HE6C796) : SetPixelV(lhdc, 10, tmph + 18, &HE8CD9A) : SetPixelV(lhdc, 11, tmph + 18, &HE5D19E) : SetPixelV(lhdc, 12, tmph + 18, &HE9D6A3) : SetPixelV(lhdc, 13, tmph + 18, &HE9D6A5) : SetPixelV(lhdc, 14, tmph + 18, &HE9D6A3) : SetPixelV(lhdc, 15, tmph + 18, &HE9D5A6) : SetPixelV(lhdc, 16, tmph + 18, &HE9D5A6) : SetPixelV(lhdc, 17, tmph + 18, &HE9D5A6)
		SetPixelV(lhdc, 2, tmph + 19, &HFEFEFE) : SetPixelV(lhdc, 3, tmph + 19, &HF8F8F8) : SetPixelV(lhdc, 4, tmph + 19, &HE6E6E6) : SetPixelV(lhdc, 5, tmph + 19, &HC8C8C8) : SetPixelV(lhdc, 6, tmph + 19, &H8C8C8C) : SetPixelV(lhdc, 7, tmph + 19, &H61605E) : SetPixelV(lhdc, 8, tmph + 19, &H656059) : SetPixelV(lhdc, 9, tmph + 19, &H857C6D) : SetPixelV(lhdc, 10, tmph + 19, &HA59C87) : SetPixelV(lhdc, 11, tmph + 19, &HC8C1A8) : SetPixelV(lhdc, 12, tmph + 19, &HD1CAB0) : SetPixelV(lhdc, 13, tmph + 19, &HD5CFB5) : SetPixelV(lhdc, 14, tmph + 19, &HD6D1B6) : SetPixelV(lhdc, 15, tmph + 19, &HD7D2BA) : SetPixelV(lhdc, 16, tmph + 19, &HD7D1BA) : SetPixelV(lhdc, 17, tmph + 19, &HD7D2BA)
		SetPixelV(lhdc, 3, tmph + 20, &HFEFEFE) : SetPixelV(lhdc, 4, tmph + 20, &HF9F9F9) : SetPixelV(lhdc, 5, tmph + 20, &HECECEC) : SetPixelV(lhdc, 6, tmph + 20, &HDADADA) : SetPixelV(lhdc, 7, tmph + 20, &HC1C1C1) : SetPixelV(lhdc, 8, tmph + 20, &H9C9B99) : SetPixelV(lhdc, 9, tmph + 20, &H7D7A73) : SetPixelV(lhdc, 10, tmph + 20, &H635E50) : SetPixelV(lhdc, 11, tmph + 20, &H58533F) : SetPixelV(lhdc, 12, tmph + 20, &H554F39) : SetPixelV(lhdc, 13, tmph + 20, &H514D36) : SetPixelV(lhdc, 14, tmph + 20, &H554F37) : SetPixelV(lhdc, 15, tmph + 20, &H57523A) : SetPixelV(lhdc, 16, tmph + 20, &H5A563D) : SetPixelV(lhdc, 17, tmph + 20, &H5A563E)
		SetPixelV(lhdc, 5, tmph + 21, &HFCFCFC) : SetPixelV(lhdc, 6, tmph + 21, &HF5F5F5) : SetPixelV(lhdc, 7, tmph + 21, &HEBEBEB) : SetPixelV(lhdc, 8, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, 9, tmph + 21, &HD6D6D6) : SetPixelV(lhdc, 10, tmph + 21, &HCECECE) : SetPixelV(lhdc, 11, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, 12, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 13, tmph + 21, &HC7C7C7) : SetPixelV(lhdc, 14, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 15, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, 16, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, 17, tmph + 21, &HC5C5C5)
		SetPixelV(lhdc, 7, tmph + 22, &HFDFDFD) : SetPixelV(lhdc, 8, tmph + 22, &HF9F9F9) : SetPixelV(lhdc, 9, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, 10, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, 11, tmph + 22, &HEEEEEE) : SetPixelV(lhdc, 12, tmph + 22, &HEDEDED) : SetPixelV(lhdc, 13, tmph + 22, &HECECEC) : SetPixelV(lhdc, 14, tmph + 22, &HECECEC) : SetPixelV(lhdc, 15, tmph + 22, &HECECEC) : SetPixelV(lhdc, 16, tmph + 22, &HECECEC) : SetPixelV(lhdc, 17, tmph + 22, &HECECEC)
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, 0, &H3C090A) : SetPixelV(lhdc, tmpw + 18, 0, &H3C090A) : SetPixelV(lhdc, tmpw + 19, 0, &H340A0A) : SetPixelV(lhdc, tmpw + 20, 0, &H300A09) : SetPixelV(lhdc, tmpw + 21, 0, &H2F080A) : SetPixelV(lhdc, tmpw + 22, 0, &H341011) : SetPixelV(lhdc, tmpw + 23, 0, &H3E2526) : SetPixelV(lhdc, tmpw + 24, 0, &H5A4C4C) : SetPixelV(lhdc, tmpw + 25, 0, &H9E9B9B) : SetPixelV(lhdc, tmpw + 26, 0, &HEEEEEE) : SetPixelV(lhdc, tmpw + 34, 0, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 1, &HDEC7BE) : SetPixelV(lhdc, tmpw + 18, 1, &HDEC7BE) : SetPixelV(lhdc, tmpw + 19, 1, &HDBC6C1) : SetPixelV(lhdc, tmpw + 20, 1, &HD9C4BF) : SetPixelV(lhdc, tmpw + 21, 1, &HD7C1B9) : SetPixelV(lhdc, tmpw + 22, 1, &HD3B5AF) : SetPixelV(lhdc, tmpw + 23, 1, &HBE9F97) : SetPixelV(lhdc, tmpw + 24, 1, &H9B6A65) : SetPixelV(lhdc, tmpw + 25, 1, &H65231E) : SetPixelV(lhdc, tmpw + 26, 1, &H642A26) : SetPixelV(lhdc, tmpw + 27, 1, &HA59696) : SetPixelV(lhdc, tmpw + 28, 1, &HF7F7F7) : SetPixelV(lhdc, tmpw + 34, 1, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 2, &HD3BCB2) : SetPixelV(lhdc, tmpw + 18, 2, &HD3BCB2) : SetPixelV(lhdc, tmpw + 19, 2, &HCDB8B3) : SetPixelV(lhdc, tmpw + 20, 2, &HCBB6B1) : SetPixelV(lhdc, tmpw + 21, 2, &HD0BBB2) : SetPixelV(lhdc, tmpw + 22, 2, &HD0B2AC) : SetPixelV(lhdc, tmpw + 23, 2, &HD6B6AF) : SetPixelV(lhdc, tmpw + 24, 2, &HDCABA6) : SetPixelV(lhdc, tmpw + 25, 2, &HDC9691) : SetPixelV(lhdc, tmpw + 26, 2, &H732E29) : SetPixelV(lhdc, tmpw + 27, 2, &H380A0A) : SetPixelV(lhdc, tmpw + 28, 2, &H6A5556) : SetPixelV(lhdc, tmpw + 29, 2, &HEAEBEA) : SetPixelV(lhdc, tmpw + 34, 2, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 3, &HD4B39A) : SetPixelV(lhdc, tmpw + 18, 3, &HD4B39A) : SetPixelV(lhdc, tmpw + 19, 3, &HD1B294) : SetPixelV(lhdc, tmpw + 20, 3, &HD0B193) : SetPixelV(lhdc, tmpw + 21, 3, &HD0AE91) : SetPixelV(lhdc, tmpw + 22, 3, &HD4B296) : SetPixelV(lhdc, tmpw + 23, 3, &HCBAA8F) : SetPixelV(lhdc, tmpw + 24, 3, &HCBAA8F) : SetPixelV(lhdc, tmpw + 25, 3, &HCCA38B) : SetPixelV(lhdc, tmpw + 26, 3, &HB77E68) : SetPixelV(lhdc, tmpw + 27, 3, &H811B09) : SetPixelV(lhdc, tmpw + 28, 3, &H720E08) : SetPixelV(lhdc, tmpw + 29, 3, &H7D5051) : SetPixelV(lhdc, tmpw + 30, 3, &HEFEEEE) : SetPixelV(lhdc, tmpw + 34, 3, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 4, &HCCAC92) : SetPixelV(lhdc, tmpw + 18, 4, &HCCAC91) : SetPixelV(lhdc, tmpw + 19, 4, &HC6A889) : SetPixelV(lhdc, tmpw + 20, 4, &HC7A98A) : SetPixelV(lhdc, tmpw + 21, 4, &HC7A589) : SetPixelV(lhdc, tmpw + 22, 4, &HC4A185) : SetPixelV(lhdc, tmpw + 23, 4, &HC6A58A) : SetPixelV(lhdc, tmpw + 24, 4, &HBF9E83) : SetPixelV(lhdc, tmpw + 25, 4, &HC39A82) : SetPixelV(lhdc, tmpw + 26, 4, &HC58C76) : SetPixelV(lhdc, tmpw + 27, 4, &HA9432F) : SetPixelV(lhdc, tmpw + 28, 4, &H861F0C) : SetPixelV(lhdc, tmpw + 29, 4, &H460B0C) : SetPixelV(lhdc, tmpw + 30, 4, &H7B6B6C) : SetPixelV(lhdc, tmpw + 31, 4, &HFAFAFA) : SetPixelV(lhdc, tmpw + 34, 4, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 5, &HCFA88A) : SetPixelV(lhdc, tmpw + 18, 5, &HCFA889) : SetPixelV(lhdc, tmpw + 19, 5, &HCBA683) : SetPixelV(lhdc, tmpw + 20, 5, &HC9A481) : SetPixelV(lhdc, tmpw + 21, 5, &HCCA480) : SetPixelV(lhdc, tmpw + 22, 5, &HCEA280) : SetPixelV(lhdc, tmpw + 23, 5, &HCCA379) : SetPixelV(lhdc, tmpw + 24, 5, &HCA9E74) : SetPixelV(lhdc, tmpw + 25, 5, &HC69971) : SetPixelV(lhdc, tmpw + 26, 5, &HC89870) : SetPixelV(lhdc, tmpw + 27, 5, &HB46A34) : SetPixelV(lhdc, tmpw + 28, 5, &H90380A) : SetPixelV(lhdc, tmpw + 29, 5, &H892509) : SetPixelV(lhdc, tmpw + 30, 5, &H8A251B) : SetPixelV(lhdc, tmpw + 31, 5, &HC4C2C2) : SetPixelV(lhdc, tmpw + 32, 5, &HFEFEFE) : SetPixelV(lhdc, tmpw + 34, 5, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 6, &HCCA587) : SetPixelV(lhdc, tmpw + 18, 6, &HCCA586) : SetPixelV(lhdc, tmpw + 19, 6, &HC9A481) : SetPixelV(lhdc, tmpw + 20, 6, &HC9A481) : SetPixelV(lhdc, tmpw + 21, 6, &HC7A07C) : SetPixelV(lhdc, tmpw + 22, 6, &HCCA17E) : SetPixelV(lhdc, tmpw + 23, 6, &HC79F74) : SetPixelV(lhdc, tmpw + 24, 6, &HC69A70) : SetPixelV(lhdc, tmpw + 25, 6, &HC59870) : SetPixelV(lhdc, tmpw + 26, 6, &HC2926A) : SetPixelV(lhdc, tmpw + 27, 6, &HB96F39) : SetPixelV(lhdc, tmpw + 28, 6, &HA04814) : SetPixelV(lhdc, tmpw + 29, 6, &H973215) : SetPixelV(lhdc, tmpw + 30, 6, &H831A0F) : SetPixelV(lhdc, tmpw + 31, 6, &H6E6966) : SetPixelV(lhdc, tmpw + 32, 6, &HF8F8F8) : SetPixelV(lhdc, tmpw + 34, 6, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 7, &HD2A77D) : SetPixelV(lhdc, tmpw + 18, 7, &HD3A77C) : SetPixelV(lhdc, tmpw + 19, 7, &HD8AA7D) : SetPixelV(lhdc, tmpw + 20, 7, &HD2A376) : SetPixelV(lhdc, tmpw + 21, 7, &HD1A373) : SetPixelV(lhdc, tmpw + 22, 7, &HCEA070) : SetPixelV(lhdc, tmpw + 23, 7, &HD2A06F) : SetPixelV(lhdc, tmpw + 24, 7, &HD19D68) : SetPixelV(lhdc, tmpw + 25, 7, &HD09A65) : SetPixelV(lhdc, tmpw + 26, 7, &HC2864F) : SetPixelV(lhdc, tmpw + 27, 7, &HAE6927) : SetPixelV(lhdc, tmpw + 28, 7, &HA95A19) : SetPixelV(lhdc, tmpw + 29, 7, &HA44A10) : SetPixelV(lhdc, tmpw + 30, 7, &H8B2E09) : SetPixelV(lhdc, tmpw + 31, 7, &H6B3E34) : SetPixelV(lhdc, tmpw + 32, 7, &HE7E6E6) : SetPixelV(lhdc, tmpw + 34, 7, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 8, &HB98E62) : SetPixelV(lhdc, tmpw + 18, 8, &HBA8E62) : SetPixelV(lhdc, tmpw + 19, 8, &HB98B5E) : SetPixelV(lhdc, tmpw + 20, 8, &HB98B5E) : SetPixelV(lhdc, tmpw + 21, 8, &HB68858) : SetPixelV(lhdc, tmpw + 22, 8, &HB48656) : SetPixelV(lhdc, tmpw + 23, 8, &HB58452) : SetPixelV(lhdc, tmpw + 24, 8, &HB5814C) : SetPixelV(lhdc, tmpw + 25, 8, &HB07A46) : SetPixelV(lhdc, tmpw + 26, 8, &HB2773F) : SetPixelV(lhdc, tmpw + 27, 8, &HB36E2C) : SetPixelV(lhdc, tmpw + 28, 8, &HB26221) : SetPixelV(lhdc, tmpw + 29, 8, &HB35A20) : SetPixelV(lhdc, tmpw + 30, 8, &H9C3E11) : SetPixelV(lhdc, tmpw + 31, 8, &H5C2A1F) : SetPixelV(lhdc, tmpw + 32, 8, &HC4C1C0) : SetPixelV(lhdc, tmpw + 33, 8, &HFDFDFD) : SetPixelV(lhdc, tmpw + 34, 8, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 9, &HC69464) : SetPixelV(lhdc, tmpw + 18, 9, &HC69564) : SetPixelV(lhdc, tmpw + 19, 9, &HC3925E) : SetPixelV(lhdc, tmpw + 20, 9, &HC3915D) : SetPixelV(lhdc, tmpw + 21, 9, &HC3925E) : SetPixelV(lhdc, tmpw + 22, 9, &HC38F5B) : SetPixelV(lhdc, tmpw + 23, 9, &HC28D55) : SetPixelV(lhdc, tmpw + 24, 9, &HC08751) : SetPixelV(lhdc, tmpw + 25, 9, &HBC844C) : SetPixelV(lhdc, tmpw + 26, 9, &HBC8147) : SetPixelV(lhdc, tmpw + 27, 9, &HB57936) : SetPixelV(lhdc, tmpw + 28, 9, &HB3702D) : SetPixelV(lhdc, tmpw + 29, 9, &HB56626) : SetPixelV(lhdc, tmpw + 30, 9, &HA25115) : SetPixelV(lhdc, tmpw + 31, 9, &H662D12) : SetPixelV(lhdc, tmpw + 32, 9, &HAEA3A1) : SetPixelV(lhdc, tmpw + 33, 9, &HF9F9F9) : SetPixelV(lhdc, tmpw + 34, 9, &HFFFFFFFF)
		SetPixelV(lhdc, tmpw + 17, 10, &HCC9B6A) : SetPixelV(lhdc, tmpw + 18, 10, &HCC9B6A) : SetPixelV(lhdc, tmpw + 19, 10, &HCA9864) : SetPixelV(lhdc, tmpw + 20, 10, &HC99763) : SetPixelV(lhdc, tmpw + 21, 10, &HC99763) : SetPixelV(lhdc, tmpw + 22, 10, &HCA9562) : SetPixelV(lhdc, tmpw + 23, 10, &HC9945D) : SetPixelV(lhdc, tmpw + 24, 10, &HC68E57) : SetPixelV(lhdc, tmpw + 25, 10, &HC48C55) : SetPixelV(lhdc, tmpw + 26, 10, &HC3884E) : SetPixelV(lhdc, tmpw + 27, 10, &HBE823E) : SetPixelV(lhdc, tmpw + 28, 10, &HB97634) : SetPixelV(lhdc, tmpw + 29, 10, &HBD6D2D) : SetPixelV(lhdc, tmpw + 30, 10, &HA8581C) : SetPixelV(lhdc, tmpw + 31, 10, &H6D3319) : SetPixelV(lhdc, tmpw + 32, 10, &HA49794) : SetPixelV(lhdc, tmpw + 33, 10, &HF6F6F6) : SetPixelV(lhdc, tmpw + 34, 10, &HFFFFFFFF)
		tmph = lh - 22
		tmpw = lw - 34
		SetPixelV(lhdc, tmpw + 17, tmph + 10, &HCC9B6A) : SetPixelV(lhdc, tmpw + 18, tmph + 10, &HCC9B6A) : SetPixelV(lhdc, tmpw + 19, tmph + 10, &HCA9864) : SetPixelV(lhdc, tmpw + 20, tmph + 10, &HC99763) : SetPixelV(lhdc, tmpw + 21, tmph + 10, &HC99763) : SetPixelV(lhdc, tmpw + 22, tmph + 10, &HCA9562) : SetPixelV(lhdc, tmpw + 23, tmph + 10, &HC9945D) : SetPixelV(lhdc, tmpw + 24, tmph + 10, &HC68E57) : SetPixelV(lhdc, tmpw + 25, tmph + 10, &HC48C55) : SetPixelV(lhdc, tmpw + 26, tmph + 10, &HC3884E) : SetPixelV(lhdc, tmpw + 27, tmph + 10, &HBE823E) : SetPixelV(lhdc, tmpw + 28, tmph + 10, &HB97634) : SetPixelV(lhdc, tmpw + 29, tmph + 10, &HBD6D2D) : SetPixelV(lhdc, tmpw + 30, tmph + 10, &HA8581C) : SetPixelV(lhdc, tmpw + 31, tmph + 10, &H6D3319) : SetPixelV(lhdc, tmpw + 32, tmph + 10, &HA49794) : SetPixelV(lhdc, tmpw + 33, tmph + 10, &HF6F6F6)
		SetPixelV(lhdc, tmpw + 17, tmph + 11, &HD0A175) : SetPixelV(lhdc, tmpw + 18, tmph + 11, &HD0A175) : SetPixelV(lhdc, tmpw + 19, tmph + 11, &HCC9D6D) : SetPixelV(lhdc, tmpw + 20, tmph + 11, &HCE9B6C) : SetPixelV(lhdc, tmpw + 21, tmph + 11, &HCB9A6A) : SetPixelV(lhdc, tmpw + 22, tmph + 11, &HCD996A) : SetPixelV(lhdc, tmpw + 23, tmph + 11, &HCA9666) : SetPixelV(lhdc, tmpw + 24, tmph + 11, &HCF9865) : SetPixelV(lhdc, tmpw + 25, tmph + 11, &HCA9460) : SetPixelV(lhdc, tmpw + 26, tmph + 11, &HC78F57) : SetPixelV(lhdc, tmpw + 27, tmph + 11, &HC1864B) : SetPixelV(lhdc, tmpw + 28, tmph + 11, &HC08143) : SetPixelV(lhdc, tmpw + 29, tmph + 11, &HB7712E) : SetPixelV(lhdc, tmpw + 30, tmph + 11, &HB16A28) : SetPixelV(lhdc, tmpw + 31, tmph + 11, &H694321) : SetPixelV(lhdc, tmpw + 32, tmph + 11, &HA59F9B) : SetPixelV(lhdc, tmpw + 33, tmph + 11, &HF4F4F4)
		SetPixelV(lhdc, tmpw + 17, tmph + 12, &HDAAA7F) : SetPixelV(lhdc, tmpw + 18, tmph + 12, &HD9AB7E) : SetPixelV(lhdc, tmpw + 19, tmph + 12, &HDBAC7C) : SetPixelV(lhdc, tmpw + 20, tmph + 12, &HDDAA7B) : SetPixelV(lhdc, tmpw + 21, tmph + 12, &HDAA979) : SetPixelV(lhdc, tmpw + 22, tmph + 12, &HDAA677) : SetPixelV(lhdc, tmpw + 23, tmph + 12, &HD8A474) : SetPixelV(lhdc, tmpw + 24, tmph + 12, &HDBA471) : SetPixelV(lhdc, tmpw + 25, tmph + 12, &HD49F6A) : SetPixelV(lhdc, tmpw + 26, tmph + 12, &HD09861) : SetPixelV(lhdc, tmpw + 27, tmph + 12, &HD0955A) : SetPixelV(lhdc, tmpw + 28, tmph + 12, &HCC8D4F) : SetPixelV(lhdc, tmpw + 29, tmph + 12, &HCA8441) : SetPixelV(lhdc, tmpw + 30, tmph + 12, &HBB7532) : SetPixelV(lhdc, tmpw + 31, tmph + 12, &H7B5434) : SetPixelV(lhdc, tmpw + 32, tmph + 12, &HB1ACAA) : SetPixelV(lhdc, tmpw + 33, tmph + 12, &HF5F5F5)
		SetPixelV(lhdc, tmpw + 17, tmph + 13, &HDCB087) : SetPixelV(lhdc, tmpw + 18, tmph + 13, &HDCB087) : SetPixelV(lhdc, tmpw + 19, tmph + 13, &HDBAF81) : SetPixelV(lhdc, tmpw + 20, tmph + 13, &HDEAF82) : SetPixelV(lhdc, tmpw + 21, tmph + 13, &HDCAF81) : SetPixelV(lhdc, tmpw + 22, tmph + 13, &HDDAD7F) : SetPixelV(lhdc, tmpw + 23, tmph + 13, &HDBAC7B) : SetPixelV(lhdc, tmpw + 24, tmph + 13, &HDDAA7A) : SetPixelV(lhdc, tmpw + 25, tmph + 13, &HD9A775) : SetPixelV(lhdc, tmpw + 26, tmph + 13, &HD7A26E) : SetPixelV(lhdc, tmpw + 27, tmph + 13, &HCE9961) : SetPixelV(lhdc, tmpw + 28, tmph + 13, &HCC945C) : SetPixelV(lhdc, tmpw + 29, tmph + 13, &HC2854D) : SetPixelV(lhdc, tmpw + 30, tmph + 13, &HAF7239) : SetPixelV(lhdc, tmpw + 31, tmph + 13, &H695C4F) : SetPixelV(lhdc, tmpw + 32, tmph + 13, &HD5D5D5) : SetPixelV(lhdc, tmpw + 33, tmph + 13, &HF8F8F8)
		SetPixelV(lhdc, tmpw + 17, tmph + 14, &HE4B88E) : SetPixelV(lhdc, tmpw + 18, tmph + 14, &HE3B88E) : SetPixelV(lhdc, tmpw + 19, tmph + 14, &HE0B486) : SetPixelV(lhdc, tmpw + 20, tmph + 14, &HE4B689) : SetPixelV(lhdc, tmpw + 21, tmph + 14, &HE2B587) : SetPixelV(lhdc, tmpw + 22, tmph + 14, &HE5B588) : SetPixelV(lhdc, tmpw + 23, tmph + 14, &HE3B483) : SetPixelV(lhdc, tmpw + 24, tmph + 14, &HE2AF7F) : SetPixelV(lhdc, tmpw + 25, tmph + 14, &HDEAC7A) : SetPixelV(lhdc, tmpw + 26, tmph + 14, &HDEAA75) : SetPixelV(lhdc, tmpw + 27, tmph + 14, &HD8A36B) : SetPixelV(lhdc, tmpw + 28, tmph + 14, &HD09860) : SetPixelV(lhdc, tmpw + 29, tmph + 14, &HC58850) : SetPixelV(lhdc, tmpw + 30, tmph + 14, &HA56930) : SetPixelV(lhdc, tmpw + 31, tmph + 14, &H7B746C) : SetPixelV(lhdc, tmpw + 32, tmph + 14, &HE8E8E8) : SetPixelV(lhdc, tmpw + 33, tmph + 14, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 15, &HE1BF93) : SetPixelV(lhdc, tmpw + 18, tmph + 15, &HE2BE93) : SetPixelV(lhdc, tmpw + 19, tmph + 15, &HE2BF91) : SetPixelV(lhdc, tmpw + 20, tmph + 15, &HE1BD8F) : SetPixelV(lhdc, tmpw + 21, tmph + 15, &HE0BC8E) : SetPixelV(lhdc, tmpw + 22, tmph + 15, &HE2BC8E) : SetPixelV(lhdc, tmpw + 23, tmph + 15, &HE4BD8C) : SetPixelV(lhdc, tmpw + 24, tmph + 15, &HE0B685) : SetPixelV(lhdc, tmpw + 25, tmph + 15, &HDCB07E) : SetPixelV(lhdc, tmpw + 26, tmph + 15, &HE1AF7C) : SetPixelV(lhdc, tmpw + 27, tmph + 15, &HDEA66E) : SetPixelV(lhdc, tmpw + 28, tmph + 15, &HD19962) : SetPixelV(lhdc, tmpw + 29, tmph + 15, &HAD875D) : SetPixelV(lhdc, tmpw + 30, tmph + 15, &H7D6851) : SetPixelV(lhdc, tmpw + 31, tmph + 15, &HB9B9B9) : SetPixelV(lhdc, tmpw + 32, tmph + 15, &HF1F1F1)
		SetPixelV(lhdc, tmpw + 17, tmph + 16, &HE7C699) : SetPixelV(lhdc, tmpw + 18, tmph + 16, &HE8C69A) : SetPixelV(lhdc, tmpw + 19, tmph + 16, &HE7C496) : SetPixelV(lhdc, tmpw + 20, tmph + 16, &HE8C597) : SetPixelV(lhdc, tmpw + 21, tmph + 16, &HE5C294) : SetPixelV(lhdc, tmpw + 22, tmph + 16, &HE8C194) : SetPixelV(lhdc, tmpw + 23, tmph + 16, &HE6BF8E) : SetPixelV(lhdc, tmpw + 24, tmph + 16, &HE7BC8C) : SetPixelV(lhdc, tmpw + 25, tmph + 16, &HE7BB8A) : SetPixelV(lhdc, tmpw + 26, tmph + 16, &HE5B37F) : SetPixelV(lhdc, tmpw + 27, tmph + 16, &HE1A971) : SetPixelV(lhdc, tmpw + 28, tmph + 16, &HD79F67) : SetPixelV(lhdc, tmpw + 29, tmph + 16, &H8E6A40) : SetPixelV(lhdc, tmpw + 30, tmph + 16, &H8A8076) : SetPixelV(lhdc, tmpw + 31, tmph + 16, &HE2E2E2) : SetPixelV(lhdc, tmpw + 32, tmph + 16, &HF9F9F9)
		SetPixelV(lhdc, tmpw + 17, tmph + 17, &HE2CC9D) : SetPixelV(lhdc, tmpw + 18, tmph + 17, &HE2CC9E) : SetPixelV(lhdc, tmpw + 19, tmph + 17, &HE3CF9C) : SetPixelV(lhdc, tmpw + 20, tmph + 17, &HDFCA98) : SetPixelV(lhdc, tmpw + 21, tmph + 17, &HE2CD9C) : SetPixelV(lhdc, tmpw + 22, tmph + 17, &HE4CC9C) : SetPixelV(lhdc, tmpw + 23, tmph + 17, &HE1C491) : SetPixelV(lhdc, tmpw + 24, tmph + 17, &HE1C18F) : SetPixelV(lhdc, tmpw + 25, tmph + 17, &HE4BD8C) : SetPixelV(lhdc, tmpw + 26, tmph + 17, &HDAB285) : SetPixelV(lhdc, tmpw + 27, tmph + 17, &HC7A582) : SetPixelV(lhdc, tmpw + 28, tmph + 17, &H806A56) : SetPixelV(lhdc, tmpw + 29, tmph + 17, &H676565) : SetPixelV(lhdc, tmpw + 30, tmph + 17, &HD2D2D2) : SetPixelV(lhdc, tmpw + 31, tmph + 17, &HF2F2F2) : SetPixelV(lhdc, tmpw + 32, tmph + 17, &HFEFEFE)
		SetPixelV(lhdc, tmpw + 17, tmph + 18, &HE9D5A6) : SetPixelV(lhdc, tmpw + 18, tmph + 18, &HE9D5A6) : SetPixelV(lhdc, tmpw + 19, tmph + 18, &HE9D6A3) : SetPixelV(lhdc, tmpw + 20, tmph + 18, &HE9D7A6) : SetPixelV(lhdc, tmpw + 21, tmph + 18, &HE2CC9B) : SetPixelV(lhdc, tmpw + 22, tmph + 18, &HE8D0A0) : SetPixelV(lhdc, tmpw + 23, tmph + 18, &HE8CF9C) : SetPixelV(lhdc, tmpw + 24, tmph + 18, &HE7C795) : SetPixelV(lhdc, tmpw + 25, tmph + 18, &HE2BC8B) : SetPixelV(lhdc, tmpw + 26, tmph + 18, &HB68F63) : SetPixelV(lhdc, tmpw + 27, tmph + 18, &H886948) : SetPixelV(lhdc, tmpw + 28, tmph + 18, &H786A5D) : SetPixelV(lhdc, tmpw + 29, tmph + 18, &HC7C7C7) : SetPixelV(lhdc, tmpw + 30, tmph + 18, &HEBEBEB) : SetPixelV(lhdc, tmpw + 31, tmph + 18, &HFCFCFC)
		SetPixelV(lhdc, tmpw + 17, tmph + 19, &HD7D2BA) : SetPixelV(lhdc, tmpw + 18, tmph + 19, &HD7D2BA) : SetPixelV(lhdc, tmpw + 19, tmph + 19, &HD7D1B9) : SetPixelV(lhdc, tmpw + 20, tmph + 19, &HD5CEB6) : SetPixelV(lhdc, tmpw + 21, tmph + 19, &HDBD3BB) : SetPixelV(lhdc, tmpw + 22, tmph + 19, &HC9C1AA) : SetPixelV(lhdc, tmpw + 23, tmph + 19, &HA9A28B) : SetPixelV(lhdc, tmpw + 24, tmph + 19, &H827E6C) : SetPixelV(lhdc, tmpw + 25, tmph + 19, &H6A665B) : SetPixelV(lhdc, tmpw + 26, tmph + 19, &H625F5A) : SetPixelV(lhdc, tmpw + 27, tmph + 19, &H8B8C8C) : SetPixelV(lhdc, tmpw + 28, tmph + 19, &HCDCDCD) : SetPixelV(lhdc, tmpw + 29, tmph + 19, &HE8E8E8) : SetPixelV(lhdc, tmpw + 30, tmph + 19, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 20, &H5A563E) : SetPixelV(lhdc, tmpw + 18, tmph + 20, &H59543D) : SetPixelV(lhdc, tmpw + 19, tmph + 20, &H58513A) : SetPixelV(lhdc, tmpw + 20, tmph + 20, &H554F38) : SetPixelV(lhdc, tmpw + 21, tmph + 20, &H59513B) : SetPixelV(lhdc, tmpw + 22, tmph + 20, &H58513E) : SetPixelV(lhdc, tmpw + 23, tmph + 20, &H646053) : SetPixelV(lhdc, tmpw + 24, tmph + 20, &H7B7973) : SetPixelV(lhdc, tmpw + 25, tmph + 20, &HA2A19F) : SetPixelV(lhdc, tmpw + 26, tmph + 20, &HC5C5C5) : SetPixelV(lhdc, tmpw + 27, tmph + 20, &HDADADA) : SetPixelV(lhdc, tmpw + 28, tmph + 20, &HEDEDED) : SetPixelV(lhdc, tmpw + 29, tmph + 20, &HFAFAFA)
		SetPixelV(lhdc, tmpw + 17, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 18, tmph + 21, &HC5C5C5) : SetPixelV(lhdc, tmpw + 19, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 20, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 21, tmph + 21, &HC6C6C6) : SetPixelV(lhdc, tmpw + 22, tmph + 21, &HC9C9C9) : SetPixelV(lhdc, tmpw + 23, tmph + 21, &HCECECE) : SetPixelV(lhdc, tmpw + 24, tmph + 21, &HD7D7D7) : SetPixelV(lhdc, tmpw + 25, tmph + 21, &HE1E1E1) : SetPixelV(lhdc, tmpw + 26, tmph + 21, &HECECEC) : SetPixelV(lhdc, tmpw + 27, tmph + 21, &HF6F6F6) : SetPixelV(lhdc, tmpw + 28, tmph + 21, &HFDFDFD)
		SetPixelV(lhdc, tmpw + 17, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 18, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 19, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 20, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 21, tmph + 22, &HECECEC) : SetPixelV(lhdc, tmpw + 22, tmph + 22, &HEDEDED) : SetPixelV(lhdc, tmpw + 23, tmph + 22, &HF0F0F0) : SetPixelV(lhdc, tmpw + 24, tmph + 22, &HF4F4F4) : SetPixelV(lhdc, tmpw + 25, tmph + 22, &HFAFAFA) : SetPixelV(lhdc, tmpw + 26, tmph + 22, &HFDFDFD)
		tmph = 11 : tmph1 = lh - 10 : tmpw = lw - 34
		'Generar lineas intermedias
		APILine(0, tmph, 0, tmph1, &HF7F7F7) : APILine(1, tmph, 1, tmph1, &HA99D9B) : APILine(2, tmph, 2, tmph1, &H632D17) : APILine(3, tmph, 3, tmph1, &HA65A1D) : APILine(4, tmph, 4, tmph1, &HB96C2E)
		APILine(5, tmph, 5, tmph1, &HBC7738) : APILine(6, tmph, 6, tmph1, &HC18242) : APILine(7, tmph, 7, tmph1, &HC2894E) : APILine(8, tmph, 8, tmph1, &HC18A52) : APILine(9, tmph, 9, tmph1, &HC59157)
		APILine(10, tmph, 10, tmph1, &HC59159) : APILine(11, tmph, 11, tmph1, &HCC9863) : APILine(12, tmph, 12, tmph1, &HCC9665) : APILine(13, tmph, 13, tmph1, &HCB9767) : APILine(14, tmph, 14, tmph1, &HC99565)
		APILine(15, tmph, 15, tmph1, &HCC9A6A) : APILine(16, tmph, 16, tmph1, &HCC9A6A) : APILine(17, tmph, 17, tmph1, &HCC9B6A) : APILine(tmpw + 17, tmph, tmpw + 17, tmph1, &HCC9B6A) : APILine(tmpw + 18, tmph, tmpw + 18, tmph1, &HCC9B6A)
		APILine(tmpw + 19, tmph, tmpw + 19, tmph1, &HCA9864) : APILine(tmpw + 20, tmph, tmpw + 20, tmph1, &HC99763) : APILine(tmpw + 21, tmph, tmpw + 21, tmph1, &HC99763) : APILine(tmpw + 22, tmph, tmpw + 22, tmph1, &HCA9562) : APILine(tmpw + 23, tmph, tmpw + 23, tmph1, &HC9945D)
		APILine(tmpw + 24, tmph, tmpw + 24, tmph1, &HC68E57) : APILine(tmpw + 25, tmph, tmpw + 25, tmph1, &HC48C55) : APILine(tmpw + 26, tmph, tmpw + 26, tmph1, &HC3884E) : APILine(tmpw + 27, tmph, tmpw + 27, tmph1, &HBE823E) : APILine(tmpw + 28, tmph, tmpw + 28, tmph1, &HB97634)
		APILine(tmpw + 29, tmph, tmpw + 29, tmph1, &HBD6D2D) : APILine(tmpw + 30, tmph, tmpw + 30, tmph1, &HA8581C) : APILine(tmpw + 31, tmph, tmpw + 31, tmph1, &H6D3319) : APILine(tmpw + 32, tmph, tmpw + 32, tmph1, &HA49794) : APILine(tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6)
		'Lineas verticales
		APILine(17, 0, lw - 17, 0, &H3C090A)
		APILine(17, 1, lw - 17, 1, &HDEC7BE)
		APILine(17, 2, lw - 17, 2, &HD3BCB2)
		APILine(17, 3, lw - 17, 3, &HD4B39A)
		APILine(17, 4, lw - 17, 4, &HCCAC92)
		APILine(17, 5, lw - 17, 5, &HCFA88A)
		APILine(17, 6, lw - 17, 6, &HCCA587)
		APILine(17, 7, lw - 17, 7, &HD2A77D)
		APILine(17, 8, lw - 17, 8, &HB98E62)
		APILine(17, 9, lw - 17, 9, &HC69464)
		APILine(17, 10, lw - 17, 10, &HCC9B6A)
		APILine(17, 11, lw - 17, 11, &HD0A175)
		tmph = lh - 22
		APILine(17, tmph + 11, lw - 17, tmph + 11, &HD0A175)
		APILine(17, tmph + 12, lw - 17, tmph + 12, &HDAAA7F)
		APILine(17, tmph + 13, lw - 17, tmph + 13, &HDCB087)
		APILine(17, tmph + 14, lw - 17, tmph + 14, &HE4B88E)
		APILine(17, tmph + 15, lw - 17, tmph + 15, &HE1BF93)
		APILine(17, tmph + 16, lw - 17, tmph + 16, &HE7C699)
		APILine(17, tmph + 17, lw - 17, tmph + 17, &HE2CC9D)
		APILine(17, tmph + 18, lw - 17, tmph + 18, &HE9D5A6)
		APILine(17, tmph + 19, lw - 17, tmph + 19, &HD7D2BA)
		APILine(17, tmph + 20, lw - 17, tmph + 20, &H5A563E)
		APILine(17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5)
		APILine(17, tmph + 22, lw - 17, tmph + 22, &HECECEC)
		Exit Sub
		
DrawMacOSXButtonPressed_Error: 
	End Sub
	
	Private Sub DrawPlastikButton(ByRef iState As isState)
		'On Error GoTo DrawPlastikButton_Error
		
		Dim tmpcolor As Integer
		
		'If Ambient.DisplayAsDefault Then iState = stateDefaulted
		Select Case iState
			Case isState.statenormal, isState.stateDefaulted
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H8s))
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), 1, 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), 1, 2, 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H60s))
				APILine(2, 0, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				APILine(MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'Border Pixels
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H15s))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H15s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H25s)))
				
				If iState = isState.stateDefaulted Or m_bFocused Then
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
					APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)))
					APILine(1, 2, MyBase.ClientRectangle.Width - 1, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)))
					APILine(1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)))
					APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)))
				End If
				
				Exit Sub
				
			Case isState.stateHot
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H18s))
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H10s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H10s)), 1, 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), 1, 2, 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H60s))
				APILine(2, 0, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				APILine(MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'Border Pixels
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H15s))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE)), &H15s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE)), -&H10s)))
				tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)))
				APILine(1, 2, MyBase.ClientRectangle.Width - 1, 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)))
				APILine(1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)))
				Exit Sub
				
			Case isState.statePressed
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H9s))
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), tmpcolor, 1, 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)), MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), 1, 2, 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H60s))
				APILine(2, 0, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				APILine(MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'Border Pixels
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H15s))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H8s)))
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H15s)))
				Exit Sub
				
			Case isState.statedisabled
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H12s))
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), 1, 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), 1, 2, 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5)
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H60s))
				APILine(2, 0, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				APILine(MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'Border Pixels
				tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H5s))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, 0, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, 1, tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
				APILine(2, 1, MyBase.ClientRectangle.Width - 2, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &H15s)))
				APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&H25s)))
				'If iState = stateDefaulted Or m_bFocused Then
				'    tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
				'    APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
				'    APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
				'    APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
				'    APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
				'End If
				Exit Sub
				
		End Select
		
		Exit Sub
		
DrawPlastikButton_Error: 
	End Sub
	
	Private Sub DrawGalaxyButton(ByRef iState As isState)
		'On Error GoTo DrawGalaxyButton_Error
		
		Dim tmpcolor As Integer
		
		If iState = isState.statenormal Then
			tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
		Else
			tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, OffsetColor(System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE)), &HFs))
		End If
		
		MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
		
		If iState = isState.statePressed Then
			DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 6)
			APILine(2, 1, MyBase.ClientRectangle.Width - 3, 1, tmpcolor)
			APILine(1, 2, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
			APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H60s)))
			APILine(MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H60s)))
		Else
			DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), 2, 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 6)
			APILine(2, 1, MyBase.ClientRectangle.Width - 3, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H60s)))
			APILine(1, 2, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H60s)))
			APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
			APILine(MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		End If
		
		tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
		
		If iState <> isState.statePressed Then
			tmpcolor = IIf(m_bFocused, OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H60s), OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s))
		Else
			tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s))
		End If
		
		APILine(2, 0, MyBase.ClientRectangle.Width - 3, 0, tmpcolor)
		APILine(0, 2, 0, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, tmpcolor)
		APILine(MyBase.ClientRectangle.Width - 2, 2, MyBase.ClientRectangle.Width - 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, 1, 1, tmpcolor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 3, 1, tmpcolor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H15s))
		APILine(3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, tmpcolor)
		APILine(MyBase.ClientRectangle.Width - 1, 3, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, tmpcolor)
		APILine(MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5, tmpcolor)
		APILine(MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)))
		Exit Sub
		
DrawGalaxyButton_Error: 
	End Sub
	
	Private Sub DrawKeramikButton(ByRef iState As isState)
		'On Error GoTo DrawKeramikButton_Error
		
		Dim tmpcolor As Integer
		
		If m_iState = isState.statenormal Then
			tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), -&HFs))
		ElseIf m_iState = isState.stateHot Then 
			tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))), &H1s))
		ElseIf m_iState = isState.statePressed Then 
			tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))), &H18s))
		Else
			tmpcolor = System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))), &HFs))
		End If
		
		MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H20s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 5, 2, MyBase.ClientRectangle.Width - 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 6)
		'Left
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H80s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H80s)), 0, 0, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), 2, 4, 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 7)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H35s)), 3, 3, 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 7)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), 4, 2, 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 7)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H5s)), 5, 2, 6, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 2, 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 5)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H35s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 3, 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 3)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4, 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 1)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4, 6, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 3)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H35s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 5, 12, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 6)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H38s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 6, 7, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 7)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 7, 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 8)
		'Right
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H80s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H80s)), MyBase.ClientRectangle.Width - 2, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H80s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)), MyBase.ClientRectangle.Width - 1, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H40s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H8s)), MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), MyBase.ClientRectangle.Width - 4, 4, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 5)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)), MyBase.ClientRectangle.Width - 5, 3, MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawVGradient(tmpcolor, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), MyBase.ClientRectangle.Width - 6, 2, MyBase.ClientRectangle.Width - 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 3)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H5s)), MyBase.ClientRectangle.Width - 7, MyBase.ClientRectangle.Width - 6, 6, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)), MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 5)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H30s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), MyBase.ClientRectangle.Width - 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2, MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 5)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H35s)), MyBase.ClientRectangle.Width - 6, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2, MyBase.ClientRectangle.Width - 5, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawVGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H5s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H25s)), MyBase.ClientRectangle.Width - 7, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2, MyBase.ClientRectangle.Width - 6, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) / 2 - 4)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H20s)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H35s)), MyBase.ClientRectangle.Width - 15, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, MyBase.ClientRectangle.Width - 7, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3)
		'top
		APILine(3, 0, MyBase.ClientRectangle.Width - 3, 0, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H30s)))
		APILine(1, 1, MyBase.ClientRectangle.Width - 1, 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H30s)))
		APILine(5, 1, MyBase.ClientRectangle.Width - 5, 1, tmpcolor) 'OffsetColor(tmpcolor, &H30)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H20s)), tmpcolor, MyBase.ClientRectangle.Width - 11, 2, MyBase.ClientRectangle.Width - 4, 3)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H20s)), tmpcolor, MyBase.ClientRectangle.Width - 10, 3, MyBase.ClientRectangle.Width - 5, 4)
		DrawHGradient(System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H20s)), tmpcolor, MyBase.ClientRectangle.Width - 9, 4, MyBase.ClientRectangle.Width - 6, 5)
		APILine(6, 3, 7, 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &H80s)))
		'bottom
		APILine(3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H10s)))
		APILine(2, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 2, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H80s)))
		APILine(7, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, MyBase.ClientRectangle.Width - 7, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, tmpcolor)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H70s)))
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixelV(MyBase.hdc, MyBase.ClientRectangle.Width - 3, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 3, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H70s)))
		APILine(MyBase.ClientRectangle.Width - 4, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 4, System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&H15s)))
		Exit Sub
		
DrawKeramikButton_Error: 
	End Sub
	
	'Important. If not included, tooltips don't change when you try to set the toltip text
	Private Sub RemoveToolTip()
		'On Error GoTo RemoveToolTip_Error
		
		Dim lR As Integer
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyBase.Extender.ToolTipText = m_sToolTipText
		
		If m_lttHwnd <> 0 Then
			'      With ttip
			'         .lSize = Len(ttip)
			'         .lHwnd = UserControl.hwnd
			'      End With
			'UPGRADE_WARNING: Couldn't resolve default property of object ttip. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lR = SendMessage(ttip.lHwnd, TTM_DELTOOLA, 0, ttip)
			DestroyWindow(m_lttHwnd)
		End If
		
		Exit Sub
		
RemoveToolTip_Error: 
	End Sub
	
	Private Sub CreateToolTip()
		'On Error GoTo CreateToolTip_Error
		
		Dim lpRect As RECT
		Dim lWinStyle As Integer
		'RemoveToolTip
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sToolTipText = MyBase.Extender.ToolTipText
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyBase.Extender.ToolTipText = ""
		ttip.lpStr = m_sToolTipText
		lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
		
		''create baloon style if desired
		If m_lToolTipType = ttStyleEnum.TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
		m_lttHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, MyBase.Handle.ToInt32, 0, VB6.GetHInstance.ToInt32, 0)
		''make our tooltip window a topmost window
		SetWindowPos(m_lttHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
		''get the rect of the parent control
		GetClientRect(MyBase.Handle.ToInt32, lpRect)
		
		''now set our tooltip info structure
		With ttip
			
			''if we want it centered, then set that flag
			If m_lttCentered Then
				.lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
			Else
				.lFlags = TTF_SUBCLASS
			End If
			
			''set the hwnd prop to our parent control's hwnd
			.lHwnd = MyBase.Handle.ToInt32
			.lId = 0
			.hInstance = VB6.GetHInstance.ToInt32
			'.lpstr = ALREADY SET
			'UPGRADE_WARNING: Couldn't resolve default property of object ttip.lpRect. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.lpRect = lpRect
		End With
		
		''add the tooltip structure
		'UPGRADE_WARNING: Couldn't resolve default property of object ttip. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SendMessage(m_lttHwnd, TTM_ADDTOOLA, 0, ttip)
		
		''if we want a title or we want an icon
		If m_sTooltiptitle <> vbNullString Or m_lToolTipIcon <> ttIconType.TTNoIcon Then
			SendMessage(m_lttHwnd, TTM_SETTITLE, CInt(m_lToolTipIcon), m_sTooltiptitle)
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not IsNothing(m_lttForeColor) Then
			SendMessage(m_lttHwnd, TTM_SETTIPTEXTCOLOR, TranslateColor(m_lttForeColor), 0)
		End If
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not IsNothing(m_lttBackColor) Then
			SendMessage(m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(m_lttBackColor), 0)
		End If
		
		Exit Sub
		
CreateToolTip_Error: 
	End Sub
	
	Private Sub m_About_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles m_About.Click
		'On Error GoTo m_About_Click_Error
		
		m_About.Visible = False
		Exit Sub
		
m_About_Click_Error: 
	End Sub
	
	Private Sub m_About_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles m_About.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'On Error GoTo m_About_MouseDown_Error
		
		Dim tmpRect As RECT
		Dim tmpcolor As Integer
		
		With m_About
			'Draw button
			SetRect(tmpRect, 290, 80, 360, 26)
			'tmpcolor = OffsetColor(GetSysColor(COLOR_BTNFACE), &HF)
			tmpcolor = GetSysColor(COLOR_BTNFACE)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawVGradientEx((m_About.hdc), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), tmpRect.Left_Renamed, tmpRect.Top_Renamed, tmpRect.Right_Renamed, tmpRect.Bottom_Renamed)
			tmpcolor = GetSysColor(COLOR_BTNSHADOW)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 2, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 2, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpRect.Right_Renamed - 3, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			tmpcolor = GetSysColor(COLOR_BTNFACE)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed + 1, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed - 1, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed + 1, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed - 1, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 1, tmpcolor)
			SetRect(tmpRect, 290, 80, 360, 106)
			.Font = VB6.FontChangeSize(.Font, 8)
			.ForeColor = System.Drawing.Color.Blue
			.Font = VB6.FontChangeUnderline(.Font, True)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "Close", -1, tmpRect, DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_SINGLELINE)
		End With
		
		Exit Sub
		
m_About_MouseDown_Error: 
	End Sub
	
	Private Sub m_About_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles m_About.Paint
		'Draw the About content
		'On Error GoTo m_About_Paint_Error
		
		Dim lwformat As Integer
		Dim tmpRect As RECT
		Dim tmpcolor As Integer
		lwformat = DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_LEFT Or DrawTextFlags.DT_SINGLELINE
		
		With m_About
			.ForeColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNTEXT))
			.Font = VB6.FontChangeUnderline(.Font, False)
			.Font = VB6.FontChangeSize(.Font, 18)
			SetRect(tmpRect, 20, 10, 220, 40)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "isButton", -1, tmpRect, lwformat)
			.Font = VB6.FontChangeSize(.Font, 10)
			SetRect(tmpRect, 160, 70, 300, 20)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "Version " & strCurrentVersion, -1, tmpRect, lwformat)
			.Font = VB6.FontChangeBold(.Font, True)
			SetRect(tmpRect, 20, 110, 250, 20)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "By Fred.cpp", -1, tmpRect, lwformat)
			.Font = VB6.FontChangeBold(.Font, False)
			SetRect(tmpRect, 20, 140, 250, 20)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "http://mx.geocities.com/fred_cpp/", -1, tmpRect, lwformat)
			'Draw button
			SetRect(tmpRect, 290, 80, 360, 26)
			'tmpcolor = OffsetColor(GetSysColor(COLOR_BTNFACE), &HF)
			tmpcolor = GetSysColor(COLOR_BTNFACE)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawVGradientEx((m_About.hdc), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), &HFs)), System.Drawing.ColorTranslator.ToOle(OffsetColor(System.Drawing.ColorTranslator.FromOle(tmpcolor), -&HFs)), tmpRect.Left_Renamed, tmpRect.Top_Renamed, tmpRect.Right_Renamed, tmpRect.Bottom_Renamed)
			tmpcolor = GetSysColor(COLOR_BTNSHADOW)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 2, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 2, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 2, tmpRect.Right_Renamed - 3, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			APILineEx(.hdc, tmpRect.Left_Renamed + 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpRect.Right_Renamed - 2, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			tmpcolor = GetSysColor(COLOR_BTNFACE)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed + 1, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed - 1, tmpRect.Top_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed + 1, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Left_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 1, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed - 1, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed, tmpcolor)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			SetPixelV(.hdc, tmpRect.Right_Renamed, tmpRect.Top_Renamed + tmpRect.Bottom_Renamed - 1, tmpcolor)
			SetRect(tmpRect, 290, 80, 360, 106)
			.Font = VB6.FontChangeSize(.Font, 8)
			.ForeColor = System.Drawing.Color.Blue
			.Font = VB6.FontChangeUnderline(.Font, True)
			'UPGRADE_ISSUE: PictureBox property m_About.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(.hdc, "Close", -1, tmpRect, DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_SINGLELINE)
		End With
		
		Exit Sub
		
m_About_Paint_Error: 
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.AccessKeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_AccessKeyPress(ByRef KeyAscii As Short)
		'When action called by Accesskey
		'On Error GoTo UserControl_AccessKeyPress_Error
		
		lPrevButton = VB6.MouseButtonConstants.LeftButton
		ctlButton_Click(Me, New System.EventArgs())
		m_iState = isState.statenormal
		Exit Sub
		
UserControl_AccessKeyPress_Error: 
	End Sub
	
	'Manage default and Cancel events
	'Looks like each time the control get's the focus, the Default
	' property is also Set, It's kinda annoying I still am Trying
	' to figure out why and how can I Implement Default And Cancel
	' properties
	'UPGRADE_ISSUE: UserControl event UserControl.AmbientChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_AmbientChanged(ByRef PropertyName As String)
		'On Error GoTo UserControl_AmbientChanged_Error
		
		Select Case PropertyName
			Case "DisplayAsDefault"
				Refresh()
		End Select
		
		Exit Sub
		
UserControl_AmbientChanged_Error: 
	End Sub
	
	Private Sub ctlButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		'On Error GoTo UserControl_Click_Error
		
		If lPrevButton = VB6.MouseButtonConstants.LeftButton Then
			If m_ButtonType = isbButtonType.isbCheckBox Then
				m_Value = Not m_Value
			End If
			
			m_iState = isState.stateHot
			Refresh()
			RaiseEvent Click(Me, Nothing)
		End If
		
		Exit Sub
		
UserControl_Click_Error: 
	End Sub
	
	Private Sub ctlButton_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
		'On Error GoTo UserControl_DblClick_Error
		
		If lPrevButton = VB6.MouseButtonConstants.LeftButton Then
			ctlButton_MouseDown(Me, New System.Windows.Forms.MouseEventArgs(1 * &H100000, 0, VB6.TwipsToPixelsX(1), VB6.TwipsToPixelsY(1), 0))
		End If
		
		Exit Sub
		
UserControl_DblClick_Error: 
	End Sub
	
	Private Sub ctlButton_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Enter
		'On Error GoTo UserControl_EnterFocus_Error
		
		m_bFocused = True
		Refresh()
		Exit Sub
		
UserControl_EnterFocus_Error: 
	End Sub
	
	Private Sub ctlButton_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Leave
		'On Error GoTo UserControl_ExitFocus_Error
		
		m_bFocused = False
		Refresh()
		Exit Sub
		
UserControl_ExitFocus_Error: 
	End Sub
	
	Private Sub ctlButton_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		'On Error GoTo UserControl_GotFocus_Error
		
		m_bFocused = True
		Refresh()
		Exit Sub
		
UserControl_GotFocus_Error: 
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.Hide was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_Hide()
		'On Error GoTo UserControl_Hide_Error
		
		m_bVisible = False
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyBase.Extender.ToolTipText = m_sToolTipText
		Exit Sub
		
UserControl_Hide_Error: 
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_InitProperties()
		'On Error GoTo UserControl_InitProperties_Error
		
		m_iStyle = 0
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.Name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_sCaption = MyBase.Extender.Name
		m_IconSize = 16
		m_Icon = Nothing
		lwFontAlign = DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_WORDBREAK 'DT_VCENTER Or DT_CENTER
		m_bEnabled = True
		m_bShowFocus = False
		m_bUseCustomColors = False
		m_lBackColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Control))
		m_lHighlightColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Highlight))
		m_lFontColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.ControlText))
		m_lFontHighlightColor = TranslateColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.ControlText))
		lPrevStyle = GetWindowLong(m_About.Handle.ToInt32, GWL_STYLE)
		m_lToolTipType = ttStyleEnum.TTBalloon
		m_CaptionAlign = isbAlign.isbCenter
		m_IconAlign = isbAlign.isbleft
		iStyleIconOffset = 4
		Exit Sub
		
UserControl_InitProperties_Error: 
	End Sub
	
	Private Sub ctlButton_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'On Error GoTo UserControl_KeyDown_Error
		
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Space
				ctlButton_Click(Me, New System.EventArgs())
				RaiseEvent Click(Me, Nothing)
			Case System.Windows.Forms.Keys.Right, System.Windows.Forms.Keys.Down
				System.Windows.Forms.SendKeys.Send("{TAB}")
			Case System.Windows.Forms.Keys.Left, System.Windows.Forms.Keys.Up
				System.Windows.Forms.SendKeys.Send("+{TAB}")
		End Select
		
		Exit Sub
		
UserControl_KeyDown_Error: 
	End Sub
	
	Private Sub ctlButton_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		'On Error GoTo UserControl_LostFocus_Error
		
		m_bFocused = False
		Refresh()
		Exit Sub
		
UserControl_LostFocus_Error: 
	End Sub
	
	Private Sub ctlButton_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'On Error GoTo UserControl_MouseDown_Error
		
		If Button = VB6.MouseButtonConstants.LeftButton Then
			m_iState = isState.statePressed
			Refresh()
		End If
		
		Exit Sub
		
UserControl_MouseDown_Error: 
	End Sub
	
	' Description: Refresh the control
	Public Overrides Sub Refresh()
		Dim tmpcolor As Integer
		Dim lTransColor As Integer
		Dim lcurrpix As Integer
		
		'UPGRADE_ISSUE: UserControl method UserControl.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		MyBase.Cls()
		iStyleIconOffset = 3
		
		If Not m_bVisible Then Exit Sub
		If DesignMode And m_iStyle <> isbStyle.isbWindowsXP Then m_iState = isState.stateHot
		If m_ButtonType = isbButtonType.isbCheckBox Then
			If m_Value Then m_iState = isState.statePressed
		End If
		
		If Not m_bEnabled Then
			m_iState = isState.statedisabled
			MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE))
		End If
		
		Dim tmpRect As RECT
		Dim bDrawThemeSuccess As Boolean
		Dim tmpStyle As isbStyle
		Select Case m_iStyle
			Case isbStyle.isbNormal
				
				'Classic Style (Win98)
				If m_iState = isState.statenormal Then
					tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
				Else
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
				End If
				
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawCtlEdge(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), IIf(m_iState = isState.statePressed, EDGE_SUNKEN, EDGE_RAISED))
				
				'If Ambient.DisplayAsDefault Then
				'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayAsDefault was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If (m_bFocused And m_bShowFocus) Or (Ambient.DisplayAsDefault And Not m_bFocused) Then
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.Color.Black)
				End If
				
				'End If
			Case isbStyle.isbSoft
				
				'Soft Style (I don't know where does It come, But I've seen this before)
				If m_iState = isState.statenormal Or m_iState = isState.statedisabled Then
					tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
				Else
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
				End If
				
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				
				Select Case m_iState
					Case isState.statenormal, isState.stateHot, isState.stateDefaulted
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawCtlEdge(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), BDR_RAISEDINNER)
					Case isState.statePressed
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawCtlEdge(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), BDR_SUNKENOUTER)
					Case isState.statedisabled
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.FromOle(tmpcolor))
				End Select
				
				'If Ambient.DisplayAsDefault Then
				'    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
				'End If
			Case isbStyle.isbFlat
				
				'Flat Style (Office 2000 like)
				If m_iState = isState.statenormal Or m_iState = isState.statedisabled Then
					tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
				Else
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
				End If
				
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				
				If m_iState = isState.statenormal Then
					'Normal (flat)
					tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
					MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.FromOle(tmpcolor))
				ElseIf m_iState = isState.stateHot Then 
					'Hover
					'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
					'UserControl.BackColor = tmpColor
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawCtlEdge(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), BDR_RAISEDINNER)
				ElseIf m_iState = isState.statePressed Then 
					'Pushed
					'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
					'UserControl.BackColor = tmpColor
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawCtlEdge(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), BDR_SUNKENOUTER)
				Else 'Disabled
					'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
					'UserControl.BackColor = tmpColor
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.FromOle(tmpcolor))
				End If
				
				'If Ambient.DisplayAsDefault Then
				'    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
				'End If
			Case isbStyle.isbJava
				'Java Style
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)))
				
				Select Case m_iState
					Case isState.statePressed
						tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNSHADOW))
					Case isState.stateHot
						tmpcolor = IIf(m_bUseCustomColors, BlendColors(m_lHighlightColor, m_lBackColor), BlendColors(GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_BTNFACE)))
					Case Else
						tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
				End Select
				CopyRect(tmpRect, m_btnRect) : InflateRect(tmpRect, -4, -4)
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				APIFillRect(MyBase.hdc, m_btnRect, tmpcolor) 'm_txtRect
				DrawJavaBorder(m_btnRect.Left_Renamed, m_btnRect.Top_Renamed, m_btnRect.Right_Renamed - m_btnRect.Left_Renamed - 1, m_btnRect.Bottom_Renamed - m_btnRect.Top_Renamed - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), tmpcolor)
				'If Ambient.DisplayAsDefault Then
				'    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
				'End If
			Case isbStyle.isbOfficeXP
				'Redmond 2002 Office Suite ( ... )
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(tmpcolor)
				
				If m_iState = isState.statenormal Then
					'If Ambient.DisplayAsDefault Then
					'    tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
					'    APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor)
					'    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
					'End If
					tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20s)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIFillRectByCoords(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				ElseIf m_iState = isState.stateHot Then 
					'Hover
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIFillRectByCoords(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MSOXPShiftColor(tmpcolor))
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.FromOle(tmpcolor))
				ElseIf m_iState = isState.statePressed Then 
					'Pushed
					tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIFillRectByCoords(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, MSOXPShiftColor(tmpcolor, &H80s))
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIRectangle(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width - 1, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - 1, System.Drawing.ColorTranslator.FromOle(tmpcolor))
				Else
					'Disabled
					tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20s)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					APIFillRectByCoords(MyBase.hdc, 0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height), tmpcolor)
				End If
				
			Case isbStyle.isbWindowsXP
				
				'WinXP (Emulated)
				If m_bUseCustomColors Then
					DrawCustomWinXPButton(m_iState)
				Else
					DrawWinXPButton(m_iState)
				End If
				
				iStyleIconOffset = 3
			Case isbStyle.isbWindowsTheme
				'Uses the current installed windows theme
				MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE))
				
				If m_iState = (isState.statenormal And m_bFocused) Then 'Or Ambient.DisplayAsDefault Then
					bDrawThemeSuccess = DrawTheme("Button", 1, isState.stateDefaulted)
				Else
					bDrawThemeSuccess = DrawTheme("Button", 1, m_iState)
				End If
				
				If Not bDrawThemeSuccess Then
					m_iStyle = Me.NonThemeStyle
				End If
				
			Case isbStyle.isbPlastik
				DrawPlastikButton(m_iState)
				iStyleIconOffset = 4
			Case isbStyle.isbGalaxy
				DrawGalaxyButton(m_iState)
				iStyleIconOffset = 4
			Case isbStyle.isbKeramik
				DrawKeramikButton(m_iState)
				iStyleIconOffset = 6
			Case isbStyle.isbMacOSX
				'¿?¿??¿?Yes! Do you like It?
				DrawMacOSXButton()
				iStyleIconOffset = 7
		End Select
		
		''''''Draw Icon
		Dim ix, iy As Integer
		Dim ni, nj As Integer
		If Not m_Icon Is Nothing Then
			If CDbl(CObj(Icon)) <> 0 Then
				If m_IconAlign = isbAlign.isbCenter Then
					ix = (MyBase.ClientRectangle.Width - m_IconSize) / 2
					iy = (VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - m_IconSize) / 2
				ElseIf m_IconAlign = isbAlign.isbbottom Then 
					ix = (MyBase.ClientRectangle.Width - m_IconSize) / 2
					iy = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - m_IconSize - iStyleIconOffset
				ElseIf m_IconAlign = isbAlign.isbTop Then 
					ix = (MyBase.ClientRectangle.Width - m_IconSize) / 2
					iy = iStyleIconOffset
				ElseIf m_IconAlign = isbAlign.isbleft Then 
					ix = iStyleIconOffset
					iy = (VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - m_IconSize) / 2
				ElseIf m_IconAlign = isbAlign.isbRight Then 
					ix = MyBase.ClientRectangle.Width - m_IconSize - iStyleIconOffset
					iy = (VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height) - m_IconSize) / 2
				End If
				
				If m_iState = isState.statePressed Then
					ix = ix + 1
					iy = iy + 1
				ElseIf m_iState = isState.stateHot Then 
					If m_iStyle = isbStyle.isbOfficeXP Then
						If m_UseMaskColor Then
							'This was added By t_eee eeee
							'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							TransBlt(MyBase.hdc, ix + 1, iy + 1, m_IconSize, m_IconSize, m_Icon, m_MaskColor, &H808080)
							'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							TransBlt(MyBase.hdc, ix - 1, iy - 1, m_IconSize, m_IconSize, m_Icon, m_MaskColor) '                        pMask.PaintPicture m_Icon, ix, iy, m_IconSize, m_IconSize, , , , , vbSrcCopy
						Else
							
							'UPGRADE_ISSUE: UserControl method ctlButton.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							PaintPicture(m_Icon, ix, iy, m_IconSize, m_IconSize)
							'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							lTransColor = GetPixel(MyBase.hdc, 1, 1)
							For nj = iy To iy + m_IconSize
								For ni = ix To ix + m_IconSize
									'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									lcurrpix = GetPixel(MyBase.hdc, ni, nj)
									If lcurrpix <> lTransColor Then
										If m_UseMaskColor Then
											If lcurrpix <> m_MaskColor Then
												'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
												SetPixelV(MyBase.hdc, ni, nj, &H808080)
											End If
										Else
											'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
											SetPixelV(MyBase.hdc, ni, nj, &H808080)
										End If
									End If
								Next ni
							Next nj
						End If
						ix = ix - 2
						iy = iy - 2
					End If
				End If
				'I'll try to mask when usemaskcolor is true
				If m_UseMaskColor Then
					If m_bEnabled Then
						'Paint in the about pic on color
						'On Error GoTo MalformedIcon
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						TransBlt(MyBase.hdc, ix, iy, m_IconSize, m_IconSize, m_Icon, m_MaskColor)
					Else
						'Disabled
						'On Error GoTo MalformedIcon
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						TransBlt(MyBase.hdc, ix, iy, m_IconSize, m_IconSize, m_Icon, m_MaskColor,  ,  , True)
					End If
				Else
MalformedIcon: 
					If m_bEnabled Then
						'UPGRADE_ISSUE: UserControl method ctlButton.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PaintPicture(m_Icon, ix, iy, m_IconSize, m_IconSize)
					Else
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						PaintIconGrayscale(MyBase.hdc, CInt(CObj(m_Icon)), ix, iy, m_IconSize, m_IconSize)
					End If
				End If
			End If
			
		End If
		
		'DrawText
		DrawCaption()
		
	End Sub
	
	Private Function lGrayScale(ByRef coloredColor As Integer) As Integer
		'Convertir un Long RGB a un LongRGB grayscale
		'For this task Imported previously developed functions from my project Opticlops
		'' Desc: Convert a RGB color to long
		'Private Function RGBToLong(RGBColor As RGB) As Long
		'    RGBToLong = RGBColor.Blue + RGBColor.Green * 265 + RGBColor.Red * 65536
		'End Function
		'
		'' Desc Convert a long into a RGB structure
		'Private Function LongToRGB(lcolor As Long) As RGB
		'    LongToRGB.Red = lcolor And &HFF
		'    LongToRGB.Green = (lcolor \ &H100) And &HFF
		'    LongToRGB.Blue = (lcolor \ &H10000) And &HFF
		'End Function
		'On Error GoTo lGrayScale_Error
		
		Dim G, R, b As Integer
		Dim neutral As Integer
		'Splitt into RGB values
		b = coloredColor And &HFFs
		G = (coloredColor \ &H100s) And &HFFs
		R = (coloredColor \ &H10000) And &HFFs
		'Obtener el promedio
		neutral = (R / 3 + G / 3 + b / 3)
		'Build Long
		lGrayScale = RGB(neutral, neutral, neutral)
		Exit Function
		
lGrayScale_Error: 
	End Function
	
	Private Sub BuildRegion()
		'On Error GoTo BuildRegion_Error
		
		If m_lRegion Then DeleteObject(m_lRegion)
		
		Select Case m_iStyle
			Case isbStyle.isbMacOSX
				m_lRegion = CreateMacOSXButtonRegion
			Case isbStyle.isbWindowsXP, isbStyle.isbPlastik
				m_lRegion = CreateWinXPregion
			Case isbStyle.isbGalaxy, isbStyle.isbKeramik
				m_lRegion = CreateGalaxyRegion
			Case Else
				m_lRegion = CreateRectRgn(0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height))
		End Select
		
		SetWindowRgn(MyBase.Handle.ToInt32, m_lRegion, True)
		Exit Sub
		
BuildRegion_Error: 
	End Sub
	
	Private Sub ctlButton_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'On Error GoTo UserControl_MouseUp_Error
		
		If Button = VB6.MouseButtonConstants.LeftButton Then
			If m_ButtonType = isbButtonType.isbCheckBox And Not m_Value Then
				m_iState = isState.statePressed
			Else
				m_iState = isState.stateHot
			End If
			
			Refresh()
		End If
		
		lPrevButton = Button
		Exit Sub
		
UserControl_MouseUp_Error: 
	End Sub
	
	Private Sub ctlButton_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		'On Error GoTo UserControl_Paint_Error
		
		Call Refresh()
		Exit Sub
		
UserControl_Paint_Error: 
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.Show was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_Show()
		'm_sToolTipText = UserControl.Extender.ToolTipText
		'UserControl.Extender.ToolTipText = ""
		'On Error GoTo UserControl_Show_Error
		
		m_bVisible = True
		ctlButton_Resize(Me, New System.EventArgs())
		Refresh()
		Exit Sub
		
UserControl_Show_Error: 
	End Sub
	
	Private Sub ctlButton_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		Dim tmpRect As RECT
		Dim lh, lw As Integer
		'On Error Resume Next
		
		If VB6.PixelsToTwipsX(MyBase.Width) < 300 Then MyBase.Width = VB6.TwipsToPixelsX(300)
		If VB6.PixelsToTwipsY(MyBase.Height) < 300 Then MyBase.Height = VB6.TwipsToPixelsY(300)
		
		lh = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		lw = MyBase.ClientRectangle.Width
		SetRect(m_btnRect, 0, 0, lw, lh)
		SetRect(m_txtRect, 4, 4, lw - 4, lh - 4)
		CopyRect(tmpRect, m_txtRect)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		DrawText(MyBase.hdc, m_sCaption, Len(m_sCaption), tmpRect, DrawTextFlags.DT_CALCRECT Or DrawTextFlags.DT_WORDBREAK)
		
		Select Case m_CaptionAlign
			Case isbAlign.isbCenter
				'SetRect m_txtRect, (lw - tmpRect.Right - tmpRect.Left) / 2, (lh - tmpRect.bottom - tmpRect.Top) / 2, (lw + tmpRect.Right - tmpRect.Left) / 2, (lh + tmpRect.bottom - tmpRect.Top) / 2
				lwFontAlign = DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_WORDBREAK
			Case isbAlign.isbleft
				CopyRect(m_txtRect, tmpRect)
				SetRect(m_txtRect, iStyleIconOffset, (lh - tmpRect.Bottom_Renamed - tmpRect.Top_Renamed) / 2, tmpRect.Right_Renamed + iStyleIconOffset, (lh + tmpRect.Bottom_Renamed - tmpRect.Top_Renamed) / 2)
				lwFontAlign = DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_LEFT Or DrawTextFlags.DT_WORDBREAK
			Case isbAlign.isbRight
				CopyRect(m_txtRect, tmpRect)
				SetRect(m_txtRect, (lw - tmpRect.Right_Renamed - tmpRect.Left_Renamed) - iStyleIconOffset, (lh - tmpRect.Bottom_Renamed - tmpRect.Top_Renamed) / 2, (lw - tmpRect.Left_Renamed) - iStyleIconOffset, (lh + tmpRect.Bottom_Renamed - tmpRect.Top_Renamed) / 2)
				lwFontAlign = DrawTextFlags.DT_VCENTER Or DrawTextFlags.DT_RIGHT Or DrawTextFlags.DT_WORDBREAK
			Case isbAlign.isbTop
				CopyRect(m_txtRect, tmpRect)
				SetRect(m_txtRect, (lw - tmpRect.Right_Renamed - tmpRect.Left_Renamed) / 2, iStyleIconOffset / 2, (lw + tmpRect.Right_Renamed - tmpRect.Left_Renamed) / 2, iStyleIconOffset / 2 + (tmpRect.Bottom_Renamed - tmpRect.Top_Renamed))
				lwFontAlign = DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_TOP Or DrawTextFlags.DT_WORDBREAK
			Case isbAlign.isbbottom
				CopyRect(m_txtRect, tmpRect)
				SetRect(m_txtRect, (lw - tmpRect.Right_Renamed - tmpRect.Left_Renamed) / 2, lh - (tmpRect.Bottom_Renamed - tmpRect.Top_Renamed) - iStyleIconOffset / 2, (lw + tmpRect.Right_Renamed - tmpRect.Left_Renamed) / 2, lh - iStyleIconOffset / 2)
				lwFontAlign = DrawTextFlags.DT_CENTER Or DrawTextFlags.DT_BOTTOM Or DrawTextFlags.DT_WORDBREAK
		End Select
		
		BuildRegion()
		Refresh()
	End Sub
	
	'''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''
	'Properties
	'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As Object)
		'On Error GoTo UserControl_ReadProperties_Error
		
		m_iState = isState.statenormal
		
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			m_Icon = PropBag.ReadProperty("Icon", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iStyle = PropBag.ReadProperty("Style", 3)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sCaption = PropBag.ReadProperty("Caption", "isButton")
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_IconSize = PropBag.ReadProperty("IconSize", 16)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", isbAlign.isbCenter)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_IconAlign = PropBag.ReadProperty("IconAlign", isbAlign.isbleft)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_iNonThemeStyle = PropBag.ReadProperty("iNonThemeStyle", isbStyle.isbWindowsXP)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bEnabled = PropBag.ReadProperty("Enabled", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bShowFocus = PropBag.ReadProperty("ShowFocus", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bUseCustomColors = PropBag.ReadProperty("USeCustomColors", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lBackColor = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lHighlightColor = PropBag.ReadProperty("HighlightColor", GetSysColor(COLOR_HIGHLIGHT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lFontColor = PropBag.ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lFontHighlightColor = PropBag.ReadProperty("FontHighlightColor", GetSysColor(COLOR_BTNTEXT))
			'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MyBase.Extender.ToolTipText = PropBag.ReadProperty("ToolTipText", m_sToolTipText)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sToolTipText = PropBag.ReadProperty("ToolTipText", "")
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sTooltiptitle = PropBag.ReadProperty("Tooltiptitle", "")
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lToolTipIcon = PropBag.ReadProperty("ToolTipIcon", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lToolTipType = PropBag.ReadProperty("ToolTipType", 1)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lttBackColor = PropBag.ReadProperty("ttBackColor", GetSysColor(COLOR_INFOTEXT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lttForeColor = PropBag.ReadProperty("ttForeColor", GetSysColor(COLOR_INFOBK))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			MyBase.Font = .ReadProperty("Font", MyBase.Font)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_ButtonType = .ReadProperty("ButtonType", isbButtonType.isbButton)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_Value = .ReadProperty("Value", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_UseMaskColor = .ReadProperty("UseMaskColor", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			MyBase.Cursor = .ReadProperty("MousePointer", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bRoundedBordersByTheme = .ReadProperty("RoundedBordersByTheme", True)
		End With
		
		If Not DesignMode Then 'If we're not in design mode
			bTrack = True
			bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
			
			If Not bTrackUser32 Then
				If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
					bTrack = False
				End If
			End If
			
			If bTrack Then
				
				'OS supports mouse leave so subclass for it
				With Me
					'Start subclassing the UserControl
					Call Subclass_Start(.Handle.ToInt32)
					Call Subclass_AddMsg(.Handle.ToInt32, WM_MOUSEMOVE, eMsgWhen.MSG_AFTER)
					Call Subclass_AddMsg(.Handle.ToInt32, WM_MOUSELEAVE, eMsgWhen.MSG_AFTER)
					Call Subclass_AddMsg(.Handle.ToInt32, WM_THEMECHANGED, eMsgWhen.MSG_AFTER)
					Call Subclass_AddMsg(.Handle.ToInt32, WM_SYSCOLORCHANGE, eMsgWhen.MSG_AFTER)
				End With
				
			End If
		End If
		
		Exit Sub
		
UserControl_ReadProperties_Error: 
	End Sub
	
	'The control is terminating - a good place to stop the subclasser
	Private Sub UserControl_Terminate()
		'On Error GoTo Catch
		
		Call Subclass_StopAll()
Catch_Renamed: 
	End Sub
	
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As Object)
		'On Error GoTo UserControl_WriteProperties_Error
		
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Icon", m_Icon)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Style", m_iStyle, 3)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Caption", m_sCaption, "")
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("IconSize", m_IconSize, 16)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("IconAlign", m_IconAlign, isbAlign.isbleft)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("CaptionAlign", m_CaptionAlign, isbAlign.isbCenter)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("iNonThemeStyle", m_iNonThemeStyle, isbStyle.isbWindowsXP)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Enabled", m_bEnabled, True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ShowFocus", m_bShowFocus, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("USeCustomColors", m_bUseCustomColors, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("BackColor", m_lBackColor, GetSysColor(COLOR_BTNFACE))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("HighlightColor", m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("FontColor", m_lFontColor, GetSysColor(COLOR_BTNTEXT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("FontHighlightColor", m_lFontHighlightColor, GetSysColor(COLOR_BTNTEXT))
			'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ToolTipText", m_sToolTipText, MyBase.Extender.ToolTipText)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Tooltiptitle", m_sTooltiptitle)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ToolTipIcon", m_lToolTipIcon)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ToolTipType", m_lToolTipType)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ttBackColor", m_lttBackColor, GetSysColor(COLOR_INFOTEXT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ttForeColor", m_lttForeColor, GetSysColor(COLOR_INFOBK))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Font", MyBase.Font)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("ButtonType", m_ButtonType, isbButtonType.isbButton)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("Value", m_Value, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("MaskColor", m_MaskColor, &HC0C0C0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("UseMaskColor", m_UseMaskColor, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("MousePointer", MyBase.Cursor, 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Call .WriteProperty("RoundedBordersByTheme", m_bRoundedBordersByTheme, True)
		End With
		
		Exit Sub
		
UserControl_WriteProperties_Error: 
	End Sub
	
	'======================================================================================================
	'UserControl private routines
	'Determine if the passed function is supported
	Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
		'On Error GoTo IsFunctionExported_Error
		
		Dim hMod As Integer
		Dim bLibLoaded As Boolean
		hMod = GetModuleHandleA(sModule)
		
		If hMod = 0 Then
			hMod = LoadLibraryA(sModule)
			
			If hMod Then
				bLibLoaded = True
			End If
		End If
		
		If hMod Then
			If GetProcAddress(hMod, sFunction) Then
				IsFunctionExported = True
			End If
		End If
		
		If bLibLoaded Then
			Call FreeLibrary(hMod)
		End If
		
		Exit Function
		
IsFunctionExported_Error: 
	End Function
	
	'Track the mouse leaving the indicated window
	Private Sub TrackMouseLeave(ByVal lng_hWnd As Integer)
		'On Error GoTo TrackMouseLeave_Error
		
		Dim tme As TRACKMOUSEEVENT_STRUCT
		
		If bTrack Then
			
			With tme
				.cbSize = Len(tme)
				.dwFlags = TRACKMOUSEEVENT_FLAGS.TME_LEAVE
				.hwndTrack = lng_hWnd
			End With
			
			If bTrackUser32 Then
				Call TrackMouseEvent(tme)
			Else
				Call TrackMouseEventComCtl(tme)
			End If
		End If
		
		Exit Sub
		
TrackMouseLeave_Error: 
	End Sub
	
	Public Function Version() As String
		'On Error GoTo Version_Error
		
		Version = strCurrentVersion
		Exit Function
		
Version_Error: 
	End Function
	
	' Description: this is the Style property.
	
	Public Property Style() As isbStyle
		Get
			'On Error GoTo Style_Error
			
			Style = m_iStyle
			Exit Property
			
Style_Error: 
		End Get
		Set(ByVal Value As isbStyle)
			'On Error GoTo Style_Error
			
			m_iStyle = Value
			RaiseEvent StyleChange()
			'For a small error when creating pixarray, I need to set the backcolor to white
			'If m_iStyle = isbMacOSX Then UserControl.BackColor = vbWhite
			ctlButton_Resize(Me, New System.EventArgs())
			Refresh()
			ctlButton_Resize(Me, New System.EventArgs())
			Refresh()
			Exit Property
			
Style_Error: 
		End Set
	End Property
	
	' Description: this is the "Caption" property.
	
	Public Property Caption() As String
		Get
			'On Error GoTo Caption_Error
			
			Caption = m_sCaption
			Exit Property
			
Caption_Error: 
		End Get
		Set(ByVal Value As String)
			'On Error GoTo Caption_Error
			
			m_sCaption = Value
			RaiseEvent CaptionChange()
			ctlButton_Resize(Me, New System.EventArgs())
			Refresh()
			Exit Property
			
Caption_Error: 
		End Set
	End Property
	
	' Description: this is the Picture Property
	
	Public Property Icon() As System.Drawing.Image
		Get
			'On Error GoTo Icon_Error
			
			Icon = m_Icon
			Exit Property
			
Icon_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Image)
			'On Error GoTo Icon_Error
			
			m_Icon = Value
			RaiseEvent IconChange()
			ctlButton_Resize(Me, New System.EventArgs())
			ctlButton_Paint(Me, New System.Windows.Forms.PaintEventArgs(Nothing, Nothing))
			Exit Property
			
Icon_Error: 
		End Set
	End Property
	
	' Description: this is the "IconAlign" property.
	
	Public Property IconAlign() As isbAlign
		Get
			'On Error GoTo IconAlign_Error
			
			IconAlign = m_IconAlign
			Exit Property
			
IconAlign_Error: 
		End Get
		Set(ByVal Value As isbAlign)
			'On Error GoTo IconAlign_Error
			
			m_IconAlign = Value
			RaiseEvent IconAlignChange()
			ctlButton_Resize(Me, New System.EventArgs())
			ctlButton_Paint(Me, New System.Windows.Forms.PaintEventArgs(Nothing, Nothing))
			Exit Property
			
IconAlign_Error: 
		End Set
	End Property
	
	' Description: this is the "IconSize" property.
	
	Public Property IconSize() As Short
		Get
			'On Error GoTo IconSize_Error
			
			IconSize = m_IconSize
			Exit Property
			
IconSize_Error: 
		End Get
		Set(ByVal Value As Short)
			'On Error GoTo IconSize_Error
			
			m_IconSize = Value
			RaiseEvent IconSizeChange()
			ctlButton_Resize(Me, New System.EventArgs())
			ctlButton_Paint(Me, New System.Windows.Forms.PaintEventArgs(Nothing, Nothing))
			Exit Property
			
IconSize_Error: 
		End Set
	End Property
	
	' Description: this is the "CaptionAlign" property.
	
	Public Property CaptionAlign() As isbAlign
		Get
			'On Error GoTo CaptionAlign_Error
			
			CaptionAlign = m_CaptionAlign
			Exit Property
			
CaptionAlign_Error: 
		End Get
		Set(ByVal Value As isbAlign)
			'On Error GoTo CaptionAlign_Error
			
			m_CaptionAlign = Value
			RaiseEvent CaptionAlignChange()
			ctlButton_Resize(Me, New System.EventArgs())
			ctlButton_Paint(Me, New System.Windows.Forms.PaintEventArgs(Nothing, Nothing))
			Exit Property
			
CaptionAlign_Error: 
		End Set
	End Property
	
	' Description: When Themed Faile, Use this style:
	
	Public Property NonThemeStyle() As isbStyle
		Get
			'On Error GoTo NonThemeStyle_Error
			
			NonThemeStyle = m_iNonThemeStyle
			Exit Property
			
NonThemeStyle_Error: 
		End Get
		Set(ByVal Value As isbStyle)
			'On Error GoTo NonThemeStyle_Error
			
			m_iNonThemeStyle = Value
			RaiseEvent NonThemeStyleChange()
			ctlButton_Resize(Me, New System.EventArgs())
			ctlButton_Paint(Me, New System.Windows.Forms.PaintEventArgs(Nothing, Nothing))
			Exit Property
			
NonThemeStyle_Error: 
		End Set
	End Property
	
	'Description: Enable or disable the control
	
	Public Shadows Property Enabled() As Boolean
		Get
			'On Error GoTo Enabled_Error
			
			Enabled = m_bEnabled
			Refresh()
			Exit Property
			
Enabled_Error: 
		End Get
		Set(ByVal Value As Boolean)
			'On Error GoTo Enabled_Error
			
			m_bEnabled = Value
			m_iState = isState.statenormal
			Refresh()
			RaiseEvent EnabledChange()
			MyBase.Enabled = m_bEnabled
			Exit Property
			
Enabled_Error: 
		End Set
	End Property
	
	'Description: Do we want to show Focus?
	
	Public Property ShowFocus() As Boolean
		Get
			'On Error GoTo ShowFocus_Error
			
			ShowFocus = m_bShowFocus
			Exit Property
			
ShowFocus_Error: 
		End Get
		Set(ByVal Value As Boolean)
			'On Error GoTo ShowFocus_Error
			
			m_bShowFocus = Value
			RaiseEvent ShowFocusChange()
			Refresh()
			Exit Property
			
ShowFocus_Error: 
		End Set
	End Property
	
	'Description: Will we use custom colors?
	'             If not, system colors will be used
	
	Public Property UseCustomColors() As Boolean
		Get
			'On Error GoTo UseCustomColors_Error
			
			UseCustomColors = m_bUseCustomColors
			Exit Property
			
UseCustomColors_Error: 
		End Get
		Set(ByVal Value As Boolean)
			'On Error GoTo UseCustomColors_Error
			
			m_bUseCustomColors = Value
			RaiseEvent UseCustomColorsChange()
			Refresh()
			Exit Property
			
UseCustomColors_Error: 
		End Set
	End Property
	
	'Description: Use this color for drawing
	
	Public Overrides Property BackColor() As System.Drawing.Color
		Get
			'On Error GoTo BackColor_Error
			
			BackColor = System.Drawing.ColorTranslator.FromOle(m_lBackColor)
			Exit Property
			
BackColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			'On Error GoTo BackColor_Error
			
			m_lBackColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent BackColorChange()
			Refresh()
			Exit Property
			
BackColor_Error: 
		End Set
	End Property
	
	'Description: Use this color for drawing
	
	Public Property HighlightColor() As System.Drawing.Color
		Get
			'On Error GoTo HighlightColor_Error
			
			HighlightColor = System.Drawing.ColorTranslator.FromOle(m_lHighlightColor)
			Exit Property
			
HighlightColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			'On Error GoTo HighlightColor_Error
			
			m_lHighlightColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent HighlightColorChange()
			Refresh()
			Exit Property
			
HighlightColor_Error: 
		End Set
	End Property
	
	'Description: Use this color for drawing normal font
	
	Public Property FontColor() As System.Drawing.Color
		Get
			'On Error GoTo FontColor_Error
			
			FontColor = System.Drawing.ColorTranslator.FromOle(m_lFontColor)
			Exit Property
			
FontColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			'On Error GoTo FontColor_Error
			
			m_lFontColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent FontColorChange()
			Refresh()
			Exit Property
			
FontColor_Error: 
		End Set
	End Property
	
	'Description: Use this color for drawing normal font
	
	Public Property FontHighlightColor() As System.Drawing.Color
		Get
			'On Error GoTo FontHighlightColor_Error
			
			FontHighlightColor = System.Drawing.ColorTranslator.FromOle(m_lFontHighlightColor)
			Exit Property
			
FontHighlightColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			'On Error GoTo FontHighlightColor_Error
			
			m_lFontHighlightColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent FontHighlightColorChange()
			Refresh()
			Exit Property
			
FontHighlightColor_Error: 
		End Set
	End Property
	
	'Description: Set TooltipText
	'This sub Is never executed:(
	
	Public Property ToolTipText() As String
		Get
			On Error GoTo ToolTipText_Error
			ToolTipText = m_sToolTipText
			Exit Property
ToolTipText_Error: 
		End Get
		Set(ByVal Value As String)
			'UserControl.Extender.ToolTipText = sToolTipText
			'On Error GoTo ToolTipText_Error
			
			m_sToolTipText = Value
			CreateToolTip()
			RaiseEvent ToolTipTextChange()
			Refresh()
			Exit Property
			
ToolTipText_Error: 
		End Set
	End Property
	
	
	Public Property ToolTipTitle() As String
		Get
			On Error GoTo ToolTipTitle_Error
			ToolTipTitle = m_sTooltiptitle
			Exit Property
ToolTipTitle_Error: 
		End Get
		Set(ByVal Value As String)
			On Error GoTo ToolTipTitle_Error
			m_sTooltiptitle = Value
			RaiseEvent TooltipTitleChange()
			Refresh()
			Exit Property
ToolTipTitle_Error: 
		End Set
	End Property
	
	'Description: Set TooltipIcon
	
	Public Property ToolTipIcon() As ttIconType
		Get
			On Error GoTo ToolTipIcon_Error
			ToolTipIcon = m_lToolTipIcon
			Exit Property
ToolTipIcon_Error: 
		End Get
		Set(ByVal Value As ttIconType)
			On Error GoTo ToolTipIcon_Error
			m_lToolTipIcon = Value
			RaiseEvent TooltipIconChange()
			Refresh()
			Exit Property
ToolTipIcon_Error: 
		End Set
	End Property
	
	
	Public Property ToolTipType() As ttStyleEnum
		Get
			On Error GoTo ToolTipType_Error
			ToolTipType = m_lToolTipType
			Exit Property
ToolTipType_Error: 
		End Get
		Set(ByVal Value As ttStyleEnum)
			On Error GoTo ToolTipType_Error
			m_lToolTipType = Value
			RaiseEvent ToolTipTypeChange()
			Refresh()
			Exit Property
ToolTipType_Error: 
		End Set
	End Property
	
	
	Public Property ToolTipBackColor() As System.Drawing.Color
		Get
			On Error GoTo ToolTipBackColor_Error
			ToolTipBackColor = System.Drawing.ColorTranslator.FromOle(m_lttBackColor)
			Exit Property
ToolTipBackColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			On Error GoTo ToolTipBackColor_Error
			m_lttBackColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent ToolTipBackColorChange()
			Refresh()
			Exit Property
ToolTipBackColor_Error: 
		End Set
	End Property
	
	
	Public Property ToolTipForeColor() As System.Drawing.Color
		Get
			On Error GoTo ToolTipForeColor_Error
			ToolTipForeColor = System.Drawing.ColorTranslator.FromOle(m_lttForeColor)
			Exit Property
ToolTipForeColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			On Error GoTo ToolTipForeColor_Error
			m_lttForeColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent ToolTipForeColorChange()
			Refresh()
			Exit Property
ToolTipForeColor_Error: 
		End Set
	End Property
	
	
	Public Overrides Property Font() As System.Drawing.Font
		Get
			On Error GoTo Font_Error
			Font = MyBase.Font
			Exit Property
Font_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Font)
			On Error GoTo Font_Error
			m_Font = Value
			MyBase.Font = Value
			Refresh()
			RaiseEvent FontChange()
			Exit Property
Font_Error: 
		End Set
	End Property
	
	
	Public Property ButtonType() As isbButtonType
		Get
			On Error GoTo ButtonType_Error
			ButtonType = m_ButtonType
			Exit Property
ButtonType_Error: 
		End Get
		Set(ByVal Value As isbButtonType)
			On Error GoTo ButtonType_Error
			m_ButtonType = Value
			RaiseEvent ButtonTypeChange()
			Refresh()
			Exit Property
ButtonType_Error: 
		End Set
	End Property
	
	
	Public Property MousePointer() As System.Windows.Forms.Cursor
		Get
			On Error GoTo MousePointer_Error
			MousePointer = MyBase.Cursor
			Exit Property
MousePointer_Error: 
		End Get
		Set(ByVal Value As System.Windows.Forms.Cursor)
			On Error GoTo MousePointer_Error
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			MyBase.Cursor = Value
			RaiseEvent MousePointerChange()
			Exit Property
MousePointer_Error: 
		End Set
	End Property
	
	
	Public Property Value() As Boolean
		Get
			On Error GoTo Value_Error
			Value = m_Value
			Exit Property
Value_Error: 
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo Value_Error
			m_Value = Value
			RaiseEvent ValueChange()
			Refresh()
			Exit Property
Value_Error: 
		End Set
	End Property
	
	
	Public Property MaskColor() As System.Drawing.Color
		Get
			On Error GoTo MaskColor_Error
			MaskColor = System.Drawing.ColorTranslator.FromOle(m_MaskColor)
			Exit Property
MaskColor_Error: 
		End Get
		Set(ByVal Value As System.Drawing.Color)
			On Error GoTo MaskColor_Error
			m_MaskColor = System.Drawing.ColorTranslator.ToOle(Value)
			RaiseEvent MaskColorChange()
			Refresh()
			Exit Property
MaskColor_Error: 
		End Set
	End Property
	
	
	Public Property UseMaskColor() As Boolean
		Get
			On Error GoTo UseMaskColor_Error
			UseMaskColor = m_UseMaskColor
			Exit Property
UseMaskColor_Error: 
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo UseMaskColor_Error
			m_UseMaskColor = Value
			RaiseEvent UseMaskColorChange()
			Refresh()
			Exit Property
UseMaskColor_Error: 
		End Set
	End Property
	
	
	Public Property RoundedBordersByTheme() As Boolean
		Get
			On Error GoTo RoundedBordersByTheme_Error
			RoundedBordersByTheme = m_bRoundedBordersByTheme
RoundedBordersByTheme_Error: 
		End Get
		Set(ByVal Value As Boolean)
			m_bRoundedBordersByTheme = Value
			RaiseEvent RoundedBordersByThemeChange()
			Refresh()
RoundedBordersByTheme_Error: 
		End Set
	End Property
	
	Public Function OpenLink(ByRef sLink As String) As Integer
		OpenLink = ShellExecute(Handle.ToInt32, "open", sLink, CStr(VariantType.Null), CStr(VariantType.Null), 1)
		Exit Function
OpenLink_Error: 
	End Function
End Class