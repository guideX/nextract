Option Strict Off
Option Explicit On
Friend Class ctlProgressBar
	Inherits System.Windows.Forms.UserControl
	Public Event ScrollingChange()
	Public Event ShowTextChange()
	Public Event MinChange()
	Public Event MaxChange()
	Public Event ImageChange()
	Public Event BrushStyleChange()
	Public Event OrientationChange()
	Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function CreateDC Lib "gdi32"  Alias "CreateDCA"(ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As Any) As Integer
	Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Short, ByVal COLORREF As Integer) As Integer
	Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Integer) As Integer
	Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawText Lib "user32"  Alias "DrawTextA"(ByVal hdc As Integer, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Integer, ByRef qrc As RECT, ByVal edge As Integer, ByVal grfFlags As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FrameRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef lpPoint As POINTAPI) As Integer
	Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Integer, ByVal nBkMode As Integer) As Integer
	Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Integer, ByVal crColor As Integer) As Integer
	Const DT_SINGLELINE As Integer = &H20s
	Const DT_CALCRECT As Integer = &H400s
	Const BF_BOTTOM As Short = &H8s
	Const BF_LEFT As Short = &H1s
	Const BF_RIGHT As Short = &H4s
	Const BF_TOP As Short = &H2s
	Const BF_RECT As Boolean = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
	Private Structure POINTAPI
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
	Public Enum BrushStyle
		HS_HORIZONTAL = 0
		HS_VERTICAL = 1
		HS_FDIAGONAL = 2
		HS_BDIAGONAL = 3
		HS_CROSS = 4
		HS_DIAGCROSS = 5
		HS_SOLID = 6
	End Enum
	Public Enum cScrolling
		ccScrollingStandard = 0
		ccScrollingSmooth = 1
		ccScrollingSearch = 2
		ccScrollingOfficeXP = 3
		ccScrollingPastel = 4
		ccScrollingJavT = 5
		ccScrollingMediaPlayer = 6
		ccScrollingCustomBrush = 7
		ccScrollingPicture = 8
		ccScrollingMetallic = 9
	End Enum
	Public Enum cOrientation
		ccOrientationHorizontal = 0
		ccOrientationVertical = 1
	End Enum
	Private m_Color As System.Drawing.Color
	Private m_hDC As Integer
	Private m_hWnd As Integer
	Private m_Max As Integer
	Private m_Min As Integer
	Private m_Value As Integer
	Private m_ShowText As Boolean
	Private m_Scrolling As cScrolling
	Private m_Orientation As cOrientation
	Private m_Brush As BrushStyle
	Private m_Picture As System.Drawing.Image
	Private m_MemDC As Boolean
	Private m_ThDC As Integer
	Private m_hBmp As Integer
	Private m_hBmpOld As Integer
	Private iFnt As System.Drawing.Font
	Private m_fnt As System.Drawing.Font
	Private hFntOld As Integer
	Private m_lWidth As Integer
	Private m_lHeight As Integer
	Private fPercent As Double
	Private TR As RECT
	Private TBR As RECT
	Private TSR As RECT
	Private AT As RECT
	Private lSegmentWidth As Integer
	Private lSegmentSpacing As Integer
	
	Public Sub DrawProgressBar()
		If m_Value > 100 Then m_Value = 100
		GetClientRect(m_hWnd, TR)
		DrawFillRectangle(TR, IIf(m_Scrolling = cScrolling.ccScrollingMediaPlayer, &H0s, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)), m_hDC)
		If m_Scrolling = cScrolling.ccScrollingMetallic Then
			DrawMetalProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingOfficeXP Then 
			DrawOfficeXPProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingPastel Then 
			DrawPastelProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingJavT Then 
			DrawJavTProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingMediaPlayer Then 
			DrawMediaProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingCustomBrush Then 
			DrawCustomBrushProgressbar()
		ElseIf m_Scrolling = cScrolling.ccScrollingPicture Then 
			DrawPictureProgressbar()
		Else
			CalcBarSize()
			PBarDraw()
			If m_Scrolling = 0 Then DrawDivisions()
			pDrawBorder()
		End If
		DrawTexto()
		If m_MemDC Then
			With Me
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				pDraw(.hdc, 0, 0, VB6.PixelsToTwipsX(.ClientRectangle.Width), VB6.PixelsToTwipsY(VB6.PixelsToTwipsY(.ClientRectangle.Height)), VB6.PixelsToTwipsY(.ClientRectangle.Left), VB6.PixelsToTwipsY(.ClientRectangle.Top))
			End With
		End If
	End Sub
	
	'==========================================================
	'/---OFFICE XP STYLE
	'==========================================================
	Private Sub DrawOfficeXPProgressbar()
		
		DrawRectangle(TR, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 100), m_hDC)
		
		With TBR
			.Left_Renamed = 1
			.Top_Renamed = 1
			.Bottom_Renamed = TR.Bottom_Renamed - 1
			.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed) * (m_Value / 100)
		End With
		
		DrawFillRectangle(TBR, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 180), m_hDC)
		
	End Sub
	'==========================================================
	'/---JAVT XP STYLE
	'==========================================================
	Private Sub DrawJavTProgressbar()
		
		DrawRectangle(TR, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 10), m_hDC)
		TBR.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed) * (m_Value / 101)
		DrawGradient(System.Drawing.ColorTranslator.ToOle(m_Color), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 100), 2, 2, TR.Right_Renamed - 2, TR.Bottom_Renamed - 5, m_hDC) ', True
		DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 250), System.Drawing.ColorTranslator.ToOle(m_Color), 3, 3, TBR.Right_Renamed, TR.Bottom_Renamed - 6, m_hDC) ', True
		DrawLine(TBR.Right_Renamed, 2, TBR.Right_Renamed, TR.Bottom_Renamed - 2, m_hDC, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 25))
		
	End Sub
	'==========================================================
	'/---PICTURE STYLE
	'==========================================================
	Private Sub DrawPictureProgressbar()
		
		Dim Brush As Integer
		Dim origBrush As Integer
		
		DrawEdge(m_hDC, TR, 2, BF_RECT) '//--- Draw ProgressBar Border
		
		If Nothing Is m_Picture Then Exit Sub '//--- In Case No Picture is Choosen
		
		'UPGRADE_ISSUE: Picture property m_Picture.handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Brush = CreatePatternBrush(m_Picture.Handle) '//-- Use Pattern Picture Draw
		origBrush = SelectObject(m_hDC, Brush)
		TBR.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed) * (m_Value / 101)
		
		'UPGRADE_ISSUE: Constant vbPatCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		PatBlt(m_hDC, 2, 2, TBR.Right_Renamed, TR.Bottom_Renamed - 4, vbPatCopy)
		
		SelectObject(m_hDC, origBrush)
		DeleteObject(Brush)
		
	End Sub
	'==========================================================
	'/---PASTEL XP STYLE
	'==========================================================
	Private Sub DrawPastelProgressbar()
		DrawEdge(m_hDC, TR, 6, BF_RECT)
		DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 140), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 200), 2, 2, TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed - 4) * (m_Value / 100), TR.Bottom_Renamed - 3, m_hDC, True)
	End Sub
	
	'==========================================================
	'/---METALLIC XP STYLE
	'==========================================================
	Private Sub DrawMetalProgressbar()
		TBR.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed - 4) * (m_Value / 100)
		
		DrawGradient(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White), &HC0C0C0, 2, 2, TR.Right_Renamed - 3, (TR.Bottom_Renamed - 3) / 2, m_hDC)
		DrawGradient(BlendColor(System.Drawing.ColorTranslator.FromOle(&HC0C0C0), System.Drawing.ColorTranslator.FromOle(&H0s), 255), &HC0C0C0, 2, (TR.Bottom_Renamed - 3) / 2, TR.Right_Renamed - 3, (TR.Bottom_Renamed - 3) / 2, m_hDC)
		DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 150), BlendColor(m_Color, System.Drawing.ColorTranslator.FromOle(&H0s), 180), 2, 2, TBR.Right_Renamed, (TR.Bottom_Renamed - 3) / 2, m_hDC)
		DrawGradient(BlendColor(m_Color, System.Drawing.ColorTranslator.FromOle(&H0s), 190), System.Drawing.ColorTranslator.ToOle(m_Color), 2, (TR.Bottom_Renamed - 3) / 2, TBR.Right_Renamed, (TR.Bottom_Renamed - 3) / 2, m_hDC)
		
		TR.Left_Renamed = TR.Left_Renamed + 3
		pDrawBorder()
		
		
	End Sub
	'==========================================================
	'/---CUSTOM BRUSH XP STYLE
	'==========================================================
	Private Sub DrawCustomBrushProgressbar()
		
		Dim hBrush As Integer
		
		DrawEdge(m_hDC, TR, 9, BF_RECT)
		
		With TBR
			.Left_Renamed = 2
			.Top_Renamed = 2
			.Bottom_Renamed = TR.Bottom_Renamed - 2
			.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed) * (m_Value / 101)
		End With
		
		hBrush = CreateHatchBrush(m_Brush, GetLngColor(System.Drawing.ColorTranslator.ToOle(Color)))
		SetBkColor(m_hDC, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 140))
		FillRect(m_hDC, TBR, hBrush)
		DeleteObject(hBrush)
		
	End Sub
	'==========================================================
	'/---MEDIA PROGRESS XP STYLE
	'==========================================================
	Private Sub DrawMediaProgressbar()
		
		DrawRectangle(TR, BlendColor(m_Color, System.Drawing.ColorTranslator.FromOle(&H0s), 200), m_hDC)
		DrawGradient(&H0, ShiftColorXP(GetLngColor(BlendColor(m_Color, System.Drawing.ColorTranslator.FromOle(&H0s), 100)), 10), 2, 2, TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed - 5) * (m_Value / 100), TR.Bottom_Renamed - 2, m_hDC, True)
		
	End Sub
	
	'==========================================================
	'/---Calculate Division Bars & Percent Values
	'==========================================================
	
	Private Sub CalcBarSize()
		
		lSegmentWidth = IIf(m_Scrolling = 0, 6, 0) '/-- Windows Default
		lSegmentSpacing = 2 '/-- Windows Default
		
		TR.Left_Renamed = TR.Left_Renamed + 3
		
		'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
		TBR = LSet(TR)
		
		fPercent = m_Value / 98
		
		If fPercent < 0# Then fPercent = 0#
		
		If m_Orientation = 0 Then
			
			'=======================================================================================
			'                                 Calc Horizontal ProgressBar
			'---------------------------------------------------------------------------------------
			
			TBR.Right_Renamed = TR.Left_Renamed + (TR.Right_Renamed - TR.Left_Renamed) * fPercent
			
			TBR.Right_Renamed = TBR.Right_Renamed - ((TBR.Right_Renamed - TBR.Left_Renamed) Mod (lSegmentWidth + lSegmentSpacing))
			
			If TBR.Right_Renamed < TR.Left_Renamed Then
				TBR.Right_Renamed = TR.Left_Renamed
			End If
			
		Else
			
			'=======================================================================================
			'                                 Calc Vertical ProgressBar
			'---------------------------------------------------------------------------------------
			fPercent = 1# - fPercent
			TBR.Top_Renamed = TR.Top_Renamed + (TR.Bottom_Renamed - TR.Top_Renamed) * fPercent
			TBR.Top_Renamed = TBR.Top_Renamed - ((TBR.Top_Renamed - TBR.Bottom_Renamed) Mod (lSegmentWidth + lSegmentSpacing))
			If TBR.Top_Renamed > TR.Bottom_Renamed Then TBR.Top_Renamed = TR.Bottom_Renamed
			
			
			
		End If
		
	End Sub
	
	'==========================================================
	'/---Draw Division Bars
	'==========================================================
	
	Private Sub DrawDivisions()
		Dim i As Integer
		Dim hBR As Integer
		
		hBR = CreateSolidBrush(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White))
		
		'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
		TSR = LSet(TR)
		
		
		If m_Orientation = 0 Then
			
			
			'=======================================================================================
			'                                 Draw Horizontal ProgressBar
			'---------------------------------------------------------------------------------------
			For i = TBR.Left_Renamed + lSegmentWidth To TBR.Right_Renamed Step lSegmentWidth + lSegmentSpacing
				TSR.Left_Renamed = i + 1
				TSR.Right_Renamed = i + 1 + lSegmentSpacing
				FillRect(m_hDC, TSR, hBR)
			Next i
			'---------------------------------------------------------------------------------------
			
		Else
			
			'=======================================================================================
			'                                  Draw Vertical ProgressBar
			'---------------------------------------------------------------------------------------
			For i = TBR.Bottom_Renamed To TBR.Top_Renamed + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
				TSR.Top_Renamed = i - 2
				TSR.Bottom_Renamed = i - 2 + lSegmentSpacing
				FillRect(m_hDC, TSR, hBR)
			Next i
			'---------------------------------------------------------------------------------------
			
		End If
		
		DeleteObject(hBR)
		
	End Sub
	
	
	'==========================================================
	'/---Draw The ProgressXP Bar Border  ;)
	'==========================================================
	
	Private Sub pDrawBorder()
		Dim RTemp As RECT
		
		TR.Left_Renamed = TR.Left_Renamed - 3
		
		'UPGRADE_WARNING: Couldn't resolve default property of object RTemp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RTemp = TR
		
		
		DrawLine(2, 1, TR.Right_Renamed - 2, 1, m_hDC, &HBEBEBE)
		DrawLine(2, TR.Bottom_Renamed - 2, TR.Right_Renamed - 2, TR.Bottom_Renamed - 2, m_hDC, &HEFEFEF)
		DrawLine(1, 2, 1, TR.Bottom_Renamed - 2, m_hDC, &HBEBEBE)
		DrawLine(2, 2, 2, TR.Bottom_Renamed - 2, m_hDC, &HEFEFEF)
		DrawLine(2, 2, TR.Right_Renamed - 2, 2, m_hDC, &HEFEFEF)
		DrawLine(TR.Right_Renamed - 2, 2, TR.Right_Renamed - 2, TR.Bottom_Renamed - 2, m_hDC, &HEFEFEF)
		
		DrawRectangle(TR, GetLngColor(&H686868), m_hDC)
		
		
		Call SetPixelV(m_hDC, 0, 0, GetLngColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)))
		Call SetPixelV(m_hDC, 0, 1, GetLngColor(&HA6ABAC))
		Call SetPixelV(m_hDC, 0, 2, GetLngColor(&H7D7E7F))
		Call SetPixelV(m_hDC, 1, 0, GetLngColor(&HA7ABAC)) '//TOP RIGHT CORNER
		Call SetPixelV(m_hDC, 1, 1, GetLngColor(&H777777))
		Call SetPixelV(m_hDC, 2, 0, GetLngColor(&H7D7E7F))
		Call SetPixelV(m_hDC, 2, 2, GetLngColor(&HBEBEBE))
		
		Call SetPixelV(m_hDC, 0, TR.Bottom_Renamed - 1, GetLngColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)))
		Call SetPixelV(m_hDC, 1, TR.Bottom_Renamed - 1, GetLngColor(&HA6ABAC))
		Call SetPixelV(m_hDC, 2, TR.Bottom_Renamed - 1, GetLngColor(&H7D7E7F))
		Call SetPixelV(m_hDC, 0, TR.Bottom_Renamed - 3, GetLngColor(&H7D7E7F)) '//BOTTOM RIGHT CORNER
		Call SetPixelV(m_hDC, 0, TR.Bottom_Renamed - 2, GetLngColor(&HA7ABAC))
		Call SetPixelV(m_hDC, 1, TR.Bottom_Renamed - 2, GetLngColor(&H777777))
		
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, 0, GetLngColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, 1, GetLngColor(&HBEBEBE))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, 2, GetLngColor(&H7D7E7F)) '//TOP LEFT CORNER
		Call SetPixelV(m_hDC, TR.Right_Renamed - 2, 2, GetLngColor(&HBEBEBE))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 2, 1, GetLngColor(&H686868))
		
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, TR.Bottom_Renamed - 1, GetLngColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, TR.Bottom_Renamed - 2, GetLngColor(&HBEBEBE))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 1, TR.Bottom_Renamed - 3, GetLngColor(&H7D7E7F))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 2, TR.Bottom_Renamed - 2, GetLngColor(&H777777)) '//TOP RIGHT CORNER
		Call SetPixelV(m_hDC, TR.Right_Renamed - 2, TR.Bottom_Renamed - 1, GetLngColor(&HBEBEBE))
		Call SetPixelV(m_hDC, TR.Right_Renamed - 3, TR.Bottom_Renamed - 1, GetLngColor(&H7D7E7F))
		
		
	End Sub
	
	
	'==========================================================
	'/---Draw The ProgressXP Bar ;)
	'==========================================================
	
	Private Sub PBarDraw()
		Dim TempRect As RECT
		Dim ITemp As Integer
		
		If m_Orientation = 0 Then
			
			If TBR.Right_Renamed <= 14 Then TBR.Right_Renamed = 12
			
			TempRect.Left_Renamed = 4
			TempRect.Right_Renamed = IIf(TBR.Right_Renamed + 4 > TR.Right_Renamed, TBR.Right_Renamed - 4, TBR.Right_Renamed)
			TempRect.Top_Renamed = 8
			TempRect.Bottom_Renamed = TR.Bottom_Renamed - 8
			
			'=======================================================================================
			'                                 Draw Horizontal ProgressBar
			'---------------------------------------------------------------------------------------
			
			
			If m_Scrolling = cScrolling.ccScrollingSearch Then
				'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
				GoSub HorizontalSearch
			Else
				DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 150), System.Drawing.ColorTranslator.ToOle(m_Color), 4, 3, TempRect.Right_Renamed, 6, m_hDC)
				DrawFillRectangle(TempRect, System.Drawing.ColorTranslator.ToOle(m_Color), m_hDC)
				DrawGradient(System.Drawing.ColorTranslator.ToOle(m_Color), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 150), 4, TempRect.Bottom_Renamed - 2, TempRect.Right_Renamed, 6, m_hDC)
			End If
		Else
			
			TempRect.Left_Renamed = 9
			TempRect.Right_Renamed = TR.Right_Renamed - 8
			TempRect.Top_Renamed = TBR.Top_Renamed
			TempRect.Bottom_Renamed = TR.Bottom_Renamed
			
			'=======================================================================================
			'                                 Draw Vertical ProgressBar
			'---------------------------------------------------------------------------------------
			
			If m_Scrolling = cScrolling.ccScrollingSearch Then
				'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
				GoSub VerticalSearch
			Else
				DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 150), System.Drawing.ColorTranslator.ToOle(m_Color), 4, TBR.Top_Renamed, 4, TR.Bottom_Renamed, m_hDC, True)
				DrawFillRectangle(TempRect, System.Drawing.ColorTranslator.ToOle(m_Color), m_hDC)
				DrawGradient(System.Drawing.ColorTranslator.ToOle(m_Color), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 150), TR.Right_Renamed - 8, TBR.Top_Renamed, 4, TR.Bottom_Renamed, m_hDC, True)
			End If
			
			'--------------------   <-------- Gradient Color From (- to +)
			'||||||||||||||||||||   <-------- Fill Color
			'--------------------   <-------- Gradient Color From (+ to -)
			
		End If
		
		Exit Sub
		
HorizontalSearch: 
		
		
		For ITemp = 0 To 2
			
			With TempRect
				.Left_Renamed = TBR.Right_Renamed + ((lSegmentSpacing + 10) * (ITemp)) - (45 * ((100 - m_Value) / 100))
				.Right_Renamed = .Left_Renamed + 10
				.Top_Renamed = 8
				.Bottom_Renamed = TR.Bottom_Renamed - 8
				DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 220 - (40 * ITemp)), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 200 - (40 * ITemp)), .Left_Renamed, 3, 9, TR.Bottom_Renamed - 2, m_hDC, True)
			End With
			
		Next ITemp
		
		'UPGRADE_WARNING: Return has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Return 
		
VerticalSearch: 
		
		
		For ITemp = 0 To 2
			
			With TempRect
				.Left_Renamed = 8
				.Right_Renamed = TR.Right_Renamed - 8
				.Top_Renamed = TBR.Top_Renamed + ((lSegmentSpacing + 10) * ITemp)
				.Bottom_Renamed = .Top_Renamed + 10
				DrawGradient(ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 220 - (40 * ITemp)), ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 200 - (40 * ITemp)), TR.Right_Renamed - 2, .Top_Renamed, 2, 9, m_hDC)
			End With
			
		Next ITemp
		
		'UPGRADE_WARNING: Return has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Return 
		
	End Sub
	
	'======================================================================
	'DRAWS THE PERCENT TEXT ON PROGRESS BAR
	Private Function DrawTexto() As Object
		Dim ThisText As String
		Dim isAlpha As Boolean
		
		If (m_Scrolling = cScrolling.ccScrollingMediaPlayer Or m_Scrolling = cScrolling.ccScrollingMetallic) Then isAlpha = True
		
		
		If m_Scrolling = cScrolling.ccScrollingSearch Then
			ThisText = "Searching.."
		Else
			ThisText = System.Math.Round(m_Value) & " %"
		End If
		
		If (m_ShowText) Then
			
			iFnt = Font '//--New Font
			hFntOld = SelectObject(m_hDC, iFnt.hFont) '//--Use the New Font
			SetBkMode(m_hDC, 1) '//--Transparent Text
			
			'//--Use the Alpha Text Color Look if Progress is MediaPlayer Style, else Normal (Gray)
			SetTextColor(m_hDC, GetLngColor(IIf(m_Scrolling = cScrolling.ccScrollingMediaPlayer, &HC0C0C0, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))))
			
			CalculateAlphaTextRect(ThisText) '//--Calculate The Text Rectangle
			
			'//-- If ProgressBar is already over the Text don't draw the old text, yust draw the Alpha Text
			'It saves some memory
			
			If ((TR.Right_Renamed * (m_Value / 100)) <= AT.Right_Renamed) Or Not isAlpha Then
				DrawText(m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE)
			End If
			
			SelectObject(m_hDC, hFntOld) 'Delete the Used Font
			
			'//--Use the Alpha Text Look if Progress is AlPhA Style
			If isAlpha Then DrawAlphaText(ThisText)
			
		End If
		
		
	End Function
	'======================================================================
	
	'======================================================================
	'ALPHA TEXT RECT FUNCTION
	Private Sub CalculateAlphaTextRect(ByVal ThisText As String)
		
		'//--Calculates the Bounding Rects Of the Text using DT_CALCRECT
		DrawText(m_hDC, ThisText, Len(ThisText), AT, DT_CALCRECT)
		AT.Left_Renamed = (TR.Right_Renamed / 2) - ((AT.Right_Renamed - AT.Left_Renamed) / 2)
		AT.Top_Renamed = (TR.Bottom_Renamed / 2) - ((AT.Bottom_Renamed - AT.Top_Renamed) / 2)
		
	End Sub
	'======================================================================
	
	'======================================================================
	'ALPHA TEXT FUNCTION
	Private Sub DrawAlphaText(ByVal ThisText As String)
		
		iFnt = Font '//--New Font
		hFntOld = SelectObject(m_hDC, iFnt.hFont) '//--Use the New Font
		SetBkMode(m_hDC, 1) '//--Transparent Text
		
		
		'//-- This is When the Text is Drawn
		'//--Gives the Media Player Text Look (Changes Color When Progress is over the Text)
		
		If (TR.Right_Renamed * (m_Value / 100)) >= AT.Left_Renamed Then
			SetTextColor(m_hDC, GetLngColor(IIf(m_Scrolling = cScrolling.ccScrollingMediaPlayer, ShiftColorXP(System.Drawing.ColorTranslator.ToOle(m_Color), 80), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White))))
			AT.Left_Renamed = (TR.Right_Renamed / 2) - ((AT.Right_Renamed - AT.Left_Renamed) / 2)
			AT.Right_Renamed = (TR.Right_Renamed * (m_Value / 100))
			DrawText(m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE)
			SelectObject(m_hDC, hFntOld)
		End If
		
	End Sub
	'======================================================================
	
	'======================================================================
	'CONVERTION FUNCTION
	Private Function GetLngColor(ByRef Color As Integer) As Integer
		
		If (Color And &H80000000) Then
			GetLngColor = GetSysColor(Color And &H7FFFFFFF)
		Else
			GetLngColor = Color
		End If
	End Function
	'======================================================================
	
	'======================================================================
	'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
	Private Sub DrawRectangle(ByRef bRECT As RECT, ByVal Color As Integer, ByVal hdc As Integer)
		
		Dim hBrush As Integer
		
		hBrush = CreateSolidBrush(Color)
		FrameRect(hdc, bRECT, hBrush)
		DeleteObject(hBrush)
		
	End Sub
	'======================================================================
	
	'======================================================================
	'DRAWS A LINE WITH A DEFINED COLOR
	Public Sub DrawLine(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal cHdc As Integer, ByVal Color As Integer)
		
		Dim Pen1 As Integer
		Dim Pen2 As Integer
		Dim Outline As Integer
		Dim POS As POINTAPI
		
		Pen1 = CreatePen(0, 1, GetLngColor(Color))
		Pen2 = SelectObject(cHdc, Pen1)
		
		MoveToEx(cHdc, X, Y, POS)
		LineTo(cHdc, Width, Height)
		
		SelectObject(cHdc, Pen2)
		DeleteObject(Pen2)
		DeleteObject(Pen1)
		
	End Sub
	'======================================================================
	
	'======================================================================
	'BLENDS AN SPECIFIED COLOR TO GET XP COLOR LOOK
	Private Function ShiftColorXP(ByVal MyColor As Integer, ByVal Base As Integer) As Integer
		
		Dim b, R, G, Delta As Integer
		
		R = (MyColor And &HFFs)
		G = ((MyColor \ &H100s) Mod &H100s)
		b = ((MyColor \ &H10000) Mod &H100s)
		
		Delta = &HFFs - Base
		
		b = Base + b * Delta \ &HFFs
		G = Base + G * Delta \ &HFFs
		R = Base + R * Delta \ &HFFs
		
		If R > 255 Then R = 255
		If G > 255 Then G = 255
		If b > 255 Then b = 255
		
		ShiftColorXP = R + 256 * G + 65536 * b
		
	End Function
	'======================================================================
	
	'======================================================================
	'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
	Public Sub DrawGradient(ByRef lEndColor As Integer, ByRef lStartcolor As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal hdc As Integer, Optional ByRef bH As Boolean = False)
		On Error Resume Next
		
		''Draw a Vertical Gradient in the current HDC
		Dim sG, sR, sB As Single
		Dim eG, eR, eB As Single
		Dim ni As Integer
		
		lEndColor = GetLngColor(lEndColor)
		lStartcolor = GetLngColor(lStartcolor)
		
		sR = (lStartcolor And &HFFs)
		sG = (lStartcolor \ &H100s) And &HFFs
		sB = CShort(lStartcolor And &HFF0000) / &H10000
		eR = (lEndColor And &HFFs)
		eG = (lEndColor \ &H100s) And &HFFs
		eB = CShort(lEndColor And &HFF0000) / &H10000
		sR = (sR - eR) / IIf(bH, X2, Y2)
		sG = (sG - eG) / IIf(bH, X2, Y2)
		sB = (sB - eB) / IIf(bH, X2, Y2)
		
		
		For ni = 0 To IIf(bH, X2, Y2)
			
			If bH Then
				DrawLine(X + ni, Y, X + ni, Y2, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB)))
			Else
				DrawLine(X, Y + ni, X2, Y + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB)))
			End If
			
		Next ni
	End Sub
	'======================================================================
	
	'======================================================================
	'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
	Private Function BlendColor(ByVal oColorFrom As System.Drawing.Color, ByVal oColorTo As System.Drawing.Color, Optional ByVal Alpha As Integer = 128) As Integer
		Dim lCFrom As Integer
		Dim lCTo As Integer
		Dim lSrcR As Integer
		Dim lSrcG As Integer
		Dim lSrcB As Integer
		Dim lDstR As Integer
		Dim lDstG As Integer
		Dim lDstB As Integer
		
		lCFrom = GetLngColor(System.Drawing.ColorTranslator.ToOle(oColorFrom))
		lCTo = GetLngColor(System.Drawing.ColorTranslator.ToOle(oColorTo))
		
		lSrcR = lCFrom And &HFFs
		lSrcG = (lCFrom And &HFF00) \ &H100
		lSrcB = (lCFrom And &HFF0000) \ &H10000
		lDstR = lCTo And &HFFs
		lDstG = (lCTo And &HFF00) \ &H100
		lDstB = (lCTo And &HFF0000) \ &H10000
		
		BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
		
	End Function
	'======================================================================
	
	'======================================================================
	'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
	Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Integer, ByVal MyHdc As Integer)
		
		Dim hBrush As Integer
		
		hBrush = CreateSolidBrush(GetLngColor(Color))
		FillRect(MyHdc, hRect, hBrush)
		DeleteObject(hBrush)
		
	End Sub
	'======================================================================
	
	'======================================================================
	'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
	Private Function ThDC(ByRef Width As Integer, ByRef Height As Integer) As Integer
		If m_ThDC = 0 Then
			If (Width > 0) And (Height > 0) Then
				pCreate(Width, Height)
			End If
		Else
			If Width > m_lWidth Or Height > m_lHeight Then
				pCreate(Width, Height)
			End If
		End If
		ThDC = m_ThDC
	End Function
	'======================================================================
	
	'======================================================================
	'CREATES THE TEMP DC
	Private Sub pCreate(ByVal Width As Integer, ByVal Height As Integer)
		Dim lhDCC As Integer
		pDestroy()
		lhDCC = CreateDC("DISPLAY", "", "", 0)
		If Not (lhDCC = 0) Then
			m_ThDC = CreateCompatibleDC(lhDCC)
			If Not (m_ThDC = 0) Then
				m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
				If Not (m_hBmp = 0) Then
					m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
					If Not (m_hBmpOld = 0) Then
						m_lWidth = Width
						m_lHeight = Height
						DeleteDC(lhDCC)
						Exit Sub
					End If
				End If
			End If
			DeleteDC(lhDCC)
			pDestroy()
		End If
	End Sub
	'======================================================================
	
	'======================================================================
	'DRAWS THE TEMP DC
	Public Sub pDraw(ByVal hdc As Integer, Optional ByVal xSrc As Integer = 0, Optional ByVal ySrc As Integer = 0, Optional ByVal WidthSrc As Integer = 0, Optional ByVal HeightSrc As Integer = 0, Optional ByVal xDst As Integer = 0, Optional ByVal yDst As Integer = 0)
		If WidthSrc <= 0 Then WidthSrc = m_lWidth
		If HeightSrc <= 0 Then HeightSrc = m_lHeight
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy)
		
	End Sub
	'======================================================================
	
	'======================================================================
	'DESTROYS THE TEMP DC
	Private Sub pDestroy()
		If Not m_hBmpOld = 0 Then
			SelectObject(m_ThDC, m_hBmpOld)
			m_hBmpOld = 0
		End If
		If Not m_hBmp = 0 Then
			DeleteObject(m_hBmp)
			m_hBmp = 0
		End If
		If Not m_ThDC = 0 Then
			DeleteDC(m_ThDC)
			m_ThDC = 0
		End If
		m_lWidth = 0
		m_lHeight = 0
	End Sub
	'======================================================================
	
	
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'===========================================================================
	'USER CONTROL EVENTS
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'===========================================================================
	
	
	Private Sub UserControl_Initialize()
		
		Dim fnt As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Me.Font = fnt
		
		With Me
			.BackColor = System.Drawing.Color.White
			'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.ScaleMode = vbPixels
		End With
		
		'----------------------------------------------------------
		'Default Values
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Me.hdc = MyBase.hdc
		Me.hWnd = MyBase.Handle.ToInt32
		m_Max = 100
		m_Min = 0
		m_Value = 0
		m_Orientation = cOrientation.ccOrientationHorizontal
		m_Scrolling = cScrolling.ccScrollingStandard
		m_Color = System.Drawing.ColorTranslator.FromOle(GetLngColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Highlight)))
		DrawProgressBar()
		'----------------------------------------------------------
		
	End Sub
	
	Private Sub ctlProgressBar_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		DrawProgressBar()
	End Sub
	
	Private Sub ctlProgressBar_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Me.hdc = MyBase.hdc
	End Sub
	
	Private Sub UserControl_Terminate()
		pDestroy()
	End Sub
	
	Public WriteOnly Property BrushStyle_Renamed() As BrushStyle
		Set(ByVal Value As BrushStyle)
			m_Brush = Value
			RaiseEvent BrushStyleChange()
		End Set
	End Property
	
	
	Public Property Color() As System.Drawing.Color
		Get
			Color = m_Color
		End Get
		Set(ByVal Value As System.Drawing.Color)
			m_Color = System.Drawing.ColorTranslator.FromOle(GetLngColor(System.Drawing.ColorTranslator.ToOle(Value)))
			DrawProgressBar()
		End Set
	End Property
	
	
	
	Public Overrides Property Font() As System.Drawing.Font
		Get
			Font = m_fnt
		End Get
		Set(ByVal Value As System.Drawing.Font)
			If IsReference(Value) And Not TypeOf Value Is String Then
				m_fnt = Value
			Else
				m_fnt = Value
			End If
		End Set
	End Property
	
	
	Public Property hWnd() As Integer
		Get
			hWnd = m_hWnd
		End Get
		Set(ByVal Value As Integer)
			m_hWnd = Value
		End Set
	End Property
	
	
	Public Property hdc() As Integer
		Get
			hdc = m_hDC
		End Get
		Set(ByVal Value As Integer)
			m_hDC = ThDC(MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height))
			If m_hDC = 0 Then
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				m_hDC = MyBase.hdc
			Else
				m_MemDC = True
			End If
		End Set
	End Property
	
	
	Public Property Image() As System.Drawing.Image
		Get
			If Nothing Is m_Picture Then Exit Property
			Image = m_Picture
		End Get
		Set(ByVal Value As System.Drawing.Image)
			m_Picture = Value
			RaiseEvent ImageChange()
			DrawProgressBar()
		End Set
	End Property
	
	
	Public Property Min() As Integer
		Get
			Min = m_Min
		End Get
		Set(ByVal Value As Integer)
			m_Min = Value
			RaiseEvent MinChange()
		End Set
	End Property
	
	
	Public Property Max() As Integer
		Get
			Max = m_Max
		End Get
		Set(ByVal Value As Integer)
			m_Max = Value
			RaiseEvent MaxChange()
		End Set
	End Property
	
	
	Public Property Orientation() As cOrientation
		Get
			Orientation = m_Orientation
		End Get
		Set(ByVal Value As cOrientation)
			m_Orientation = Value
			RaiseEvent OrientationChange()
			DrawProgressBar()
		End Set
	End Property
	
	
	Public Property Scrolling() As cScrolling
		Get
			Scrolling = m_Scrolling
		End Get
		Set(ByVal Value As cScrolling)
			m_Scrolling = Value
			RaiseEvent ScrollingChange()
			DrawProgressBar()
		End Set
	End Property
	
	
	Public Property ShowText() As Boolean
		Get
			ShowText = m_ShowText
		End Get
		Set(ByVal Value As Boolean)
			m_ShowText = Value
			RaiseEvent ShowTextChange()
			DrawProgressBar()
		End Set
	End Property
	
	
	Public Property Value() As Integer
		Get
			Value = ((m_Value / 100) * m_Max) / IIf(m_Min > 0, m_Min, 1)
		End Get
		Set(ByVal Value As Integer)
			On Error Resume Next
			m_Value = ((Value * 100) / m_Max) + m_Min
			DrawProgressBar()
		End Set
	End Property
	
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As Object)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Font", Font)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("BrushStyle", m_Brush, 4)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Color", m_Color, System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Highlight))
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Image", m_Picture, Nothing)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Max", m_Max, 100)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Min", m_Min, 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Orientation", m_Orientation, cOrientation.ccOrientationHorizontal)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Scrolling", m_Scrolling, cScrolling.ccScrollingStandard)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("ShowText", m_ShowText, False)
		'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Call PropBag.WriteProperty("Value", m_Value, 0)
	End Sub
	
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As Object)
		On Error Resume Next
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_Brush = PropBag.ReadProperty("BrushStyle", 4)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Color = System.Drawing.ColorTranslator.FromOle(PropBag.ReadProperty("Color", System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Highlight)))
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_Picture = PropBag.ReadProperty("Image", Nothing)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Max = PropBag.ReadProperty("Max", 100)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Min = PropBag.ReadProperty("Min", 0)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Orientation = PropBag.ReadProperty("Orientation", cOrientation.ccOrientationHorizontal)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Scrolling = PropBag.ReadProperty("Scrolling", cScrolling.ccScrollingStandard)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.ShowText = PropBag.ReadProperty("ShowText", False)
		'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.Value = PropBag.ReadProperty("Value", 0)
	End Sub
End Class