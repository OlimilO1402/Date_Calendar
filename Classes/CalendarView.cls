VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Calendar        As Calendar
Private m_Canvas          As PictureBox 'Control ' As Printer AndAs PictureBox
Private m_LightMode       As Boolean ' True= Lightmode=white, false=Darkmode=black(Darkgrey)
Private m_HasMonthNames   As Boolean
Private m_HasWeekDayNames As Boolean
Private m_HasWeekNumbers  As Boolean
Private m_MarginCal       As Margin
Private m_MarginMon       As Margin ' not in use
Private m_MarginDay       As Margin ' not in use
Private m_ColorNormalGrid As Long   ' grey
Private m_ColorFestivlDay As Long   ' purple
Private m_ColorMouseOver  As Long   ' yellow
Private m_ColorSelected   As Long   ' red
Private m_ColorBirthDay   As Long   ' green
Private m_ColorMarriage   As Long
Private m_ColorWeekday    As Long   ' white
Private m_ColorSaturday   As Long   ' lighlight blue
Private m_ColorSunday     As Long   ' light blue
Private m_ColorLNWeekday  As Long   ' white
Private m_ColorLNSaturday As Long   ' lightlight grey
Private m_ColorLNSunday   As Long   ' light grey
Private m_ColTmpWeekday   As Long
Private m_ColTmpSaturday  As Long
Private m_ColTmpSunday    As Long
Private m_FontMonthName   As StdFont
Private m_FontDayNrName   As StdFont
Private m_FontWeekNr      As StdFont

Private m_DayRectWidth    As Double
Private m_DayRectHeight   As Double

Private Sub Class_Initialize()
    '
End Sub
Private Sub Class_Terminate()
    Me.Dispose
End Sub

Friend Sub New_(aCalendar As Calendar, Canvas As PictureBox)
    Set m_Calendar = aCalendar
    Set m_Canvas = Canvas
    
    m_LightMode = True
    m_HasMonthNames = True
    m_HasWeekDayNames = True
    m_HasWeekNumbers = True
    m_MarginCal = New_Margin(10)
    m_MarginMon = New_Margin(10)
    m_MarginDay = New_Margin(10)
    
    m_ColorNormalGrid = RGB(240, 240, 240)
    m_ColorFestivlDay = RGB(222, 141, 245)
    m_ColorMouseOver = RGB(255, 255, 0)
    m_ColorSelected = RGB(255, 0, 0)
    m_ColorBirthDay = RGB(0, 255, 0)
    m_ColorMarriage = RGB(250, 250, 250)
    m_ColorWeekday = RGB(255, 255, 255)
    m_ColorSaturday = RGB(230, 244, 253)
    m_ColorSunday = RGB(137, 189, 226)
    m_ColorLNWeekday = RGB(255, 255, 255)
    m_ColorLNSaturday = RGB(200, 202, 201)
    m_ColorLNSunday = RGB(157, 157, 157)
    Set m_FontDayNrName = New_StdFont("Segoe UI") '("Comic Sans MS")
    Set m_FontMonthName = New_StdFont("Segoe Print", 10, True) '("Comic Sans MS")
    Set m_FontWeekNr = New_StdFont("Segoe UI") '("Comic Sans MS")
    
    Resize
    
End Sub

Friend Sub NewC(other As CalendarView)
    With other
        Set m_Canvas = .Canvas
        m_LightMode = .LightMode
        m_HasMonthNames = .HasMonthNames
        m_HasWeekDayNames = True
        m_HasWeekNumbers = True
        m_MarginCal = .MarginCal ' New_Margin(10)
        m_MarginMon = .MarginMon ' New_Margin(10)
        m_MarginDay = .MarginDay ' New_Margin(10)
        
        m_ColorNormalGrid = RGB(240, 240, 240)
        m_ColorFestivlDay = RGB(222, 141, 245)
        m_ColorMouseOver = RGB(255, 255, 0)
        m_ColorSelected = RGB(255, 0, 0)
        m_ColorBirthDay = RGB(0, 255, 0)
        m_ColorMarriage = RGB(250, 250, 250)
        m_ColorWeekday = RGB(255, 255, 255)
        m_ColorSaturday = RGB(230, 244, 253)
        m_ColorSunday = RGB(137, 189, 226)
        m_ColorLNWeekday = RGB(255, 255, 255)
        m_ColorLNSaturday = RGB(200, 202, 201)
        m_ColorLNSunday = RGB(157, 157, 157)
        Set m_FontDayNrName = New_StdFont("Segoe UI") '("Comic Sans MS")
        Set m_FontMonthName = New_StdFont("Segoe Print", 10, True) '("Comic Sans MS")
        'Set .FontMonthName = New_StdFont("Comic Sans MS")
        Set m_FontWeekNr = New_StdFont("Segoe UI") '("Comic Sans MS")
        '.HasDecLastYear = True
        '.HasJanNextYear = True
    End With
End Sub

'Public Property Get CalendarView_DayWidth(this As CalendarView, CalYear As CalendarYear) As Double
'    With this
'        Dim n As Long: n = UBound(CalYear.Months) - LBound(CalYear.Months) + 1 + IIf(.HasDecLastYear, 1, 0) + IIf(.HasJanNextYear, 1, 0)
'        CalendarView_DayWidth = (.Canvas.ScaleWidth - .MarginCal.Left - .MarginCal.Right) / n
'    End With
'End Property
'

'Private Function CalcDayRectWidth() As Double
'    CalcDayRectWidth = (m_Canvas.ScaleWidth - m_MarginCal.Left - m_MarginCal.Right) / m_Calendar.CountMonth
'End Function
'
'Public Property Get CalendarView_DayHeight(this As CalendarView) As Double
'    With this
'        Dim n As Double: n = 32
'        CalendarView_DayHeight = (.Canvas.ScaleHeight - .MarginCal.Top - .MarginCal.Bottom - IIf(.HasMonthNames, .FontMonthName.Size, 0)) / n
'    End With
'End Property
'Private Function CalcDayRectHeight() As Double
'    CalcDayRectHeight = (m_Canvas.ScaleHeight - m_MarginCal.Top - m_MarginCal.Bottom) / 32
'End Function

Friend Function Clone() As CalendarView
    Set Clone = New CalendarView: Clone.NewC Me
End Function

Private Function CalcRectWidth() As Double
    Dim n As Integer: n = m_Calendar.CountMonth
    CalcRectWidth = (m_Canvas.ScaleWidth - m_MarginCal.Left - m_MarginCal.Right) / n
End Function

Private Function CalcRectHeight() As Double
    Dim n As Integer: n = 32
    CalcRectHeight = (m_Canvas.ScaleHeight - m_MarginCal.Top - m_MarginCal.Bottom - IIf(m_HasMonthNames, m_FontMonthName.Size, 0)) / n
End Function

Friend Sub Dispose()
    Set m_Canvas = Nothing
    Set m_FontDayNrName = Nothing
    Set m_FontMonthName = Nothing
    Set m_FontWeekNr = Nothing
End Sub

Public Sub Resize()
    m_DayRectWidth = CalcRectWidth
    m_DayRectHeight = CalcRectHeight
End Sub

Public Property Get Canvas() As Control
    Set Canvas = m_Canvas
End Property

Public Property Get CanvasPrinter() As Printer
    If TypeOf m_Canvas Is Printer Then
        Set CanvasPrinter = m_Canvas
    End If
End Property
Public Property Set CanvasPrinter(ByVal Value As Printer)
    Set m_Canvas = Printer
End Property

Public Property Get Calendar() As Calendar
    Set Calendar = m_Calendar
End Property
Public Property Set Calendar(Value As Calendar)
    Set m_Calendar = Value
End Property

Public Property Get LightMode() As Boolean
    LightMode = m_LightMode
End Property
Public Property Let LightMode(ByVal Value As Boolean)
    m_LightMode = Value
End Property

Public Property Get HasMonthNames() As Boolean
    HasMonthNames = m_HasMonthNames
End Property
Public Property Let HasMonthNames(ByVal Value As Boolean)
    m_HasMonthNames = Value
End Property

Public Property Get HasWeekDayNames() As Boolean
    HasWeekDayNames = m_HasWeekDayNames
End Property
Public Property Let HasWeekDayNames(ByVal Value As Boolean)
    m_HasWeekDayNames = Value
End Property

Public Property Get HasWeekNumbers() As Boolean
    HasWeekNumbers = m_HasWeekNumbers
End Property
Public Property Let HasWeekNumbers(ByVal Value As Boolean)
    m_HasWeekNumbers = Value
End Property

Friend Property Get MarginCal() As Margin
    MarginCal = m_MarginCal
End Property
Friend Property Let MarginCal(Value As Margin)
    m_MarginCal = Value
End Property

Friend Property Get MarginMon() As Margin     'not in use
    MarginMon = m_MarginMon
End Property
Friend Property Let MarginMon(Value As Margin)
    m_MarginMon = Value
End Property

Friend Property Get MarginDay() As Margin     'not in use
    MarginDay = m_MarginDay
End Property
Friend Property Let MarginDay(Value As Margin)
    m_MarginDay = Value
End Property

Public Property Get ColorNormalGrid() As Long 'grey
    ColorNormalGrid = m_ColorNormalGrid
End Property
Public Property Let ColorNormalGrid(ByVal Value As Long)
    m_ColorNormalGrid = Value
End Property

Public Property Get ColorFestivlDay() As Long 'purple
    ColorFestivlDay = m_ColorFestivlDay
End Property
Public Property Let ColorFestivlDay(ByVal Value As Long)
    m_ColorFestivlDay = Value
End Property

Public Property Get ColorSelected() As Long 'purple
    ColorSelected = m_ColorSelected
End Property
Public Property Let ColorSelected(ByVal Value As Long)
    m_ColorSelected = Value
End Property

Public Property Get ColorBirthDay() As Long 'green
    ColorBirthDay = m_ColorBirthDay
End Property
Public Property Let ColorBirthDay(ByVal Value As Long)
    m_ColorBirthDay = Value
End Property

Public Property Get ColorWeekday() As Long  'white
    ColorWeekday = m_ColorWeekday
End Property
Public Property Let ColorWeekday(ByVal Value As Long)
    m_ColorWeekday = Value
End Property

Public Property Get ColorSaturday() As Long 'lighlight blue
    ColorSaturday = m_ColorSaturday
End Property
Public Property Let ColorSaturday(ByVal Value As Long)
    m_ColorSaturday = Value
End Property

Public Property Get ColorSunday() As Long   'light blue
    ColorSunday = m_ColorSunday
End Property
Public Property Let ColorSunday(ByVal Value As Long)
    m_ColorSunday = Value
End Property

Public Property Get ColorLNWeekday() As Long 'white
    ColorLNWeekday = m_ColorLNWeekday
End Property
Public Property Let ColorLNWeekday(ByVal Value As Long)
    m_ColorLNWeekday = Value
End Property

Public Property Get ColorLNSaturday() As Long
    ColorLNSaturday = m_ColorLNSaturday
End Property
Public Property Let ColorLNSaturday(ByVal Value As Long)
    m_ColorLNSaturday = Value
End Property

Public Property Get ColorLNSunday() As Long
    ColorLNSunday = m_ColorLNSunday
End Property
Public Property Let ColorLNSunday(ByVal Value As Long)
    m_ColorLNSunday = Value
End Property

Public Property Get ColTmpWeekday() As Long
    ColTmpWeekday = m_ColTmpWeekday
End Property
Public Property Let ColTmpWeekday(ByVal Value As Long)
    m_ColTmpWeekday = Value
End Property

Public Property Get ColTmpSaturday() As Long
    ColTmpSaturday = m_ColTmpSaturday
End Property
Public Property Let ColTmpSaturday(ByVal Value As Long)
    m_ColTmpSaturday = Value
End Property

Public Property Get ColTmpSunday() As Long
    ColTmpSunday = m_ColTmpSunday
End Property
Public Property Let ColTmpSunday(ByVal Value As Long)
    m_ColTmpSunday = Value
End Property

Public Property Get FontMonthName() As StdFont
    Set FontMonthName = m_FontMonthName
End Property
Public Property Let FontMonthName(ByVal Value As StdFont)
    Set m_FontMonthName = Value
End Property

Public Property Get FontDayNrName() As StdFont
    Set FontDayNrName = m_FontDayNrName
End Property
Public Property Let FontDayNrName(ByVal Value As StdFont)
    Set m_FontDayNrName = Value
End Property

Public Property Get FontWeekNr() As StdFont
    Set FontWeekNr = m_FontWeekNr
End Property
Public Property Let FontWeekNr(ByVal Value As StdFont)
    Set m_FontWeekNr = Value
End Property

Public Sub Draw()
    With m_Calendar
        m_Canvas.CurrentX = m_MarginCal.Left
        Dim cy As CalendarYear
        Set cy = .LastYear: If Not cy Is Nothing Then DrawYear cy
        Set cy = .ThisYear: If Not cy Is Nothing Then DrawYear cy
        Set cy = .NextYear: If Not cy Is Nothing Then DrawYear cy
    End With
End Sub

Public Sub DrawYear(CalYear As CalendarYear)
    'If CalYear Is Nothing Then Exit Sub
    Dim cm As CalendarMonth
    Dim i As Long
    For i = CalYear.StartMonth To CalYear.EndMonth
        Set cm = CalYear.Month(i)
        Me.DrawMonth cm
        m_Canvas.CurrentX = i * m_DayRectWidth
    Next
End Sub

Public Sub DrawMonth(CalMonth As CalendarMonth)
'Try: On Error GoTo Catch
    Dim x As Double: x = m_Canvas.CurrentX
    Dim y As Double: y = m_MarginCal.Top
    Dim ny As Integer
    If m_HasMonthNames Then
        Set m_Canvas.Font = m_FontMonthName
        Dim s As String: s = MonthName(CalMonth.Month) & " '" & Right(CStr(CalMonth.Year.Year), 2)
        m_Canvas.Print s
        ny = ny + 1
        m_Canvas.CurrentY = m_MarginCal.Top + ny * m_DayRectHeight
    End If
    'm_Canvas.CurrentX = x
    Dim i As Integer
    For i = 1 To CalMonth.CountDays
        DrawDay CalMonth.Day(i)
        ny = ny + 1
        m_Canvas.CurrentY = m_MarginCal.Top + ny * m_DayRectHeight
    Next
    'm_Canvas.CurrentX = x + m_DayRectWidth
    m_Canvas.CurrentY = y
    'm_DayRectWidth
    'm_DayRectHeight
'Catch:
End Sub

Public Sub DrawDay(CalDay As CalendarDay)
'Try: On Error GoTo Catch
    Dim fc As Long: fc = m_Canvas.ForeColor
    Dim x As Double: x = m_Canvas.CurrentX
    Dim y As Double: y = m_Canvas.CurrentY
    Dim wd As VbDayOfWeek: wd = Weekday(CalDay.DayDate)
    Dim c As Long: c = IIf(wd = vbSaturday, m_ColTmpSaturday, IIf(wd = VbDayOfWeek.vbSunday, m_ColTmpSunday, m_ColTmpWeekday))
    
    m_Canvas.Line (x, y)-(x + m_DayRectWidth - 1, y + m_DayRectHeight - 1), c, BF
    
    Select Case True
    Case CalDay.IsMouseOver
        m_Canvas.DrawWidth = 2
        c = RGB(255, 0, 0)
    Case CalDay.IsFestivalDay
        m_Canvas.DrawWidth = 2
        c = m_ColorFestivlDay
    Case Else
        m_Canvas.DrawWidth = 1
        c = m_ColorNormalGrid
    End Select
    m_Canvas.Line (x, y)-(x + m_DayRectWidth - 1, y + m_DayRectHeight - 1), c, B
    
    m_Canvas.CurrentX = x
    m_Canvas.CurrentY = y
    
    Dim s As String
    s = CStr(CalDay.DayDate) & " " & MTime.Weekday_ToStr(wd, vbSunday, True)
    If CalDay.IsFestivalDay Then
        s = s & " " & CalDay.FestivalDay.Name ' MDECalendar.ELegalFestivals_ToStr(CalDay.FestivalIndex)
    Else
        If wd = vbMonday Then
            s = s & " " & "KW " & MTime.WeekOfYearISO(CalDay.DayDate)
        End If
    End If
    
    Set m_Canvas.Font = m_FontDayNrName
    If wd = vbSunday Then
        m_Canvas.ForeColor = RGB(255, 255, 255)
    End If
    
    m_Canvas.Print s
    
    m_Canvas.CurrentX = x
    m_Canvas.CurrentY = y '+ m_DayRectHeight
    m_Canvas.ForeColor = fc
'Catch:
End Sub

'
'Public Sub CalendarDay_MouseOut(this As CalendarDay, CalYear As CalendarYear)
'    With this
'        Dim m As Integer: m = Month(.Date)
'        If m = 0 Then Exit Sub
'        If .Day = 0 Then Exit Sub
'        CalYear.Months(Month(.Date)).Days(.Day).MouseOver = False
'    End With
'End Sub
'Public Property Get CalendarView_DayWidth(this As CalendarView, CalYear As CalendarYear) As Double
'    With this
'        Dim n As Long: n = UBound(CalYear.Months) - LBound(CalYear.Months) + 1 + IIf(.HasDecLastYear, 1, 0) + IIf(.HasJanNextYear, 1, 0)
'        CalendarView_DayWidth = (.Canvas.ScaleWidth - .MarginCal.Left - .MarginCal.Right) / n
'    End With
'End Property
'
'Public Property Get CalendarView_DayHeight(this As CalendarView) As Double
'    With this
'        Dim n As Double: n = 32
'        CalendarView_DayHeight = (.Canvas.ScaleHeight - .MarginCal.Top - .MarginCal.Bottom - IIf(.HasMonthNames, .FontMonthName.Size, 0)) / n
'    End With
'End Property

Public Function CalendarDayFromMouseCoords(Calendar As Calendar, ByVal MouseX As Single, ByVal MouseY As Single) As CalendarDay
    With Calendar
        MouseX = MouseX - m_MarginCal.Left
        MouseY = MouseY - m_MarginCal.Top
        Dim m As Integer: m = CInt(MouseX \ m_DayRectWidth)        ' x-axis
        Dim d As Integer: d = CInt(MouseY \ m_DayRectHeight) '- 1  ' y-axis
        Set CalendarDayFromMouseCoords = m_Calendar.GetDayFromGrid(m, d)
        'Dim dly As Integer: dly = IIf(.HasDecLastYear, 1, 0)
        'Dim jny As Integer: jny = IIf(.HasJanNextYear, 1, 0)
        'Dim cm As Integer: cm = m_Calendar.CountMonth
        'If 0 <= m And m <= cm Then
        '    If 0 < d And d < MTime.DaysInMonth(CalYear.Year, m) Then
        '        CalYear.Months(m).Days(d).MouseOver = True
        '        CalendarDayFromMouseCoords = CalYear.Months(m).Days(d)
        '    End If
        'End If
    End With
End Function

'Public Sub DrawYear(CalYear As CalendarYear)
''Try: On Error GoTo Catch
'    With this
'        Dim nx As Integer
'        .Canvas.CurrentX = .MarginCal.Left
'        .Canvas.CurrentY = .MarginCal.Top
'
'        .TmpDayWidth = CalendarView_DayWidth(this, CalYear)
'        .TmpDayHeight = CalendarView_DayHeight(this)
'
'        If .HasDecLastYear Then
'            Dim CalLastYear As CalendarYear:  CalLastYear = New_CalendarYear(CalYear.Year - 1, 12, 12)
'            Dim DecLastYear As CalendarMonth: DecLastYear = New_CalendarMonth(CalLastYear, 12)
'            .ColTmpWeekday = .ColorLNWeekday
'            .ColTmpSaturday = .ColorLNSaturday
'            .ColTmpSunday = .ColorLNSunday
'            CalendarView_DrawMonth this, DecLastYear
'            nx = nx + 1
'            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
'        End If
'
'        .ColTmpWeekday = .ColorWeekday
'        .ColTmpSaturday = .ColorSaturday
'        .ColTmpSunday = .ColorSunday
'
'        Dim m As Integer
'        For m = LBound(CalYear.Months) To UBound(CalYear.Months)
'            CalendarView_DrawMonth this, CalYear.Months(m)
'            nx = nx + 1
'            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
'        Next
'
'        If .HasJanNextYear Then
'            Dim CalNextYear As CalendarYear:  CalNextYear = New_CalendarYear(CalYear.Year + 1, 1, 1)
'            Dim JanNextYear As CalendarMonth: JanNextYear = New_CalendarMonth(CalNextYear, 1)
'            .ColTmpWeekday = .ColorLNWeekday
'            .ColTmpSaturday = .ColorLNSaturday
'            .ColTmpSunday = .ColorLNSunday
'            CalendarView_DrawMonth this, JanNextYear
'            nx = nx + 1
'            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
'        End If
'    End With
''Catch:
'End Sub
'
'Public Sub DrawDay(CalDay As CalendarDay)
''Try: On Error GoTo Catch
'    With this
'        Dim fc As Long: fc = .Canvas.ForeColor
'        Dim x As Double: x = .Canvas.CurrentX
'        Dim y As Double: y = .Canvas.CurrentY
'        Dim wd As VbDayOfWeek: wd = Weekday(CalDay.Date)
'        Dim c As Long: c = IIf(wd = vbSaturday, .ColTmpSaturday, IIf(wd = VbDayOfWeek.vbSunday, .ColTmpSunday, .ColTmpWeekday))
'
'        .Canvas.Line (x, y)-(x + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, BF
'
'        Select Case True
'        Case CalDay.MouseOver
'            .Canvas.DrawWidth = 2
'            c = RGB(255, 0, 0)
'        Case CalDay.FestivalIndex
'            .Canvas.DrawWidth = 2
'            c = .ColorFestivlDay
'        Case Else
'            .Canvas.DrawWidth = 1
'            c = .ColorNormalGrid
'        End Select
'        .Canvas.Line (x, y)-(x + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, B
'
'        .Canvas.CurrentX = x
'        .Canvas.CurrentY = y
'
'        Dim S As String
'        S = CStr(CalDay.Day) & " " & VbWeekDay_ToStr(wd, vbSunday, True)
'        If CalDay.FestivalIndex Then
'            S = S & " " & MDECalendar.ELegalFestivals_ToStr(CalDay.FestivalIndex)
'        Else
'            If wd = vbMonday Then
'                S = S & " " & "KW " & WeekOfYearISO(CalDay.Date)
'            End If
'        End If
'
'        Set .Canvas.Font = .FontDayNrName
'        If wd = vbSunday Then
'            .Canvas.ForeColor = RGB(255, 255, 255)
'        End If
'
'        .Canvas.Print S
'
'        .Canvas.CurrentX = x
'        .Canvas.CurrentY = y
'        .Canvas.ForeColor = fc
'    End With
''Catch:
'End Sub

