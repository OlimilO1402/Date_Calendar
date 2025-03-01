Attribute VB_Name = "MDECalendar"
Option Explicit

'for computing the legal or religious festivals / holidays of one year in every country of germany
'in this enum-series the exponent of the enum const matches the land-key (see AGS: https://de.wikipedia.org/wiki/Amtlicher_Gemeindeschl%C3%BCssel)

Public Enum EGermanLand
    SchleswigHolstein = &H2&        ' 2 ^ 01 ' Land
    Hamburg = &H4&                  ' 2 ^ 02 ' Freie und Hansestadt
    Niedersachsen = &H8&            ' 2 ^ 03 ' Land
    Bremen = &H10&                  ' 2 ^ 04 ' Freie und Hansestadt
    NordrheinWestfalen = &H20&      ' 2 ^ 05 ' Land
    Hessen = &H40&                  ' 2 ^ 06 ' Land
    Rheinlandpfalz = &H80&          ' 2 ^ 07 ' Land
    BadenWuerttemberg = &H100&      ' 2 ^ 08 ' Land
    Bayern = &H200&                 ' 2 ^ 09 ' Freistaat
    Bayern_Augsburg = &H201&        '
    Saarland = &H400&               ' 2 ^ 10 ' Land
    Berlin = &H800&                 ' 2 ^ 11 ' Stadtstaat
    Brandenburg = &H1000&           ' 2 ^ 12 ' Land
    MecklenburgVorpommern = &H2000& ' 2 ^ 13 ' Land
    Sachsen = &H4000&               ' 2 ^ 14 ' Freistaat
    SachsenAnhalt = &H8000&         ' 2 ^ 15 ' Land
    Thueringen = &H10000            ' 2 ^ 16 ' Freistaat
    AllLands = &H1FFFE
End Enum

Public Enum ELegalFestivals
    NewYear = 1               'Neujahr                   ' 1  01.01.
    HolyThreeKings            'Heilige Drei Könige       ' 2  06.01.
    InternationalWomensDay    'Internationaler Frauentag ' 3  08.03.
    GoodFriday                'Karfreitag                ' 4  2 days before easter sunday
    EasterSunday              'Ostersonntag              ' 5  calculate according to Gauss
    EasterMonday              'Ostermontag               ' 6  1 day after easter sunday
    LaborDay                  'Tag der Arbeit            ' 7  01.05.
    Mothersday                'Muttertag                 ' 8  2. sunday in may
    AscensionOfChrist         'Christihimmelfahrt        ' 9  10 days before pentecost sunday
    PentecostSunday           'Pfingstsonntag            '10  7 weeks = 49 days after easter sunday
    PentecostMonday           'Pfingstmontag             '11   1 days after pentecost sunday
    CorpusChristi             'Fronleichnam              '12  10 days after pentecost monday
    AugsburgPeaceFestival     'AugsburgerFriedensfest    'AugsburgerFriedensfest    '13  08.08. only in the city of Augsburg
    AssumptionDay             'MariaeHimmelfahrt         '14  15.08.
    WorldChildrensDay         'Weltkindertag             '15  20.09.
    DayOfGermanUnity          'TagDerDeutschenEinheit    '16  03.10. 'national festival
    ReformationDay            'Reformationstag           '17  31.10. 'protestantic festival
    AllSaintsDay              'Allerheiligen             '18  01.11.
    PrayerAndRepentanceDay    'BussUndBettag             '19  20.11 '10 days before first advent sunday
    FirstWarning              'ErsterAdvent              '20
    '                         '                          '21  24.12.
    ChristmasDayFirst = 22      'Weihnachtsfeiertag1 = 22  '22  25.12.
    ChristmasDaySecond        'Weihnachtsfeiertag2       '23  26.12.
    '                         '                          '24  31.12.
    MaxLegalFestivals
End Enum

Public Enum EContractFestivals
    ChristmasEve = 21         'Heiligabend = 21          '21  24.12. (according to job agreement maybe half holiday)
    NewYearsEve = 24          'Silvester = 24            '24  31.12. (according to job agreement maybe half holiday)
End Enum

Public Type LegalFestival
    Date     As Date
    Festival As ELegalFestivals
    Land     As EGermanLand
End Type

Public Type PersonalBirthday
    Name     As String
    BirthDay As Date ' the actual day in the year of birth
End Type

Public Type CalendarDay
    Day  As Integer
    Date As Date
    FestivalIndex As Integer '0 = no festivalday
    BirthDays()   As PersonalBirthday
    MouseOver As Boolean
    Selected  As Boolean
End Type

Public Type CalendarMonth
    Year   As Integer
    Month  As Integer
    Days() As CalendarDay
End Type

Public Type CalendarYear
    Year     As Integer
    Months() As CalendarMonth
    Fests()  As LegalFestival
End Type

Public Type Calendar
    LastYear As CalendarYear
    ThisYear As CalendarYear
    NextYear As CalendarYear
End Type

Public Type Margin
    Left   As Double
    Top    As Double
    Right  As Double
    Bottom As Double
End Type

Public Type CalendarView
    Canvas          As Control ' As Printer AndAs PictureBox
    HasDecLastYear  As Boolean
    HasJanNextYear  As Boolean
    HasMonthNames   As Boolean
    HasWeekDayNames As Boolean
    HasWeekNumbers  As Boolean
    MarginCal       As Margin
    MarginMon       As Margin 'not in use
    MarginDay       As Margin 'not in use
    ColorNormalGrid As Long 'grey
    ColorFestivlDay As Long 'purple
    ColorMouseOver  As Long 'yellow
    ColorSelected   As Long 'red
    ColorBirthDay   As Long 'green
    ColorWeekday    As Long 'white
    ColorSaturday   As Long 'lighlight blue
    ColorSunday     As Long 'light blue
    ColorLNWeekday  As Long 'white
    ColorLNSaturday As Long 'lightlight grey
    ColorLNSunday   As Long 'light grey
    ColTmpWeekday   As Long
    ColTmpSaturday  As Long
    ColTmpSunday    As Long
    FontMonthName   As StdFont
    FontDayNrName   As StdFont
    FontWeekNr      As StdFont
    TmpDayWidth     As Double
    TmpDayHeight    As Double
End Type

' v ############################## v '       the legal and religious holidays / festivals       ' v ############################## v '
Private Function New_LegalFestival(ByVal aDate As Date, ByVal aFest As ELegalFestivals, ByVal aLand As EGermanLand) As LegalFestival
    With New_LegalFestival:  .Date = aDate: .Festival = aFest: .Land = aLand: End With
End Function

Public Function ELegalFestivals_ToStr(ByVal e As ELegalFestivals) As String
    Dim S As String
    Select Case e
    Case ELegalFestivals.NewYear:                   S = "Neujahr"           ' 1  01.01.
    Case ELegalFestivals.HolyThreeKings:            S = "Heilige 3 Könige"  ' 2  06.01.
    Case ELegalFestivals.InternationalWomensDay:    S = "Internat.Frauent"  ' 3  08.03.
    Case ELegalFestivals.GoodFriday:                S = "Karfreitag"        ' 4  2 days before easter-sunday
    Case ELegalFestivals.EasterSunday:              S = "Ostersonntag"      ' 5  calculate accoding to Gauss"
    Case ELegalFestivals.EasterMonday:              S = "Ostermontag"       ' 6  1 day after easter-sunday
    Case ELegalFestivals.LaborDay:                  S = "Tag Der Arbeit"    ' 7  01.05."
    Case ELegalFestivals.Mothersday:                S = "Muttertag"         ' 8  2. sunday in may
    Case ELegalFestivals.AscensionOfChrist:         S = "Christi Himmelf."  ' 9  10 days before Pentecost-sunday"
    Case ELegalFestivals.PentecostSunday:           S = "Pfingstsonntag"    '10  7 weeks = 49 days after easter-sunday"
    Case ELegalFestivals.PentecostMonday:           S = "Pfingstmontag"     '11  1 day after Pentecost-sunday"
    Case ELegalFestivals.CorpusChristi:             S = "Fronleichnam"      '12  10 days after Pentecost-monday"
    Case ELegalFestivals.AugsburgPeaceFestival:     S = "Augsbg.Friedensf." '13  08.08."
    Case ELegalFestivals.AssumptionDay:             S = "Mariä Himmelf."    '14  15.08."
    Case ELegalFestivals.WorldChildrensDay:             S = "Weltkindertag"     '15  20.09."
    Case ELegalFestivals.DayOfGermanUnity:          S = "Tag d.Dt.Einheit"  '16  03.10."
    Case ELegalFestivals.ReformationDay:           S = "Reformationstag"   '17  31.10."
    Case ELegalFestivals.AllSaintsDay:             S = "Allerheiligen"     '18
    Case ELegalFestivals.PrayerAndRepentanceDay:    S = "Buß- & Bettag"     '19  20.11
    
    Case ELegalFestivals.FirstWarning:              S = "Erster Advent"     '19  20.11
    
    Case EContractFestivals.ChristmasEve:           S = "Heiligabend"       '20  24.12.
    
    Case ELegalFestivals.ChristmasDayFirst:         S = "1. Weihnachtsf."   '21  25.12.
    Case ELegalFestivals.ChristmasDaySecond:        S = "2. Weihnachtsf."   '22  26.12.
    
    Case EContractFestivals.NewYearsEve:            S = "Silvester"         '23  31.12.
    End Select
    ELegalFestivals_ToStr = S
End Function

Public Function GetFestivals(ByVal Year As Integer) As LegalFestival()
    'returns an array of festivals for a given year
    ReDim Fests(0 To ELegalFestivals.MaxLegalFestivals) As LegalFestival
    Dim i As Long
    Dim EasterSunday As Date: EasterSunday = MTime.OsternShort2(Year)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 1, 1), ELegalFestivals.NewYear, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 1, 6), ELegalFestivals.HolyThreeKings, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.SachsenAnhalt)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 3, 8), ELegalFestivals.InternationalWomensDay, EGermanLand.Berlin Or EGermanLand.MecklenburgVorpommern)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday - 2, ELegalFestivals.GoodFriday, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday, ELegalFestivals.EasterSunday, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 1, ELegalFestivals.EasterMonday, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 5, 1), ELegalFestivals.LaborDay, EGermanLand.AllLands)
    
    Dim Mothersday As Date: Mothersday = MTime.Mothersday(Year)
    i = i + 1:    Fests(i) = New_LegalFestival(Mothersday, ELegalFestivals.Mothersday, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 39, ELegalFestivals.AscensionOfChrist, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 49, ELegalFestivals.PentecostSunday, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 50, ELegalFestivals.PentecostMonday, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 60, ELegalFestivals.CorpusChristi, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.Hessen Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland Or EGermanLand.Sachsen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 8), ELegalFestivals.AugsburgPeaceFestival, EGermanLand.Bayern_Augsburg)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 15), ELegalFestivals.AssumptionDay, EGermanLand.Saarland Or EGermanLand.Bayern)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 9, 20), ELegalFestivals.WorldChildrensDay, EGermanLand.Thueringen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 3), ELegalFestivals.DayOfGermanUnity, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 31), ELegalFestivals.ReformationDay, EGermanLand.Brandenburg Or EGermanLand.Bremen Or EGermanLand.Hamburg Or EGermanLand.MecklenburgVorpommern Or EGermanLand.Niedersachsen Or EGermanLand.Sachsen Or EGermanLand.SachsenAnhalt Or EGermanLand.SchleswigHolstein Or EGermanLand.Thueringen)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 11, 1), ELegalFestivals.AllSaintsDay, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland)
    
    Dim AdvSund1 As Date: AdvSund1 = AdventSunday1(Year)
    'Der Buß- und Bettag findet jedes Jahr am Mittwoch vor Totensonntag und damit genau elf Tage vor dem ersten Adventssonntag statt
    i = i + 1:    Fests(i) = New_LegalFestival(AdvSund1 - 11, ELegalFestivals.PrayerAndRepentanceDay, EGermanLand.Sachsen)
    
    i = i + 1:    Fests(i) = New_LegalFestival(AdvSund1, ELegalFestivals.FirstWarning, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 24), EContractFestivals.ChristmasEve, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 25), ELegalFestivals.ChristmasDayFirst, EGermanLand.AllLands)
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 26), ELegalFestivals.ChristmasDaySecond, EGermanLand.AllLands)
    
    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 31), EContractFestivals.NewYearsEve, EGermanLand.AllLands)
    
    GetFestivals = Fests
End Function

Public Property Get Festivals_Index(this() As LegalFestival, ByVal aDate As Date) As Integer
    'returns the index in the array if aDate is a legal, religious or festival holiday otherwise 0
    Dim i As Integer
    For i = LBound(this) To UBound(this)
        If this(i).Date = aDate Then
            Festivals_Index = i
            Exit Property
        End If
    Next
End Property
' ^ ############################## ^ '       the legal and religious holidays / festivals       ' ^ ############################## ^ '

'fuck how to do this properly?
Public Function New_CalendarYear(ByVal Year As Integer, _
                                 Optional ByVal StartMonth As Integer = 1, _
                                 Optional ByVal EndMonth As Integer = 12, _
                                 Optional ByVal includeLastDec As Boolean = False, _
                                 Optional ByVal includeNextJan As Boolean = False) As CalendarYear
    Dim Y As CalendarYear
    Y.Year = Year
    Y.Fests = GetFestivals(Year)
    StartMonth = IIf(0 < StartMonth And StartMonth <= 12, StartMonth, 1)
    EndMonth = IIf(StartMonth <= EndMonth And EndMonth <= 12, EndMonth, 12)
    ReDim Y.Months(StartMonth To EndMonth)
    Dim m As Integer
    For m = StartMonth To EndMonth
        Y.Months(m) = New_CalendarMonth(Y, m)
    Next
    New_CalendarYear = Y
End Function

Public Function New_CalendarMonth(CalYear As CalendarYear, ByVal Month As Integer) As CalendarMonth
    With New_CalendarMonth
        .Year = CalYear.Year
        .Month = Month
        Dim mds As Integer: mds = DaysInMonth(.Year, Month)
        ReDim .Days(1 To mds)
        Dim d As Integer
        For d = 1 To mds
            .Days(d) = New_CalendarDay(CalYear, Month, d)
        Next
    End With
End Function

Public Function New_CalendarDay(CalYear As CalendarYear, ByVal Month As Integer, ByVal Day As Integer) As CalendarDay
    With New_CalendarDay
        .Day = Day
        .Date = DateSerial(CalYear.Year, Month, Day)
        .FestivalIndex = Festivals_Index(CalYear.Fests, .Date)
    End With
End Function

Public Function New_Calendar(LastY As CalendarYear, ThisY As CalendarYear, NextY As CalendarYear) As Calendar
    With New_Calendar
        .LastYear = LastY
        .ThisYear = ThisY
        .NextYear = NextY
    End With
End Function

Public Function New_StdFont(ByVal FontName As String, Optional ByVal Size As Single = 10, Optional ByVal IsBold As Boolean = False, Optional ByVal IsItalic As Boolean = False, Optional ByVal IsStrikedthrough As Boolean = False, Optional ByVal IsUnderlined As Boolean = False) As StdFont
    Set New_StdFont = New StdFont: New_StdFont.Name = FontName
    With New_StdFont
        .Size = Size
        .Bold = IsBold
        .Italic = IsItalic
        .Strikethrough = IsStrikedthrough
        .Underline = IsUnderlined
        '.Weight
        '.Charset
    End With
End Function

Public Function StdFont_Clone(ByVal other As StdFont) As StdFont
    Set StdFont_Clone = New StdFont
    With StdFont_Clone
        .Name = other.Name
        .Size = other.Size
        .Bold = other.Bold
        .Italic = other.Italic
        .Weight = other.Weight
        .Charset = other.Charset
        .Underline = other.Underline
        .Strikethrough = other.Strikethrough
    End With
End Function

Public Function New_Margin(MargLeft_Or_LTRB, Optional MargTop, Optional MargRight, Optional MargBottom) As Margin
    With New_Margin
        .Left = CDbl(MargLeft_Or_LTRB)
        .Top = IIf(IsMissing(MargTop), .Left, MargTop)
        .Right = IIf(IsMissing(MargRight), .Left, MargRight)
        .Bottom = IIf(IsMissing(MargBottom), .Left, MargBottom)
    End With
End Function

Public Function New_CalendarView(Canvas As PictureBox) As CalendarView
    With New_CalendarView
        Set .Canvas = Canvas
        .ColorNormalGrid = RGB(240, 240, 240)
        .ColorFestivlDay = RGB(222, 141, 245)
        .ColorMouseOver = RGB(255, 255, 0)
        .ColorSelected = RGB(255, 0, 0)
        .ColorBirthDay = RGB(0, 255, 0)
        .ColorWeekday = RGB(255, 255, 255)
        .ColorSaturday = RGB(230, 244, 253)
        .ColorSunday = RGB(137, 189, 226)
        .ColorLNWeekday = RGB(255, 255, 255)
        .ColorLNSaturday = RGB(200, 202, 201)
        .ColorLNSunday = RGB(157, 157, 157)
        Set .FontDayNrName = New_StdFont("Segoe UI") '("Comic Sans MS")
        Set .FontMonthName = New_StdFont("Segoe Print", 10, True) '("Comic Sans MS")
        'Set .FontMonthName = New_StdFont("Comic Sans MS")
        Set .FontWeekNr = New_StdFont("Segoe UI") '("Comic Sans MS")
        .HasDecLastYear = True
        .HasJanNextYear = True
        .HasMonthNames = True
        .HasWeekDayNames = True
        .HasWeekNumbers = True
        .MarginCal = New_Margin(10)
        '.MarginCalLeft = 10 'px
        '.MarginCalTop = 10 'px
        '.MarginCalRight = 10 'px
        '.MarginCalBottom = 10 'px
    End With
End Function

Public Function CalendarView_Clone(other As CalendarView) As CalendarView
    With CalendarView_Clone
        Set .Canvas = other.Canvas
        .HasDecLastYear = other.HasDecLastYear
        .HasJanNextYear = other.HasJanNextYear
        .HasMonthNames = other.HasMonthNames
        .HasWeekDayNames = other.HasWeekDayNames
        .HasWeekNumbers = other.HasWeekNumbers
        .MarginCal = other.MarginCal
        .MarginMon = other.MarginMon
        .MarginDay = other.MarginDay
        .ColorNormalGrid = other.ColorNormalGrid
        .ColorFestivlDay = other.ColorFestivlDay
        .ColorMouseOver = other.ColorMouseOver
        .ColorWeekday = other.ColorWeekday
        .ColorSaturday = other.ColorSaturday
        .ColorSunday = other.ColorSunday
        .ColorLNWeekday = other.ColorLNWeekday
        .ColorLNSaturday = other.ColorLNSaturday
        .ColorLNSunday = other.ColorLNSunday
        .ColTmpWeekday = other.ColTmpWeekday
        .ColTmpSaturday = other.ColTmpSaturday
        .ColTmpSunday = other.ColTmpSunday
        Set .FontMonthName = StdFont_Clone(other.FontMonthName)
        Set .FontDayNrName = StdFont_Clone(other.FontDayNrName)
        Set .FontWeekNr = StdFont_Clone(other.FontWeekNr)
        .TmpDayWidth = other.TmpDayWidth
        .TmpDayHeight = other.TmpDayHeight
    End With
End Function

Public Function CalendarView_Dispose(this As CalendarView)
    With this
        Set .Canvas = Nothing
        Set .FontDayNrName = Nothing
        Set .FontMonthName = Nothing
        Set .FontWeekNr = Nothing
    End With
End Function

Public Sub CalendarDay_MouseOut(this As CalendarDay, CalYear As CalendarYear)
    With this
        Dim m As Integer: m = Month(.Date)
        If m = 0 Then Exit Sub
        If .Day = 0 Then Exit Sub
        CalYear.Months(Month(.Date)).Days(.Day).MouseOver = False
    End With
End Sub

Public Function CalendarView_CalendarDayFromMouseCoords(this As CalendarView, CalYear As CalendarYear, ByVal MouseX As Single, ByVal MouseY As Single) As CalendarDay
    With this
        MouseX = MouseX - .MarginCal.Left
        MouseY = MouseY - .MarginCal.Top
        Dim m As Integer: m = CInt(MouseX \ .TmpDayWidth)         ' x-axis
        Dim d As Integer: d = CInt(MouseY \ .TmpDayHeight) '- 1  ' y-axis
        Dim dly As Integer: dly = IIf(.HasDecLastYear, 1, 0)
        Dim jny As Integer: jny = IIf(.HasJanNextYear, 1, 0)
        If 0 <= m And m <= UBound(CalYear.Months) + 1 + dly + jny Then
            If 0 < d And d < MTime.DaysInMonth(CalYear.Year, m) Then
                CalYear.Months(m).Days(d).MouseOver = True
                CalendarView_CalendarDayFromMouseCoords = CalYear.Months(m).Days(d)
            End If
        End If
    End With
End Function

Public Property Get CalendarView_DayWidth(this As CalendarView, CalYear As CalendarYear) As Double
    With this
        Dim n As Long: n = UBound(CalYear.Months) - LBound(CalYear.Months) + 1 + IIf(.HasDecLastYear, 1, 0) + IIf(.HasJanNextYear, 1, 0)
        CalendarView_DayWidth = (.Canvas.ScaleWidth - .MarginCal.Left - .MarginCal.Right) / n
    End With
End Property

Public Property Get CalendarView_DayHeight(this As CalendarView) As Double
    With this
        Dim n As Double: n = 32
        CalendarView_DayHeight = (.Canvas.ScaleHeight - .MarginCal.Top - .MarginCal.Bottom - IIf(.HasMonthNames, .FontMonthName.Size, 0)) / n
    End With
End Property

Public Sub CalendarView_DrawYear(this As CalendarView, CalYear As CalendarYear)
'Try: On Error GoTo Catch
    With this
        Dim nx As Integer
        .Canvas.CurrentX = .MarginCal.Left
        .Canvas.CurrentY = .MarginCal.Top
        
        .TmpDayWidth = CalendarView_DayWidth(this, CalYear)
        .TmpDayHeight = CalendarView_DayHeight(this)
        
        If .HasDecLastYear Then
            Dim CalLastYear As CalendarYear:  CalLastYear = New_CalendarYear(CalYear.Year - 1, 12, 12)
            Dim DecLastYear As CalendarMonth: DecLastYear = New_CalendarMonth(CalLastYear, 12)
            .ColTmpWeekday = .ColorLNWeekday
            .ColTmpSaturday = .ColorLNSaturday
            .ColTmpSunday = .ColorLNSunday
            CalendarView_DrawMonth this, DecLastYear
            nx = nx + 1
            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
        End If
        
        .ColTmpWeekday = .ColorWeekday
        .ColTmpSaturday = .ColorSaturday
        .ColTmpSunday = .ColorSunday

        Dim m As Integer
        For m = LBound(CalYear.Months) To UBound(CalYear.Months)
            CalendarView_DrawMonth this, CalYear.Months(m)
            nx = nx + 1
            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
        Next
        
        If .HasJanNextYear Then
            Dim CalNextYear As CalendarYear:  CalNextYear = New_CalendarYear(CalYear.Year + 1, 1, 1)
            Dim JanNextYear As CalendarMonth: JanNextYear = New_CalendarMonth(CalNextYear, 1)
            .ColTmpWeekday = .ColorLNWeekday
            .ColTmpSaturday = .ColorLNSaturday
            .ColTmpSunday = .ColorLNSunday
            CalendarView_DrawMonth this, JanNextYear
            nx = nx + 1
            .Canvas.CurrentX = .MarginCal.Left + nx * .TmpDayWidth
        End If
    End With
'Catch:
End Sub

Public Sub CalendarView_DrawMonth(this As CalendarView, CalMonth As CalendarMonth)
'Try: On Error GoTo Catch
    With this
        Dim X As Double: X = .Canvas.CurrentX
        Dim Y As Double: Y = .MarginCal.Top
        Dim ny As Integer
        If .HasMonthNames Then
            Set .Canvas.Font = .FontMonthName
            Dim S As String: S = MonthName(CalMonth.Month) & " '" & Right(CStr(CalMonth.Year), 2)
            .Canvas.Print S
            ny = ny + 1
            .Canvas.CurrentY = .MarginCal.Top + ny * .TmpDayHeight
        End If
        .Canvas.CurrentX = X
        Dim d As Integer
        Dim L As Integer: L = LBound(CalMonth.Days)
        Dim u As Integer: u = UBound(CalMonth.Days)
        For d = L To u
            CalendarView_DrawDay this, CalMonth.Days(d)
            ny = ny + 1
            .Canvas.CurrentY = .MarginCal.Top + ny * .TmpDayHeight
        Next
        .Canvas.CurrentY = X
        .Canvas.CurrentY = Y
    End With
'Catch:
End Sub

Public Sub CalendarView_DrawDay(this As CalendarView, CalDay As CalendarDay)
'Try: On Error GoTo Catch
    With this
        Dim fc As Long: fc = .Canvas.ForeColor
        Dim X As Double: X = .Canvas.CurrentX
        Dim Y As Double: Y = .Canvas.CurrentY
        Dim wd As VbDayOfWeek: wd = Weekday(CalDay.Date)
        Dim c As Long: c = IIf(wd = vbSaturday, .ColTmpSaturday, IIf(wd = VbDayOfWeek.vbSunday, .ColTmpSunday, .ColTmpWeekday))
        
        .Canvas.Line (X, Y)-(X + .TmpDayWidth - 1, Y + .TmpDayHeight - 1), c, BF
        
        Select Case True
        Case CalDay.MouseOver
            .Canvas.DrawWidth = 2
            c = RGB(255, 0, 0)
        Case CalDay.FestivalIndex
            .Canvas.DrawWidth = 2
            c = .ColorFestivlDay
        Case Else
            .Canvas.DrawWidth = 1
            c = .ColorNormalGrid
        End Select
        .Canvas.Line (X, Y)-(X + .TmpDayWidth - 1, Y + .TmpDayHeight - 1), c, B
        
        .Canvas.CurrentX = X
        .Canvas.CurrentY = Y
        
        Dim S As String
        S = CStr(CalDay.Day) & " " & VbWeekDay_ToStr(wd, vbSunday, True)
        If CalDay.FestivalIndex Then
            S = S & " " & MDECalendar.ELegalFestivals_ToStr(CalDay.FestivalIndex)
        Else
            If wd = vbMonday Then
                S = S & " " & "KW " & WeekOfYearISO(CalDay.Date)
            End If
        End If
        
        Set .Canvas.Font = .FontDayNrName
        If wd = vbSunday Then
            .Canvas.ForeColor = RGB(255, 255, 255)
        End If
        
        .Canvas.Print S
        
        .Canvas.CurrentX = X
        .Canvas.CurrentY = Y
        .Canvas.ForeColor = fc
    End With
'Catch:
End Sub
