Attribute VB_Name = "MDECalendar"
Option Explicit

'for computing the legal or religious festivals / holidays of one year in every country of germany
'in this enum-series the exponent of the enum const matches the land-key (see AGS: https://de.wikipedia.org/wiki/Amtlicher_Gemeindeschl%C3%BCssel)
'
'Public Enum EGermanLand
'    SchleswigHolstein = &H2&        ' 2 ^ 01 ' Land
'    Hamburg = &H4&                  ' 2 ^ 02 ' Freie und Hansestadt
'    Niedersachsen = &H8&            ' 2 ^ 03 ' Land
'    Bremen = &H10&                  ' 2 ^ 04 ' Freie und Hansestadt
'    NordrheinWestfalen = &H20&      ' 2 ^ 05 ' Land
'    Hessen = &H40&                  ' 2 ^ 06 ' Land
'    Rheinlandpfalz = &H80&          ' 2 ^ 07 ' Land
'    BadenWuerttemberg = &H100&      ' 2 ^ 08 ' Land
'    Bayern = &H200&                 ' 2 ^ 09 ' Freistaat
'    Bayern_Augsburg = &H201&        '
'    Saarland = &H400&               ' 2 ^ 10 ' Land
'    Berlin = &H800&                 ' 2 ^ 11 ' Stadtstaat
'    Brandenburg = &H1000&           ' 2 ^ 12 ' Land
'    MecklenburgVorpommern = &H2000& ' 2 ^ 13 ' Land
'    Sachsen = &H4000&               ' 2 ^ 14 ' Freistaat
'    SachsenAnhalt = &H8000&         ' 2 ^ 15 ' Land
'    Thueringen = &H10000            ' 2 ^ 16 ' Freistaat
'    AllLands = &H1FFFE
'End Enum

Public Enum ELegalFestivals
    NewYearsDay = 1             'Neujahr                   ' 1  01.01.
    Epiphany = 2                'Heilige Drei Kˆnige       ' 2  06.01.
    'ValentinesDay = 3           'Valentinstag              ' 3  14.02.
    'AshWednesday = 4            'Aschermittwoch            ' 4  46 days before easter sunday
    InternationalWomensDay = 5  'Internationaler Frauentag ' 5  08.03.
    GoodFriday                  'Karfreitag                ' 6  2 days before easter sunday
    EasterSunday                'Ostersonntag              ' 7  calculate according to Gauss
    EasterMonday                'Ostermontag               ' 8  1 day after easter sunday
    LaborDay                    'Tag der Arbeit            ' 9  01.05.
    'Mothersday = 10             'Muttertag                 '10  2. sunday in may
    AscensionOfChrist = 11      'Christihimmelfahrt        '11  10 days before pentecost sunday
    WhitSunday                  'Pfingstsonntag            '12  7 weeks = 49 days after easter sunday (aka Pentecost sunday)
    WhitMonday                  'Pfingstmontag             '13   1 days after pentecost sunday (aka Pentecost monday)
    CorpusChristi               'Fronleichnam              '14  10 days after pentecost monday
    AugsburgPeaceFestival       'AugsburgerFriedensfest    '15  AugsburgerFriedensfest    '13  08.08. only in the city of Augsburg
    AssumptionDay               'MariaeHimmelfahrt         '16  15.08.
    WorldChildrensDay           'Weltkindertag             '17  20.09.
    DayOfGermanUnity            'TagDerDeutschenEinheit    '18  03.10. 'national festival
    ReformationDay              'Reformationstag           '19  31.10. 'protestantic festival
    AllSaintsDay                'Allerheiligen             '20  01.11.
    'StMartin                    'St.Martin                 '21
    PrayerAndRepentanceDay      'BussUndBettag             '22  20.11 '10 days before first advent sunday
    'ChristmasEve = 24           'Heiligabend               '24  24.12.
    ChristmasDayFirst = 25      'Weihnachtsfeiertag1 = 22  '25  25.12.
    ChristmasDaySecond = 26     'Weihnachtsfeiertag2       '26  26.12.
    'NewYearsEve                'Silvester                 '27  31.12.
    MaxFestivals
End Enum

Public Enum EAdditionalFestivals
    ValentinesDay = 3           'Valentinstag              ' 3  14.02.
    AshWednesday = 4            'Aschermittwoch            ' 4  46 days before easter sunday
    Mothersday = 10             'Muttertag                 '10  2. sunday in may
    StMartin = 21               'StMartin                  '21
    FirstAdvent = 23            'ErsterAdvent              '23
    ChristmasEve = 24           'Heiligabend               '24  24.12. (according to job agreement maybe half holiday)
    NewYearsEve = 27            'Silvester                 '27  31.12. (according to job agreement maybe half holiday)
End Enum

'Public Type LegalFestival
'    Date     As Date
'    Festival As ELegalFestivals
'    Land     As EGermanLand
'End Type

Public Enum EEventType
    'what else ... ?
    BirthDay '(icon candles on a bid-cake)
    Marriage '(icon crossed wedding rings)
    'SpecialEvent '(icon champaign-glasses) like Firm-Jubil‰um
    Anniversary   'Firmenjubil‰um
    Holiday       'Urlaub von ... bis
    BusinessEvent 'Messe  von ... bis
    'what else ... ?
End Enum
'
'Public Type PersonalEvent
'    EventType As EPersonalEventType
'    Name      As String
'    Date      As Date
'End Type

'Public Type PersonalBirthday
'    Name     As String
'    BirthDay As Date ' the actual day in the year of birth
'End Type

'Public Type CalendarDay
'    Day  As Integer
'    Date As Date
'    FestivalIndex As Integer '0 = no festivalday
'    BirthDays()   As PersonalBirthday
'    MouseOver As Boolean
'    Selected  As Boolean
'End Type

'Public Type CalendarMonth
'    Year   As Integer
'    Month  As Integer
'    Days() As CalendarDay
'End Type

'Public Type CalendarYear
'    Year     As Integer
'    Months() As CalendarMonth
'    Fests()  As LegalFestival
'End Type

'Public Type Calendar
'    LastYear As CalendarYear
'    ThisYear As CalendarYear
'    NextYear As CalendarYear
'End Type

Public Type Rectangle
    X0 As Double
    Y0 As Double
    X1 As Double
    Y1 As Double
End Type

Public Type Margin
    Left   As Double
    Top    As Double
    Right  As Double
    Bottom As Double
End Type

'Public Type CalendarView
'    Canvas          As Control ' As Printer AndAs PictureBox
'    HasDecLastYear  As Boolean
'    HasJanNextYear  As Boolean
'    HasMonthNames   As Boolean
'    HasWeekDayNames As Boolean
'    HasWeekNumbers  As Boolean
'    MarginCal       As Margin
'    MarginMon       As Margin 'not in use
'    MarginDay       As Margin 'not in use
'    ColorNormalGrid As Long 'grey
'    ColorFestivlDay As Long 'purple
'    ColorMouseOver  As Long 'yellow
'    ColorSelected   As Long 'red
'    ColorBirthDay   As Long 'green
'    ColorWeekday    As Long 'white
'    ColorSaturday   As Long 'lighlight blue
'    ColorSunday     As Long 'light blue
'    ColorLNWeekday  As Long 'white
'    ColorLNSaturday As Long 'lightlight grey
'    ColorLNSunday   As Long 'light grey
'    ColTmpWeekday   As Long
'    ColTmpSaturday  As Long
'    ColTmpSunday    As Long
'    FontMonthName   As StdFont
'    FontDayNrName   As StdFont
'    FontWeekNr      As StdFont
'    TmpDayWidth     As Double
'    TmpDayHeight    As Double
'End Type

' v ############################## v '       the legal and religious holidays / festivals       ' v ############################## v '
'Private Function New_LegalFestival(ByVal aDate As Date, ByVal aFest As ELegalFestivals, ByVal aLand As EGermanLand) As LegalFestival
'    With New_LegalFestival:  .Date = aDate: .Festival = aFest: .Land = aLand: End With
'End Function

Public Function Festivals_ToStr(ByVal e As Long) As String

    Dim S As String
    Select Case e
    Case ELegalFestivals.NewYearsDay:               S = "Neujahr"                    ' 1  01.01.
    Case ELegalFestivals.Epiphany:                  S = "Heilige Drei Kˆnige"        ' 2  06.01.
    
    Case EAdditionalFestivals.ValentinesDay = 3:    S = "Valentinstag"               ' 3  14.02.
    Case EAdditionalFestivals.AshWednesday = 4:     S = "Aschermittwoch"             ' 4  46 days before easter sunday
    
    Case ELegalFestivals.InternationalWomensDay:    S = "Internationaler Frauentag"  ' 5  08.03.
    Case ELegalFestivals.GoodFriday:                S = "Karfreitag"                 ' 6  2 days before easter-sunday
    Case ELegalFestivals.EasterSunday:              S = "Ostersonntag"               ' 7  calculate accoding to Gauss"
    Case ELegalFestivals.EasterMonday:              S = "Ostermontag"                ' 8  1 day after easter-sunday
    Case ELegalFestivals.LaborDay:                  S = "Tag Der Arbeit"             ' 9  01.05."
    
    Case EAdditionalFestivals.Mothersday:           S = "Muttertag"                  '10  2. sunday in may
    
    Case ELegalFestivals.AscensionOfChrist:         S = "Christi Himmelfahrt"        '11  10 days before Pentecost-sunday"
    Case ELegalFestivals.WhitSunday:                S = "Pfingstsonntag"             '12  7 weeks = 49 days after easter-sunday"
    Case ELegalFestivals.WhitMonday:                S = "Pfingstmontag"              '13  1 day after Pentecost-sunday"
    Case ELegalFestivals.CorpusChristi:             S = "Fronleichnam"               '14  10 days after Pentecost-monday"
    Case ELegalFestivals.AugsburgPeaceFestival:     S = "Augsburger Friedensfest"    '15  08.08."
    Case ELegalFestivals.AssumptionDay:             S = "Mari‰ Himmelfahrt"          '16  15.08."
    Case ELegalFestivals.WorldChildrensDay:         S = "Weltkindertag"              '17  20.09."
    Case ELegalFestivals.DayOfGermanUnity:          S = "Tag der Deutschen Einheit"  '18  03.10."
    Case ELegalFestivals.ReformationDay:            S = "Reformationstag"            '19  31.10."
    Case ELegalFestivals.AllSaintsDay:              S = "Allerheiligen"              '20
    
    Case EAdditionalFestivals.StMartin:             S = "St.Martin, Faschingsanfang" '21
    
    Case ELegalFestivals.PrayerAndRepentanceDay:    S = "Buﬂ- & Bettag"              '22  20.11
    
    Case EAdditionalFestivals.FirstAdvent:          S = "Erster Advent"              '23  29.11
    Case EAdditionalFestivals.ChristmasEve:         S = "Heiligabend"                '24  24.12.
    
    Case ELegalFestivals.ChristmasDayFirst:         S = "1. Weihnachtsfeiertag"            '25  25.12.
    Case ELegalFestivals.ChristmasDaySecond:        S = "2. Weihnachtsfeiertag"            '26  26.12.
    
    Case EAdditionalFestivals.NewYearsEve:          S = "Silvester"                  '27  31.12.
    End Select
    Festivals_ToStr = S
End Function

Public Function GetFestivals(ByVal Year As Integer) As Collection 'Of FestivalDay 'LegalFestival 'As LegalFestival()
    
    'returns a collection of festivals for a given year
    
    Dim Festivals As Collection: Set Festivals = New Collection
    
    Dim EasterSunday As Date: EasterSunday = MTime.OsternShort2(Year)
    Dim Mothersday   As Date:   Mothersday = MTime.Mothersday(Year)
    Dim AdventSunday As Date: AdventSunday = MTime.AdventSunday1(Year)
    
    Dim fd As FestivalDay
    
    Set fd = MNew.FestivalDay(DateSerial(Year, 1, 1), ELegalFestivals.NewYearsDay):            Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 1, 6), ELegalFestivals.Epiphany):               Col_AddOrGet Festivals, fd ', fd.Key
    
    Set fd = MNew.FestivalDay(EasterSunday - 46, EAdditionalFestivals.AshWednesday):           Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 2, 14), EAdditionalFestivals.ValentinesDay):    Col_AddOrGet Festivals, fd ', fd.Key
    'in 2024 valentinesday and ashwednesday was the same day, -> we add ashwednesday first
    'Debug.Print fd.FestivalDate & " " & fd.WeekdayName & " " & fd.Name
    
    Set fd = MNew.FestivalDay(DateSerial(Year, 3, 8), ELegalFestivals.InternationalWomensDay): Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(EasterSunday - 2, ELegalFestivals.GoodFriday):                   Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(EasterSunday, ELegalFestivals.EasterSunday):                     Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(EasterSunday + 1, ELegalFestivals.EasterMonday):                 Col_AddOrGet Festivals, fd ', fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 5, 1), ELegalFestivals.LaborDay):               Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(Mothersday, EAdditionalFestivals.Mothersday):                    Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(EasterSunday + 39, ELegalFestivals.AscensionOfChrist):           Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(EasterSunday + 49, ELegalFestivals.WhitSunday):                  Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(EasterSunday + 50, ELegalFestivals.WhitMonday):                  Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(EasterSunday + 60, ELegalFestivals.CorpusChristi):               Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 8, 8), ELegalFestivals.AugsburgPeaceFestival):  Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 8, 15), ELegalFestivals.AssumptionDay):         Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 9, 20), ELegalFestivals.WorldChildrensDay):     Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 10, 3), ELegalFestivals.DayOfGermanUnity):      Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 10, 31), ELegalFestivals.ReformationDay):       Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 11, 1), ELegalFestivals.AllSaintsDay):          Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(DateSerial(Year, 11, 11), EAdditionalFestivals.StMartin):        Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(AdventSunday - 11, ELegalFestivals.PrayerAndRepentanceDay):      Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(AdventSunday, EAdditionalFestivals.FirstAdvent):                 Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 12, 24), EAdditionalFestivals.ChristmasEve):    Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(DateSerial(Year, 12, 25), ELegalFestivals.ChristmasDayFirst):    Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    Set fd = MNew.FestivalDay(DateSerial(Year, 12, 26), ELegalFestivals.ChristmasDaySecond):   Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key
    
    Set fd = MNew.FestivalDay(DateSerial(Year, 12, 31), EAdditionalFestivals.NewYearsEve):     Col_AddOrGet Festivals, fd 'Festivals.Add fd, fd.Key

    Set GetFestivals = Festivals

'    i = i + 1:    Fests(i) = MNew.FestivalDay(DateSerial(Year, 1, 1), ELegalFestivals.NewYear, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 1, 6), ELegalFestivals.HolyThreeKings, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.SachsenAnhalt)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 3, 8), ELegalFestivals.InternationalWomensDay, EGermanLand.Berlin Or EGermanLand.MecklenburgVorpommern)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday - 2, ELegalFestivals.GoodFriday, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday, ELegalFestivals.EasterSunday, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 1, ELegalFestivals.EasterMonday, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 5, 1), ELegalFestivals.LaborDay, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(Mothersday, ELegalFestivals.Mothersday, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 39, ELegalFestivals.AscensionOfChrist, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 49, ELegalFestivals.PentecostSunday, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 50, ELegalFestivals.PentecostMonday, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(EasterSunday + 60, ELegalFestivals.CorpusChristi, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.Hessen Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland Or EGermanLand.Sachsen)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 8), ELegalFestivals.AugsburgPeaceFestival, EGermanLand.Bayern_Augsburg)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 8, 15), ELegalFestivals.AssumptionDay, EGermanLand.Saarland Or EGermanLand.Bayern)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 9, 20), ELegalFestivals.WorldChildrensDay, EGermanLand.Thueringen)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 3), ELegalFestivals.DayOfGermanUnity, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 10, 31), ELegalFestivals.ReformationDay, EGermanLand.Brandenburg Or EGermanLand.Bremen Or EGermanLand.Hamburg Or EGermanLand.MecklenburgVorpommern Or EGermanLand.Niedersachsen Or EGermanLand.Sachsen Or EGermanLand.SachsenAnhalt Or EGermanLand.SchleswigHolstein Or EGermanLand.Thueringen)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 11, 1), ELegalFestivals.AllSaintsDay, EGermanLand.BadenWuerttemberg Or EGermanLand.Bayern Or EGermanLand.NordrheinWestfalen Or EGermanLand.Rheinlandpfalz Or EGermanLand.Saarland)
'
'    'Der Buﬂ- und Bettag findet jedes Jahr am Mittwoch vor Totensonntag und damit genau elf Tage vor dem ersten Adventssonntag statt
'    i = i + 1:    Fests(i) = New_LegalFestival(AdvSund1 - 11, ELegalFestivals.PrayerAndRepentanceDay, EGermanLand.Sachsen)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(AdvSund1, ELegalFestivals.FirstWarning, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 24), EContractFestivals.ChristmasEve, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 25), ELegalFestivals.ChristmasDayFirst, EGermanLand.AllLands)
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 26), ELegalFestivals.ChristmasDaySecond, EGermanLand.AllLands)
'
'    i = i + 1:    Fests(i) = New_LegalFestival(DateSerial(Year, 12, 31), EContractFestivals.NewYearsEve, EGermanLand.AllLands)
'
    'Set GetFestivals = Festivals
End Function
'
'Public Property Get Festivals_Index(this() As LegalFestival, ByVal aDate As Date) As Integer
'    'returns the index in the array if aDate is a legal, religious or festival holiday otherwise 0
'    Dim i As Integer
'    For i = LBound(this) To UBound(this)
'        If this(i).Date = aDate Then
'            Festivals_Index = i
'            Exit Property
'        End If
'    Next
'End Property
' ^ ############################## ^ '       the legal and religious holidays / festivals       ' ^ ############################## ^ '

'fuck how to do this properly?
'Public Function New_CalendarYear(ByVal Year As Integer, _
'                                 Optional ByVal StartMonth As Integer = 1, _
'                                 Optional ByVal EndMonth As Integer = 12, _
'                                 Optional ByVal includeLastDec As Boolean = False, _
'                                 Optional ByVal includeNextJan As Boolean = False) As CalendarYear
'    Dim Y As CalendarYear
'    Y.Year = Year
'    Y.Fests = GetFestivals(Year)
'    StartMonth = IIf(0 < StartMonth And StartMonth <= 12, StartMonth, 1)
'    EndMonth = IIf(StartMonth <= EndMonth And EndMonth <= 12, EndMonth, 12)
'    ReDim Y.Months(StartMonth To EndMonth)
'    Dim m As Integer
'    For m = StartMonth To EndMonth
'        Y.Months(m) = New_CalendarMonth(Y, m)
'    Next
'    New_CalendarYear = Y
'End Function
'
'Public Function New_CalendarMonth(CalYear As CalendarYear, ByVal Month As Integer) As CalendarMonth
'    With New_CalendarMonth
'        .Year = CalYear.Year
'        .Month = Month
'        Dim mds As Integer: mds = DaysInMonth(.Year, Month)
'        ReDim .Days(1 To mds)
'        Dim d As Integer
'        For d = 1 To mds
'            .Days(d) = New_CalendarDay(CalYear, Month, d)
'        Next
'    End With
'End Function
'
'Public Function New_CalendarDay(CalYear As CalendarYear, ByVal Month As Integer, ByVal Day As Integer) As CalendarDay
'    With New_CalendarDay
'        .Day = Day
'        .Date = DateSerial(CalYear.Year, Month, Day)
'        .FestivalIndex = Festivals_Index(CalYear.Fests, .Date)
'    End With
'End Function
'
'Public Function New_Calendar(LastY As CalendarYear, ThisY As CalendarYear, NextY As CalendarYear) As Calendar
'    With New_Calendar
'        .LastYear = LastY
'        .ThisYear = ThisY
'        .NextYear = NextY
'    End With
'End Function

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
'
'Public Function New_CalendarView(Canvas As PictureBox) As CalendarView
'    With New_CalendarView
'        Set .Canvas = Canvas
'        .ColorNormalGrid = RGB(240, 240, 240)
'        .ColorFestivlDay = RGB(222, 141, 245)
'        .ColorMouseOver = RGB(255, 255, 0)
'        .ColorSelected = RGB(255, 0, 0)
'        .ColorBirthDay = RGB(0, 255, 0)
'        .ColorWeekday = RGB(255, 255, 255)
'        .ColorSaturday = RGB(230, 244, 253)
'        .ColorSunday = RGB(137, 189, 226)
'        .ColorLNWeekday = RGB(255, 255, 255)
'        .ColorLNSaturday = RGB(200, 202, 201)
'        .ColorLNSunday = RGB(157, 157, 157)
'        Set .FontDayNrName = New_StdFont("Segoe UI") '("Comic Sans MS")
'        Set .FontMonthName = New_StdFont("Segoe Print", 10, True) '("Comic Sans MS")
'        'Set .FontMonthName = New_StdFont("Comic Sans MS")
'        Set .FontWeekNr = New_StdFont("Segoe UI") '("Comic Sans MS")
'        .HasDecLastYear = True
'        .HasJanNextYear = True
'        .HasMonthNames = True
'        .HasWeekDayNames = True
'        .HasWeekNumbers = True
'        .MarginCal = New_Margin(10)
'        '.MarginCalLeft = 10 'px
'        '.MarginCalTop = 10 'px
'        '.MarginCalRight = 10 'px
'        '.MarginCalBottom = 10 'px
'    End With
'End Function
'
'Public Function CalendarView_Clone(other As CalendarView) As CalendarView
'    With CalendarView_Clone
'        Set .Canvas = other.Canvas
'        .HasDecLastYear = other.HasDecLastYear
'        .HasJanNextYear = other.HasJanNextYear
'        .HasMonthNames = other.HasMonthNames
'        .HasWeekDayNames = other.HasWeekDayNames
'        .HasWeekNumbers = other.HasWeekNumbers
'        .MarginCal = other.MarginCal
'        .MarginMon = other.MarginMon
'        .MarginDay = other.MarginDay
'        .ColorNormalGrid = other.ColorNormalGrid
'        .ColorFestivlDay = other.ColorFestivlDay
'        .ColorMouseOver = other.ColorMouseOver
'        .ColorWeekday = other.ColorWeekday
'        .ColorSaturday = other.ColorSaturday
'        .ColorSunday = other.ColorSunday
'        .ColorLNWeekday = other.ColorLNWeekday
'        .ColorLNSaturday = other.ColorLNSaturday
'        .ColorLNSunday = other.ColorLNSunday
'        .ColTmpWeekday = other.ColTmpWeekday
'        .ColTmpSaturday = other.ColTmpSaturday
'        .ColTmpSunday = other.ColTmpSunday
'        Set .FontMonthName = StdFont_Clone(other.FontMonthName)
'        Set .FontDayNrName = StdFont_Clone(other.FontDayNrName)
'        Set .FontWeekNr = StdFont_Clone(other.FontWeekNr)
'        .TmpDayWidth = other.TmpDayWidth
'        .TmpDayHeight = other.TmpDayHeight
'    End With
'End Function
'
'Public Function CalendarView_Dispose(this As CalendarView)
'    With this
'        Set .Canvas = Nothing
'        Set .FontDayNrName = Nothing
'        Set .FontMonthName = Nothing
'        Set .FontWeekNr = Nothing
'    End With
'End Function
'
'Public Sub CalendarDay_MouseOut(this As CalendarDay, CalYear As CalendarYear)
'    With this
'        Dim m As Integer: m = Month(.Date)
'        If m = 0 Then Exit Sub
'        If .Day = 0 Then Exit Sub
'        CalYear.Months(Month(.Date)).Days(.Day).MouseOver = False
'    End With
'End Sub
'
'Public Function CalendarView_CalendarDayFromMouseCoords(this As CalendarView, CalYear As CalendarYear, ByVal MouseX As Single, ByVal MouseY As Single) As CalendarDay
'    With this
'        MouseX = MouseX - .MarginCal.Left
'        MouseY = MouseY - .MarginCal.Top
'        Dim m As Integer: m = CInt(MouseX \ .TmpDayWidth)         ' x-axis
'        Dim d As Integer: d = CInt(MouseY \ .TmpDayHeight) '- 1  ' y-axis
'        Dim dly As Integer: dly = IIf(.HasDecLastYear, 1, 0)
'        Dim jny As Integer: jny = IIf(.HasJanNextYear, 1, 0)
'        If 0 <= m And m <= UBound(CalYear.Months) + 1 + dly + jny Then
'            If 0 < d And d < MTime.DaysInMonth(CalYear.Year, m) Then
'                CalYear.Months(m).Days(d).MouseOver = True
'                CalendarView_CalendarDayFromMouseCoords = CalYear.Months(m).Days(d)
'            End If
'        End If
'    End With
'End Function
'
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
'
'Public Sub CalendarView_DrawYear(this As CalendarView, CalYear As CalendarYear)
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
'Public Sub CalendarView_DrawMonth(this As CalendarView, CalMonth As CalendarMonth)
''Try: On Error GoTo Catch
'    With this
'        Dim X As Double: X = .Canvas.CurrentX
'        Dim y As Double: y = .MarginCal.Top
'        Dim ny As Integer
'        If .HasMonthNames Then
'            Set .Canvas.Font = .FontMonthName
'            Dim s As String: s = MonthName(CalMonth.Month) & " '" & Right(CStr(CalMonth.Year), 2)
'            .Canvas.Print s
'            ny = ny + 1
'            .Canvas.CurrentY = .MarginCal.Top + ny * .TmpDayHeight
'        End If
'        .Canvas.CurrentX = X
'        Dim d As Integer
'        Dim L As Integer: L = LBound(CalMonth.Days)
'        Dim u As Integer: u = UBound(CalMonth.Days)
'        For d = L To u
'            CalendarView_DrawDay this, CalMonth.Days(d)
'            ny = ny + 1
'            .Canvas.CurrentY = .MarginCal.Top + ny * .TmpDayHeight
'        Next
'        .Canvas.CurrentY = X
'        .Canvas.CurrentY = y
'    End With
''Catch:
'End Sub
'
'Public Sub CalendarView_DrawDay(this As CalendarView, CalDay As CalendarDay)
''Try: On Error GoTo Catch
'    With this
'        Dim fc As Long: fc = .Canvas.ForeColor
'        Dim X As Double: X = .Canvas.CurrentX
'        Dim y As Double: y = .Canvas.CurrentY
'        Dim wd As VbDayOfWeek: wd = Weekday(CalDay.Date)
'        Dim c As Long: c = IIf(wd = vbSaturday, .ColTmpSaturday, IIf(wd = VbDayOfWeek.vbSunday, .ColTmpSunday, .ColTmpWeekday))
'
'        .Canvas.Line (X, y)-(X + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, BF
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
'        .Canvas.Line (X, y)-(X + .TmpDayWidth - 1, y + .TmpDayHeight - 1), c, B
'
'        .Canvas.CurrentX = X
'        .Canvas.CurrentY = y
'
'        Dim s As String
'        s = CStr(CalDay.Day) & " " & VbWeekDay_ToStr(wd, vbSunday, True)
'        If CalDay.FestivalIndex Then
'            s = s & " " & MDECalendar.ELegalFestivals_ToStr(CalDay.FestivalIndex)
'        Else
'            If wd = vbMonday Then
'                s = s & " " & "KW " & WeekOfYearISO(CalDay.Date)
'            End If
'        End If
'
'        Set .Canvas.Font = .FontDayNrName
'        If wd = vbSunday Then
'            .Canvas.ForeColor = RGB(255, 255, 255)
'        End If
'
'        .Canvas.Print s
'
'        .Canvas.CurrentX = X
'        .Canvas.CurrentY = y
'        .Canvas.ForeColor = fc
'    End With
''Catch:
'End Sub
