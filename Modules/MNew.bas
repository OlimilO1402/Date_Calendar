Attribute VB_Name = "MNew"
Option Explicit

Public Function Calendar(ByVal ThisYear As CalendarYear, Optional ByVal LastYear As CalendarYear = Nothing, _
                                                         Optional ByVal NextYear As CalendarYear = Nothing) As Calendar
    Set Calendar = New Calendar: Calendar.New_ ThisYear, LastYear, NextYear
End Function

Public Function GetDefaultYear() As Integer
    Dim dNow  As Date:     dNow = DateTime.Now
    Dim Year  As Integer:  Year = DateTime.Year(dNow)
    Dim Month As Integer: Month = DateTime.Month(dNow)
    Dim Day   As Integer:   Day = DateTime.Day(dNow)
    GetDefaultYear = Year + IIf((Month = 12) And (Day = 15), 1, 0)
End Function

Public Function CalendarDefault() As Calendar
    Dim thisY As Integer: thisY = GetDefaultYear
    Set CalendarDefault = MNew.Calendar(MNew.CalendarYear(thisY, 1, 12), _
                                        MNew.CalendarYear(thisY - 1, 12, 12), _
                                        MNew.CalendarYear(thisY + 1, 1, 1))
End Function

Public Function CalendarYear(ByVal Year As Integer, Optional ByVal StartMonth As Integer = 1, _
                                                    Optional ByVal EndMonth As Integer = 12) As CalendarYear
    Set CalendarYear = New CalendarYear: CalendarYear.New_ Year, StartMonth, EndMonth
End Function

Public Function CalendarMonth(CalYear As CalendarYear, ByVal Month As Integer) As CalendarMonth
    Set CalendarMonth = New CalendarMonth: CalendarMonth.New_ CalYear, Month
End Function

Public Function CalendarDay(ByVal Month As CalendarMonth, ByVal DayInMonth As Integer) As CalendarDay
    Set CalendarDay = New CalendarDay: CalendarDay.New_ Month, DayInMonth
End Function

Public Function CalendarView(aCalendar As Calendar, Canvas As PictureBox) As CalendarView
    Set CalendarView = New CalendarView: CalendarView.New_ aCalendar, Canvas
End Function

Public Function FestivalDay(ByVal aDate As Date, ByVal EFestival As Long) As FestivalDay
    Set FestivalDay = New FestivalDay: FestivalDay.New_ aDate, EFestival
End Function

Public Function PersonalEvent(aDate As Date, aName As String, EventName As String) As PersonalEvent
    Set PersonalEvent = New PersonalEvent: PersonalEvent.New_ aDate, aName, EventName
End Function

Public Function CalendarEvent(ByVal EventFromDate As Date, ByVal EventToDate As Date, ByVal EventName As String) As CalendarEvent
    Set CalendarEvent = New CalendarEvent: CalendarEvent.New_ EventFromDate, EventToDate, EventName
End Function
