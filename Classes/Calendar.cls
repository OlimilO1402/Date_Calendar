VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_LastYear As CalendarYear
Private m_ThisYear As CalendarYear
Private m_NextYear As CalendarYear

Friend Sub New_(ByVal ThisYear As CalendarYear, Optional ByVal LastYear As CalendarYear = Nothing, Optional ByVal NextYear As CalendarYear = Nothing)
    Set m_ThisYear = ThisYear
    Set m_LastYear = LastYear
    Set m_NextYear = NextYear
End Sub

Public Property Get LastYear() As CalendarYear
    Set LastYear = m_LastYear
End Property

Public Property Get ThisYear() As CalendarYear
    Set ThisYear = m_ThisYear
End Property

Public Property Get NextYear() As CalendarYear
    Set NextYear = m_NextYear
End Property

Public Property Get CountMonth() As Integer
    Dim c As Long
    If Not m_LastYear Is Nothing Then c = c + m_LastYear.CountMonth
    If Not m_ThisYear Is Nothing Then c = c + m_ThisYear.CountMonth
    If Not m_NextYear Is Nothing Then c = c + m_NextYear.CountMonth
    CountMonth = c
End Property

Public Function GetMonthFromGrid(ByVal m As Integer) As CalendarMonth
    If m <= 0 Then Exit Function
    Dim cm As Integer: cm = m_LastYear.CountMonth
    If m <= cm Then Set GetMonthFromGrid = m_LastYear.Month(m): Exit Function
    m = m - cm
    cm = m_ThisYear.CountMonth
    If m <= cm Then Set GetMonthFromGrid = m_ThisYear.Month(m): Exit Function
    m = m - cm
    cm = m_NextYear.CountMonth
    If m <= cm Then Set GetMonthFromGrid = m_NextYear.Month(m): Exit Function
End Function

Public Function GetDayFromGrid(ByVal m As Integer, d As Integer) As CalendarDay
    Dim CalMonth As CalendarMonth: Set CalMonth = GetMonthFromGrid(m)
    Dim cd As Integer: cd = CalMonth.CountDays
    If d <= 0 Or cd < d Then Exit Function
    Set GetDayFromGrid = CalMonth.Day(d)
End Function

