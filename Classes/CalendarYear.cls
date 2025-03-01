VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Year        As Integer
Private m_StartMonth  As Integer
Private m_EndMonth    As Integer
Private m_Months()    As CalendarMonth
Private m_Festivals   As Collection 'Of FestivalDay 'ehem LegalFestival

Friend Sub New_(ByVal Year As Integer, Optional ByVal StartMonth As Integer = 1, _
                                       Optional ByVal EndMonth As Integer = 12)
    m_Year = Year
    m_StartMonth = IIf(0 < StartMonth And StartMonth <= 12, StartMonth, 1)
    m_EndMonth = IIf(m_StartMonth <= EndMonth And EndMonth <= 12, EndMonth, 12)
    ReDim m_Months(m_StartMonth To m_EndMonth)
    Dim m As Integer
    For m = m_StartMonth To m_EndMonth
        Set m_Months(m) = MNew.CalendarMonth(Me, m)
    Next
    Set m_Festivals = MDECalendar.GetFestivals(m_Year)
    Dim fd As FestivalDay
    For Each fd In m_Festivals
        m = fd.Month
        If m_StartMonth <= m And m <= m_EndMonth Then
            Dim d As Integer: d = fd.DayInMonth
            Dim cd As CalendarDay: Set cd = m_Months(m).Day(d)
            If Not cd Is Nothing Then Set cd.FestivalDay = fd
        End If
    Next
End Sub

Public Property Get Year() As Integer
    Year = m_Year
End Property

Public Property Get StartMonth() As Integer
    StartMonth = m_StartMonth
End Property

Public Property Get EndMonth() As Integer
    EndMonth = m_EndMonth
End Property

Public Property Get CountMonth() As Integer
    CountMonth = m_EndMonth - m_StartMonth + 1
End Property

Public Property Get Month(ByVal Index As Integer) As CalendarMonth 'Index=1 = Jan, Index=2=Feb, Index=12=Dec
    If Index < m_StartMonth Or m_EndMonth < Index Then Exit Property
    Set Month = m_Months(Index)
End Property
