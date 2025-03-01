VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Year   As CalendarYear
Private m_Month  As Integer
Private m_nDays  As Integer
Private m_Days() As CalendarDay

Private m_Caret  'As Rectangle

Friend Sub New_(CalYear As CalendarYear, ByVal Month As Integer)
    Set m_Year = CalYear
    m_Month = Month
    Dim y As Integer: y = m_Year.Year
    m_nDays = MTime.DaysInMonth(y, m_Month)
    ReDim m_Days(1 To m_nDays)
    Dim d As Integer
    For d = 1 To m_nDays
        Set m_Days(d) = MNew.CalendarDay(Me, d)
    Next
End Sub

Public Property Get Year() As CalendarYear
    Set Year = m_Year
End Property

Public Property Get Month() As Integer
    Month = m_Month
End Property

Public Property Get CountDays() As Integer
    CountDays = m_nDays
End Property
    
Public Property Get Days() As CalendarDay()
    Days = m_Days()
End Property

Public Property Get Day(ByVal Index As Integer) As CalendarDay
    If Index < 1 Or m_nDays < Index Then Exit Property
    Set Day = m_Days(Index)
End Property
