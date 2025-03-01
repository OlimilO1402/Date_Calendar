VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Month         As CalendarMonth
Private m_DayInMonth    As Integer
Private m_Date          As Date

Private m_FestivalDay   As FestivalDay 'Integer '0 = no festivalday
Private m_Events        As Collection  'Of PersonalEvent 'Birthday / marriage etc
Private m_MouseOver     As Boolean
Private m_Selected      As Boolean

Private m_Caret         As Rectangle

Friend Sub New_(ByVal Month As CalendarMonth, ByVal DayInMonth As Integer)
    Set m_Month = Month
    m_DayInMonth = DayInMonth
    m_Date = DateSerial(Me.Year.Year, Me.Month.Month, m_DayInMonth)
End Sub

Public Property Get Year() As CalendarYear
    Set Year = m_Month.Year
End Property

Public Property Get Month() As CalendarMonth
    Set Month = m_Month
End Property

Public Property Get DayInMonth() As Integer
    DayInMonth = m_DayInMonth
End Property

Public Property Get DayDate() As Date
    DayDate = m_Date
End Property

Public Property Get IsFestivalDay() As Boolean
    IsFestivalDay = Not m_FestivalDay Is Nothing
End Property

Public Property Get FestivalDay() As FestivalDay
    Set FestivalDay = m_FestivalDay
End Property
Public Property Set FestivalDay(Value As FestivalDay)
    Set m_FestivalDay = Value
End Property

Public Property Get PersonalEvents() As Collection 'Of PersonalEvent      'Birthday / marriage etc
    Set PersonalEvents = m_Events
End Property

Friend Property Get IsMouseOver() As Boolean
    IsMouseOver = m_MouseOver
End Property

Friend Property Get MouseOver() As Boolean
    MouseOver = m_MouseOver
End Property
Friend Property Let MouseOver(ByVal Value As Boolean)
    m_MouseOver = Value
End Property

Friend Property Get Selected() As Boolean
    Selected = m_Selected
End Property
Friend Property Let Selected(ByVal Value As Boolean)
    m_Selected = Value
End Property

Friend Property Get Caret() As Rectangle
    Caret = m_Caret
End Property
Friend Property Let Caret(Value As Rectangle)
    m_Caret = Value
End Property

Public Sub MouseOut()
    m_MouseOver = False
End Sub
