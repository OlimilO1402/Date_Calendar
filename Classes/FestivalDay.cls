VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FestivalDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Date     As Date
Private m_Festival As Long
'Private m_Lands    As 'Mdecalendar.EGermanLand 'alle lands or only one or more

Friend Sub New_(ByVal aDate As Date, ByVal EFestival As Long)
    m_Date = aDate: m_Festival = EFestival
End Sub

Public Property Get Key() As String
    Key = CStr(Me.FestivalDate)
End Property

Public Property Get FestivalDate() As Date
    FestivalDate = m_Date
End Property

Public Property Get Month() As Integer
    Month = DateTime.Month(m_Date)
End Property

Public Property Get DayInMonth() As Integer
    DayInMonth = DateTime.Day(m_Date)
End Property

Public Property Get Name() As String
    Name = MDECalendar.Festivals_ToStr(m_Festival)
End Property

Public Property Get WeekdayName() As String
    Dim wd As VbDayOfWeek: wd = DateTime.Weekday(m_Date, vbMonday)
    WeekdayName = Weekday_ToStr(wd)
End Property
