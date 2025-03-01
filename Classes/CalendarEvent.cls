VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_DateFrom As Date
Private m_DateTo   As Date
Private m_Name     As String

Friend Sub New_(ByVal EventFromDate As Date, ByVal EventToDate As Date, ByVal EventName As String)
    m_DateFrom = EventFromDate: m_DateTo = EventToDate: m_Name = EventName
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    Name = m_Name
End Property

Public Property Get DateFrom() As Date
    DateFrom = m_DateFrom
End Property
Public Property Let DateFrom(ByVal Value As Date)
    m_DateFrom = Value
End Property

Public Property Get DateTo() As Date
    DateTo = m_DateTo
End Property
Public Property Let DateTo(ByVal Value As Date)
    m_DateTo = Value
End Property

