VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonalEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Date       As Date
Private m_PersonName As String
Private m_EventName  As String

Friend Sub New_(aDate As Date, PersonName As String, EventName As String)
    m_Date = aDate: m_PersonName = PersonName: m_EventName = EventName
End Sub

Public Property Get EventDate() As Date
    EventDate = m_Date
End Property
Public Property Let EventDate(ByVal Value As Date)
    m_Date = Value
End Property

Public Property Get PersonName() As String
    PersonName = m_PersonName
End Property
Public Property Let PersonName(ByVal Value As String)
    m_PersonName = Value
End Property

Public Property Get EventName() As String
    EventName = m_EventName
End Property
Public Property Let EventName(ByVal Value As String)
    m_EventName = Value
End Property

