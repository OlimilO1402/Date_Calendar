VERSION 5.00
Begin VB.Form FCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19935
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FCalendar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   19935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSetColorAndFonts 
      Caption         =   "Set Colors And Fonts..."
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      ToolTipText     =   "Uses ""Microsoft Print to PDF"" by default"
      Top             =   60
      Width           =   2415
   End
   Begin VB.CommandButton BtnPrintToPDF 
      Caption         =   "Print to pdf..."
      Height          =   375
      Left            =   12840
      TabIndex        =   9
      ToolTipText     =   "Uses ""Microsoft Print to PDF"" by default"
      Top             =   60
      Width           =   1695
   End
   Begin VB.ComboBox CmbMonthTo 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   60
      Width           =   1455
   End
   Begin VB.ComboBox CmbMonthFrom 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   60
      Width           =   1455
   End
   Begin VB.CheckBox ChkNextJan 
      Caption         =   "Next January"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.CheckBox ChkLastDec 
      Caption         =   "Last December"
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.ComboBox CmbYear 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   60
      Width           =   1215
   End
   Begin VB.PictureBox PBCalendar 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   599
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1079
      TabIndex        =   10
      Top             =   480
      Width           =   16215
   End
   Begin VB.Label LblMouseOverDate 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   255
      Left            =   14640
      TabIndex        =   11
      Top             =   120
      Width           =   570
   End
   Begin VB.Label LblMonthTo 
      AutoSize        =   -1  'True
      Caption         =   "to:"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   225
   End
   Begin VB.Label LblMonthFrom 
      AutoSize        =   -1  'True
      Caption         =   "from:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
   Begin VB.Label LblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "FCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Calendar  As MDECalendar.CalendarYear
Private m_CalView   As MDECalendar.CalendarView
Private m_PaperSize As PrinterObjectConstants
Private m_PapOrient As PrinterObjectConstants
Private m_CalDayMouseOver As MDECalendar.CalendarDay

Private Sub Form_Load()
    Me.Caption = "My Calendar v" & App.Major & "." & App.Minor & "." & App.Revision
    m_CalView = MDECalendar.New_CalendarView(Me.PBCalendar)
    FillCombos
    CmbYear.ListIndex = Year(Now) - 1900 + 1
    CmbMonthFrom.ListIndex = 0
    CmbMonthTo.ListIndex = 11
    m_PaperSize = PrinterObjectConstants.vbPRPSA4
    m_PapOrient = PrinterObjectConstants.vbPRORLandscape
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = PBCalendar.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then PBCalendar.Move L, T, W, H
End Sub

Private Sub FillCombos()
    FillYears CmbYear
    FillMonths CmbMonthFrom
    FillMonths CmbMonthTo
End Sub

Private Sub FillYears(cb As ComboBox)
    cb.Clear: Dim Y As Integer: For Y = 1900 To 2100: cb.AddItem Y: Next
End Sub
Private Sub FillMonths(cb As ComboBox)
    cb.Clear: Dim m As Integer: For m = 1 To 12: cb.AddItem MonthName(m): Next
End Sub

Private Sub CmbYear_Click()
    Debug.Print "Click"
    UpdateData
End Sub

'Private Sub CmbYear_Change()
'    Debug.Print "Change"
'    UpdateData
'End Sub

Private Sub CmbYear_Scroll()
    Debug.Print "Scroll"
    UpdateData
End Sub

Private Sub CmbMonthFrom_Click()
    CmbMonthTo.ListIndex = Max(CmbMonthFrom.ListIndex, CmbMonthTo.ListIndex)
    UpdateData
End Sub

Private Sub CmbMonthTo_Click()
    CmbMonthFrom.ListIndex = Min(CmbMonthFrom.ListIndex, CmbMonthTo.ListIndex)
    UpdateData
End Sub

Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Private Sub ChkLastDec_Click()
    UpdateData
End Sub

Private Sub ChkNextJan_Click()
    UpdateData
End Sub

Private Sub UpdateData()
    Dim Y As Integer: Y = CmbYear.ListIndex + 1900
    Dim mf As Integer: mf = CmbMonthFrom.ListIndex + 1
    Dim mt As Integer: mt = CmbMonthTo.ListIndex + 1
    m_Calendar = MDECalendar.New_CalendarYear(Y, mf, mt)
    'm_CalView = MDECalendar.New_CalendarView(m_Calendar, Me.PBCalendar)
    'Set m_CalView.Canvas = Me.PBCalendar
    m_CalView = MDECalendar.CalendarView_Clone(m_CalView)
    
    m_CalView.HasDecLastYear = ChkLastDec.Value = CheckBoxConstants.vbChecked
    m_CalView.HasJanNextYear = ChkNextJan.Value = CheckBoxConstants.vbChecked
    UpdateView
End Sub

Private Sub UpdateView()
    PBCalendar.Cls
    MDECalendar.CalendarView_DrawYear m_CalView, m_Calendar
End Sub

Private Sub PBCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CalDay As CalendarDay: CalDay = CalendarView_CalendarDayFromMouseCoords(m_CalView, m_Calendar, X, Y)
    LblMouseOverDate.Caption = CalDay.Date
    If CalDay.Date <> m_CalDayMouseOver.Date Then
        CalendarDay_MouseOut m_CalDayMouseOver, m_Calendar
        m_CalDayMouseOver = CalDay
        UpdateView
    End If
End Sub

Private Sub PBCalendar_Paint()
    UpdateView
End Sub

Private Sub PBCalendar_Resize()
    UpdateView
End Sub

Private Sub BtnSetColorAndFonts_Click()
    If FCalSettings.ShowDialog(Me, m_CalView) <> VbMsgBoxResult.vbOK Then Exit Sub
    UpdateView
End Sub

Private Function SelectPrinter(ByVal PrinterName As String) As Printer
    Dim i As Long
    For i = 0 To Printers.Count - 1
        If UCase(Printers(i).DeviceName) = UCase(PrinterName) Then 'e.g.: "Microsoft Print to PDF"
            Set SelectPrinter = Printers(i)
            Exit Function
        End If
    Next
End Function

Private Sub BtnPrintToPDF_Click()
Try: On Error GoTo Catch
    
    Dim CalView As CalendarView: CalView = MDECalendar.CalendarView_Clone(m_CalView)
        
    If FPaperSettings.ShowDialog(Me, m_PaperSize, m_PapOrient) <> VbMsgBoxResult.vbOK Then Exit Sub
    
    
    Set Printer = SelectPrinter("Microsoft Print to PDF")
    Set CalView.Canvas = Printer
    Printer.PaperSize = m_PaperSize
    Printer.Orientation = m_PapOrient
    'Debug.Print Printer.DriverName '= "winspool"
    'Debug.Print Printer.DeviceName '= "Microsoft Print to PDF"
    'Printer.NewPage
    MDECalendar.CalendarView_DrawYear CalView, m_Calendar
    Printer.EndDoc
    Printer.KillDoc
    
    'Exit Sub
    GoTo Finally
Catch:
    If Err.Number = 482 Then
        'User selected Cancel
        'MsgBox "There was an error, User selected Cancel or "
        Debug.Print Err.LastDllError
    Else
        MsgBox "Error during printing!" & vbCrLf & Err.Number & " " & Err.Description
    End If
    'Debug.Print Err.Number
Finally:
    MDECalendar.CalendarView_Dispose CalView
End Sub

