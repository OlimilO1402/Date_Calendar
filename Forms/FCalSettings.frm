VERSION 5.00
Begin VB.Form FCalSettings 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Calendar Settings"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBFestivalDay 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   20
      Top             =   2760
      Width           =   375
   End
   Begin VB.PictureBox PBNormalGrid 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   19
      Top             =   2760
      Width           =   375
   End
   Begin VB.PictureBox PBColorSundayLN 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox PBColorSaturdayLN 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox PBColorWeekdayLN 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox PBColorSunday 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox PBColorSaturday 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox PBColorWeekday 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton BtnSetFontDayNrName 
      Caption         =   "..."
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton BtnSetFontMonthName 
      Caption         =   "..."
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Festival-Day Border:"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   2760
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Normal Grid:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fonts:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Last/Next Year:"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "This Year:"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sundays:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Saturdays:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Weekdays:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label LblColors 
      AutoSize        =   -1  'True
      Caption         =   "Colors:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label LblFontDayNrName 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "12 So Muttertag"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblFontMonthname 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "September '24 "
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FCalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CalView As MDECalendar.CalendarView
Private m_Result  As VbMsgBoxResult
Private m_FontDlg As FontDialog
Private m_ColrDlg As ColorDialog

Private Sub Form_Load()
    Set m_FontDlg = New FontDialog
    Set m_ColrDlg = New ColorDialog
End Sub

Friend Function ShowDialog(FOwner As Form, CalV_inout As CalendarView) As VbMsgBoxResult
    m_CalView = CalV_inout ' CalendarView_Clone(CalV_inout)
    UpdateView
    Me.Show vbModal, FOwner
    If m_Result = vbCancel Then Exit Function
    CalV_inout = m_CalView 'CalendarView_Clone(m_CalView)
End Function

Private Sub UpdateView()
    Set LblFontMonthname.Font = m_CalView.FontMonthName
    Set LblFontDayNrName.Font = m_CalView.FontDayNrName
    PBColorWeekday.BackColor = m_CalView.ColorWeekday
    PBColorSaturday.BackColor = m_CalView.ColorSaturday
    PBColorSunday.BackColor = m_CalView.ColorSunday
    PBColorWeekdayLN.BackColor = m_CalView.ColorLNWeekday
    PBColorSaturdayLN.BackColor = m_CalView.ColorLNSaturday
    PBColorSundayLN.BackColor = m_CalView.ColorLNSunday
    PBNormalGrid.BackColor = m_CalView.ColorNormalGrid
    PBFestivalDay.BackColor = m_CalView.ColorFestivlDay
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub BtnSetFontMonthName_Click()
    Set m_FontDlg.Font = m_CalView.FontMonthName
    If m_FontDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    Set m_CalView.FontMonthName = m_FontDlg.Font
    UpdateView
End Sub

Private Sub BtnSetFontDayNrName_Click()
    Set m_FontDlg.Font = m_CalView.FontDayNrName
    If m_FontDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    Set m_CalView.FontDayNrName = m_FontDlg.Font
    UpdateView
End Sub

Private Sub PBColorWeekday_Click()
    m_ColrDlg.Color = PBColorWeekday.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorWeekday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBColorSaturday_Click()
    m_ColrDlg.Color = PBColorSaturday.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorSaturday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBColorSunday_Click()
    m_ColrDlg.Color = PBColorSunday.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorSunday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBColorWeekdayLN_Click()
    m_ColrDlg.Color = PBColorWeekdayLN.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorLNWeekday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBColorSaturdayLN_Click()
    m_ColrDlg.Color = PBColorSaturdayLN.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorLNSaturday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBColorSundayLN_Click()
    m_ColrDlg.Color = PBColorSundayLN.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorLNSunday = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBNormalGrid_Click()
    m_ColrDlg.Color = PBNormalGrid.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorNormalGrid = m_ColrDlg.Color
    UpdateView
End Sub

Private Sub PBFestivalDay_Click()
    m_ColrDlg.Color = PBFestivalDay.BackColor
    If m_ColrDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_CalView.ColorFestivlDay = m_ColrDlg.Color
    UpdateView
End Sub

