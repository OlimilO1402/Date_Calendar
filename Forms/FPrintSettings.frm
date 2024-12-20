VERSION 5.00
Begin VB.Form FPaperSettings 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Paper-Settings"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ListBox LstPaperSize 
      Height          =   735
      ItemData        =   "FPrintSettings.frx":0000
      Left            =   120
      List            =   "FPrintSettings.frx":000A
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton OptPortrait 
      Caption         =   "Portrait"
      Height          =   975
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   120
      Width           =   690
   End
   Begin VB.OptionButton OptLandscape 
      Caption         =   "Landscape"
      Height          =   690
      Left            =   2040
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
End
Attribute VB_Name = "FPaperSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PaperSize As PrinterObjectConstants
Private m_PapOrient As PrinterObjectConstants
Private m_Result    As VbMsgBoxResult

Public Function ShowDialog(FOwner As Form, PaperSize_inout As PrinterObjectConstants, PapOrient_inout As PrinterObjectConstants) As VbMsgBoxResult
    m_PaperSize = PaperSize_inout
    m_PapOrient = PapOrient_inout
    UpdateView
    Me.Show vbModal, FOwner
    PaperSize_inout = m_PaperSize
    PapOrient_inout = m_PapOrient
    ShowDialog = m_Result
End Function

Private Sub UpdateView()
    If m_PapOrient = PrinterObjectConstants.vbPRORLandscape Then
        OptLandscape.Value = True
        OptPortrait.Value = False
    Else
        OptLandscape.Value = False
        OptPortrait.Value = True
    End If
    If m_PaperSize = PrinterObjectConstants.vbPRPSA4 Then
        LstPaperSize.Selected(0) = True
        LstPaperSize.Selected(1) = False
    Else
        LstPaperSize.Selected(0) = False
        LstPaperSize.Selected(1) = True
    End If
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub LstPaperSize_Click()
    If LstPaperSize.ListIndex = 0 Then
        LstPaperSize.Selected(0) = True
        LstPaperSize.Selected(1) = False
        m_PaperSize = PrinterObjectConstants.vbPRPSA4
    ElseIf LstPaperSize.ListIndex = 1 Then
        LstPaperSize.Selected(0) = False
        LstPaperSize.Selected(1) = True
        m_PaperSize = PrinterObjectConstants.vbPRPSA3
    End If
End Sub

Private Sub LstPaperSize_ItemCheck(Item As Integer)
    If Item = 0 Then
        LstPaperSize.Selected(1) = False
        m_PaperSize = PrinterObjectConstants.vbPRPSA4
    ElseIf Item = 1 Then
        LstPaperSize.Selected(0) = False
        m_PaperSize = PrinterObjectConstants.vbPRPSA3
    End If
End Sub

Private Sub OptLandscape_Click()
    m_PapOrient = PrinterObjectConstants.vbPRORLandscape
End Sub

Private Sub OptPortrait_Click()
    m_PapOrient = PrinterObjectConstants.vbPRORPortrait
End Sub
