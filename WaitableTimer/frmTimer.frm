VERSION 5.00
Begin VB.Form frmTimer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Waitable timer"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTimer 
      Caption         =   "Waitable timer"
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1890
      Begin VB.CommandButton cmdSetTimer 
         Caption         =   "Set"
         Height          =   375
         Left            =   225
         TabIndex        =   2
         Top             =   990
         Width           =   1440
      End
      Begin VB.TextBox txtTimeClock 
         Height          =   360
         Left            =   150
         TabIndex        =   1
         Top             =   555
         Width           =   1620
      End
      Begin VB.Label lblInfo 
         Caption         =   "Date:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   315
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // CTrickWait - demonstration of waiting for the waitable timer tick.
' // © The trick, 2015 - 2021
' //

Option Explicit

Private Const INFINITE  As Long = -1

Private Declare Function CreateWaitableTimer Lib "kernel32" _
                         Alias "CreateWaitableTimerW" ( _
                         ByRef lpTimerAttributes As Any, _
                         ByVal bManualReset As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" ( _
                         ByVal hTimer As Long, _
                         ByVal lpDueTime As Long, _
                         ByVal lPeriod As Long, _
                         ByVal pfnCompletionRoutine As Long, _
                         ByVal lpArgToCompletionRoutine As Long, _
                         ByVal fResume As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32" ( _
                         ByVal vTime As Date, _
                         ByRef lpSystemTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
                         ByRef st As Any, _
                         ByRef ft As Currency) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" ( _
                         ByRef lpLocalFileTime As Currency, _
                         ByRef lpFileTime As Currency) As Long

Private WithEvents m_cTimer As CTrickWait   ' // Asynchronous waiter
Attribute m_cTimer.VB_VarHelpID = -1
Private m_hTimer            As Long         ' // Handle of the waitable timer

Private Sub cmdSetTimer_Click()

    On Error GoTo cancel_error
    
    Dim dDate       As Date
    Dim iSysTime(8) As Integer
    Dim crFileTime  As Currency
    Dim crLocalTime As Currency
    
    ' // To system time
    dDate = CDate(txtTimeClock)
    VariantTimeToSystemTime dDate, iSysTime(0)
    SystemTimeToFileTime iSysTime(0), crLocalTime
    LocalFileTimeToFileTime crLocalTime, crFileTime
    
    ' // Set the waitable timer
    SetWaitableTimer m_hTimer, VarPtr(crFileTime), 0, 0, 0, 0
    
    ' // If already waiting
    If m_cTimer.IsActive Then
        Select Case MsgBox("Timer enabled. Abort?", vbYesNo Or vbQuestion)
        Case vbYes: m_cTimer.Abort
        Case Else: Exit Sub
        End Select
    End If
    
    ' // Launch the waiting
    m_cTimer.vbWaitForSingleObject m_hTimer, INFINITE
    
    Exit Sub
    
cancel_error:
    
    MsgBox "Error", vbExclamation
    
End Sub

' // Event occurs after the tick of the waitable timer
Private Sub m_cTimer_OnWait( _
            ByVal Handle As Long, _
            ByVal Result As Long)
    MsgBox "Timer event." & vbNewLine & "Handle = " & Handle & vbNewLine & "Result = " & Result
End Sub

Private Sub Form_Load()

    Set m_cTimer = New CTrickWait
    
    ' // Create the waitable timer
    m_hTimer = CreateWaitableTimer(ByVal 0&, False, 0)
    
    ' // Default - the current time
    txtTimeClock = Now
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    CloseHandle m_hTimer
End Sub
