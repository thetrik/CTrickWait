VERSION 5.00
Begin VB.Form frmProcessNotify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process notify"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProcess 
      Caption         =   "Wait process"
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3045
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   390
         Left            =   780
         TabIndex        =   2
         Top             =   1020
         Width           =   1440
      End
      Begin VB.TextBox txtProcess 
         Height          =   360
         Left            =   105
         TabIndex        =   1
         Text            =   "notepad"
         Top             =   555
         Width           =   2820
      End
      Begin VB.Label lblInfo 
         Caption         =   "Execute file:"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmProcessNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // CTrickWait - demonstration of waiting for the end of the process.
' // © The trick, 2015-2021
' //

Option Explicit

Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long

Private Const SYNCHRONIZE   As Long = &H100000
Private Const INFINITE      As Long = -1

Private WithEvents m_cProcess As CTrickWait  ' // Asynchronous waiter
Attribute m_cProcess.VB_VarHelpID = -1

Private Sub cmdRun_Click()
    Dim lPID         As Long
    Dim hProcess    As Long
    
    ' // If already waiting
    If m_cProcess.IsActive Then
        Select Case MsgBox("Process enabled. Abort?", vbYesNo Or vbQuestion)
        Case vbYes: m_cProcess.Abort
        Case Else: Exit Sub
        End Select
    End If
    
    ' // Launch the process
    lPID = Shell(txtProcess)
    
    ' // Open the process for the synchronization
    hProcess = OpenProcess(SYNCHRONIZE, False, lPID)

    ' // Launch the asynchronous waiting
    m_cProcess.vbWaitForSingleObject hProcess, INFINITE
    
End Sub

Private Sub Form_Load()
    Set m_cProcess = New CTrickWait
End Sub

' // Event occurs when the process will end
Private Sub m_cProcess_OnWait( _
            ByVal Handle As Long, _
            ByVal Result As Long)

    MsgBox "Process has ended." & vbNewLine & "Handle = " & Handle & vbNewLine & "Result = " & Result
    CloseHandle Handle
    
End Sub
