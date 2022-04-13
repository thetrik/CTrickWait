VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form frmFileMonitor 
   Caption         =   "File monitor"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMonitor 
      Caption         =   "File monitor"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5160
      Begin ComctlLib.ListView lstMonitor 
         Height          =   4275
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "#"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Activity"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtMonitor 
         Height          =   360
         Left            =   135
         TabIndex        =   2
         Top             =   495
         Width           =   3510
      End
      Begin VB.CommandButton cmdMonitor 
         Caption         =   "Start"
         Height          =   390
         Left            =   3690
         TabIndex        =   1
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblInfo 
         Caption         =   "Directory:"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmFileMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // CTrickWait - demonstration of monitoring of the file operations
' // © The trick, 2015-2021
' //

Option Explicit

Private Const MAX_PATH = 260

Private Type OVERLAPPED
    Internal        As Long
    InternalHigh    As Long
    offset          As Long
    OffsetHigh      As Long
    hEvent          As Long
End Type

Private Type FILE_NOTIFY_INFORMATION
    dwNextEntryOffset           As Long
    dwAction                    As Long
    dwFileNameLength            As Long
    wcFileName(MAX_PATH * 2)    As Byte
End Type

Private Declare Function ReadDirectoryChanges Lib "kernel32.dll" _
                         Alias "ReadDirectoryChangesW" ( _
                         ByVal m_hDirectory As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nBufferLength As Long, _
                         ByVal bWatchSubTree As Long, _
                         ByVal dwNotifyFilter As Long, _
                         ByVal lpBytesReturned As Long, _
                         ByRef lpOverlapped As Any, _
                         ByVal lpCompletionRoutine As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" _
                         Alias "CreateFileW" ( _
                         ByVal lpFileName As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwShareMode As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByVal dwCreationDisposition As Long, _
                         ByVal dwFlagsAndAttributes As Long, _
                         ByVal hTemplateFile As Long) As Long
Private Declare Function CancelIo Lib "kernel32" ( _
                         ByVal hFile As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" _
                         Alias "CreateEventW" ( _
                         ByVal lpEventAttributes As Long, _
                         ByVal bManualReset As Long, _
                         ByVal bInitialState As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function ResetEvent Lib "kernel32" ( _
                         ByVal m_hEvent As Long) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long

Private Const INFINITE                      As Long = -1
Private Const FILE_LIST_DIRECTORY           As Long = &H1
Private Const FILE_SHARE_DELETE             As Long = &H4
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const FILE_FLAG_BACKUP_SEMANTICS    As Long = &H2000000
Private Const FILE_FLAG_OVERLAPPED          As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = &H3
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const FILE_NOTIFY_CHANGE_FILE_NAME  As Long = 1
Private Const FILE_NOTIFY_CHANGE_DIR_NAME   As Long = 2
Private Const FILE_ACTION_ADDED             As Long = &H1
Private Const FILE_ACTION_REMOVED           As Long = &H2
Private Const FILE_ACTION_RENAMED_OLD_NAME  As Long = &H4
Private Const FILE_ACTION_RENAMED_NEW_NAME  As Long = &H5

Private WithEvents m_cMonitor   As CTrickWait ' // Asynchronous waiter
Attribute m_cMonitor.VB_VarHelpID = -1

Private m_hDirectory    As Long         ' // Handle of the monitored directory
Private m_hEvent        As Long         ' // Handle of the asynchronous event
Private m_bBufEvents()  As Byte         ' // Buffer for the notifications
Private m_tOvp          As OVERLAPPED   ' // Structure which allows to do the asynchronous monitoring
Private m_iWatchSubdir  As Long         ' // Whether you want to track subdirectories as well


Private Sub cmdMonitor_Click()
        
    ' // Check if already opened then stop
    If m_hDirectory Then
    
        ' // Abort the waiting
        m_cMonitor.Abort
        ' // Close the directory handle and the event handle
        CloseHandle m_hEvent:     m_hEvent = 0
        CloseHandle m_hDirectory: m_hDirectory = 0
        ' // Change caption on button
        cmdMonitor.Caption = "Start"
        
        Exit Sub
        
    End If
    
    ' // Open directory for the m_cMonitoring
    m_hDirectory = CreateFile(StrPtr(txtMonitor), FILE_LIST_DIRECTORY, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, 0)
    ' // If error occured then exit
    If m_hDirectory = INVALID_HANDLE_VALUE Then MsgBox "Error open directory", vbExclamation: Exit Sub
    
    ' // Creating the event for the notifications
    m_hEvent = CreateEvent(0, True, True, 0)
    
    ' // Handle error
    If m_hEvent = 0 Then
    
        CloseHandle m_hDirectory: m_hDirectory = 0
        MsgBox "Error create notify event", vbExclamation
        Exit Sub
        
    End If
    ' // Fill the OVERLAPPED structure for the asynchronous call
    m_tOvp.hEvent = m_hEvent
    
    ' // Allocate the buffer for the notifications
    ReDim m_bBufEvents(16383)
    
    ' // Start the m_cMonitor in the asynchronous mode
    If ReadDirectoryChanges(m_hDirectory, m_bBufEvents(0), UBound(m_bBufEvents) + 1, m_iWatchSubdir, FILE_NOTIFY_CHANGE_FILE_NAME _
Or FILE_NOTIFY_CHANGE_DIR_NAME Or FILE_ACTION_ADDED Or FILE_ACTION_REMOVED, 0, m_tOvp, 0) = 0 Then
        ' // Handle error
        MsgBox "Error start m_cMonitor", vbExclamation
        CloseHandle m_hEvent:     m_hEvent = 0
        CloseHandle m_hDirectory: m_hDirectory = 0
        Exit Sub
        
    End If
    
    ' // Launch the asynchronous waiting
    m_cMonitor.vbWaitForSingleObject m_hEvent, INFINITE
    
    cmdMonitor.Caption = "Stop"
    lstMonitor.ListItems.Clear
    
End Sub

Private Sub Form_Load()

    Set m_cMonitor = New CTrickWait
    txtMonitor = Environ("WINDIR")
    m_iWatchSubdir = 1
    
End Sub

' // Event occurs if the directory being changed by the file operation that have monitored
Private Sub m_cMonitor_OnWait( _
            ByVal hHandle As Long, _
            ByVal lResult As Long)
    Dim tNotify As FILE_NOTIFY_INFORMATION
    Dim lIndex  As Long
    Dim sName   As String

    ' // Walk through the notifications buffer
    Do
    
        ' // Copy to the temporary structure
        memcpy tNotify, m_bBufEvents(lIndex), Len(tNotify)
        
        ' // Retrive the file name
        sName = """" & Left$(tNotify.wcFileName, tNotify.dwFileNameLength \ 2) & """"
        
        ' // Add to list
        With lstMonitor.ListItems.Add(, , lstMonitor.ListItems.Count + 1)
        
            ' // Check the kind of the notification
            Select Case tNotify.dwAction
            Case FILE_ACTION_ADDED:             .SubItems(1) = "ADDED"              ' // File being added
            Case FILE_ACTION_REMOVED:           .SubItems(1) = "REMOVED"            ' // File being deleted
            Case FILE_ACTION_RENAMED_OLD_NAME:  .SubItems(1) = "RENAMED (old sName)" ' // File being resNamed, this is the old sName"
            Case FILE_ACTION_RENAMED_NEW_NAME:  .SubItems(1) = "RENAMED (new sName)" ' // File being resNamed, this is the new sName"
            End Select
            
            .SubItems(2) = sName
            
        End With

        ' // Walk to the next entry
        lIndex = lIndex + tNotify.dwNextEntryOffset
        
        ' // Repeat while the notifications exists
    Loop While tNotify.dwNextEntryOffset
    
    ' // Reset event
    ResetEvent hHandle
    
    ' // Fill again the OVERLAPPED structure for the asynchronous call
    m_tOvp.hEvent = hHandle
    
    ' // Start the monitor in the asynchronous mode
    Call ReadDirectoryChanges(m_hDirectory, m_bBufEvents(0), UBound(m_bBufEvents) + 1, m_iWatchSubdir, _
FILE_NOTIFY_CHANGE_FILE_NAME Or FILE_NOTIFY_CHANGE_DIR_NAME Or FILE_ACTION_ADDED Or FILE_ACTION_REMOVED, 0, m_tOvp, 0)
    
    ' // Abort the previous waiting
    m_cMonitor.Abort
    
    ' // Launch new
    m_cMonitor.vbWaitForSingleObject hHandle, INFINITE

End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)

    ' // Abort the waiting
    m_cMonitor.Abort

    ' // If the monitoring is active then abort the process
    If m_hDirectory Then
        CancelIo m_hDirectory
    End If
    
    ' // Close all handles
    CloseHandle m_hDirectory
    CloseHandle m_hEvent
    
End Sub
