VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrickWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CTrickWait.cls - waiting for kernel objects asynchronously
' // © The trick, 2014-2021
' // ver. 2.0
' //

Option Explicit
Option Base 0

Private Const MODULE_NAME               As String = "CTrickWait"
Private Const MODULE_VERSION            As String = "2.0"
Private Const ERR_NOT_INITIALIZED       As String = "Object isn't initialized"
Private Const ERR_PENDING_WAIT          As String = "Object already waits"
Private Const E_NOT_VALID_STATE         As Long = &H8007139F
Private Const E_PENDING                 As Long = &H8000000A
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40&
Private Const MEM_COMMIT                As Long = &H1000&
Private Const MEM_RESERVE               As Long = &H2000&
Private Const MEM_RELEASE               As Long = &H8000&
Private Const CC_STDCALL                As Long = 4
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_MAP_WRITE            As Long = &H2
Private Const FILE_MAP_READ             As Long = &H4
Private Const FILE_MAP_EXECUTE          As Long = &H20

Private Enum eWaitingMode
    WAIT_MODE_EXIT = 0              ' // Exit request
    WAIT_WAITFORSINGLEOBJECT = 1    ' // WaitForSingleObject call
    WAIT_WAITFORMULTIPLEOBJECTS = 2 ' // WaitForMultipleObjects call
End Enum

Private Type tSThreadArg
    hHandle                     As Long
    dwTime                      As Long
End Type

Private Type tMThreadArg
    pHandles                    As Long
    dwTime                      As Long
    lCount                      As Long
    lWaitAll                    As Long
End Type

Private Type tThreadParams
    
    ' // + ------------------------------------
    cResetNotifierObject        As IUnknown     ' // To uninitialize thread use this object instance
                                                ' // When ref counter reaches to zero it uninitializes code
                                                ' // It should avoid problems with End/Stop button
                                                ' // because the runtime always releases object variables
                                                ' // so we'll have the chance to uninitialize resources
    pVtbl                       As Long
    dwRefCounter                As Long
    ' // +-------------------------------------
    
    hControl                    As Long         ' // Control event handle. This event wait for waiting operations
    hWnd                        As Long         ' // Communicate window handle
    bRegistered                 As Long         ' // If window class was registered then true
    hThread                     As Long         ' // Thread handle
    hInstance                   As Long         ' // Application hInstance
    dwWaitActive                As Long         ' // 1 - if there is active waiting operation
    cHostObject                 As CTrickWait   ' // This class instance (used to generate the event)
    pfnAPCAbort                 As Long         ' // APC abort routine address. When user call Abort method it transmits call to thread
                                                ' // a wait function returns WAIT_IO_COMPLETION
    pfnAPIs(20)                 As Long
    lIUnknownVTbl(2)            As Long
    eWaitMode                   As eWaitingMode
    bData(0 To 15)              As Byte         ' // Union tSThreadArg/tMThreadArg
    
End Type

Private Declare Function MapViewOfFile Lib "kernel32" ( _
                         ByVal hFileMappingObject As Long, _
                         ByVal dwDesiredAccess As Long, _
                         ByVal dwFileOffsetHigh As Long, _
                         ByVal dwFileOffsetLow As Long, _
                         ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" ( _
                         ByVal lpBaseAddress As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" _
                         Alias "CreateFileMappingW" ( _
                         ByVal hFile As Long, _
                         ByRef lpFileMappingAttributes As Any, _
                         ByVal flProtect As Long, _
                         ByVal dwMaximumSizeHigh As Long, _
                         ByVal dwMaximumSizeLow As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function SetEvent Lib "kernel32" ( _
                         ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByRef pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function QueueUserAPC Lib "kernel32" ( _
                         ByVal pfnAPC As Long, _
                         ByVal hThread As Long, _
                         ByVal dwData As Long) As Long
Private Declare Function SleepEx Lib "kernel32" ( _
                         ByVal dwMilliseconds As Long, _
                         ByVal bAlertable As Long) As Long
Private Declare Function InterlockedIncrement Lib "kernel32" ( _
                         ByRef lpAddend As Any) As Long
Private Declare Function InterlockedDecrement Lib "kernel32" ( _
                         ByRef lpAddend As Any) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function PutMem4 Lib "msvbvm60" ( _
                         ByRef pDst As Any, _
                         ByVal lValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Public Event OnWait( _
             ByVal hHandle As Long, _
             ByVal lResult As Long)

Private m_tParams   As tThreadParams
Private m_bIsInit   As Boolean
Private m_pCode     As Long
Private m_hMapping  As Long

' // Start waiting
Public Function vbWaitForSingleObject( _
                ByVal hHandle As Long, _
                ByVal dwMilliseconds As Long) As Boolean
                
    Const FUNCTION_NAME As String = "vbWaitForSingleObject"
    
    Dim tParam  As tSThreadArg
    
    If Not m_bIsInit Then
        Err.Raise E_NOT_VALID_STATE, MODULE_NAME & "::" & FUNCTION_NAME, ERR_NOT_INITIALIZED
    ElseIf IsActive Then
        Err.Raise E_PENDING, MODULE_NAME & "::" & FUNCTION_NAME, ERR_PENDING_WAIT
    End If
    
    tParam.hHandle = hHandle
    tParam.dwTime = dwMilliseconds
    
    memcpy m_tParams.bData(0), tParam, Len(tParam)
    
    m_tParams.eWaitMode = WAIT_WAITFORSINGLEOBJECT

    Set m_tParams.cHostObject = Me

    vbWaitForSingleObject = SetEvent(m_tParams.hControl)
    
End Function

' // Start waiting
Public Function vbWaitForMultipleObjects( _
                ByVal nCount As Long, _
                ByVal lpHandles As Long, _
                ByVal bWaitAll As Long, _
                ByVal dwMilliseconds As Long) As Boolean
    
    Const FUNCTION_NAME As String = "vbWaitForMultipleObjects"
    
    Dim tParam  As tMThreadArg
    
    If Not m_bIsInit Then
        Err.Raise E_NOT_VALID_STATE, MODULE_NAME & "::" & FUNCTION_NAME, ERR_NOT_INITIALIZED
    ElseIf IsActive Then
        Err.Raise E_PENDING, MODULE_NAME & "::" & FUNCTION_NAME, ERR_PENDING_WAIT
    End If
    
    ' // Thread is in waiting state (on hControl event)
    tParam.lCount = nCount
    tParam.pHandles = lpHandles
    tParam.dwTime = dwMilliseconds
    tParam.lWaitAll = bWaitAll
    
    ' // Move parameters
    memcpy m_tParams.bData(0), tParam, Len(tParam)
    
    ' // Set wait mode
    m_tParams.eWaitMode = WAIT_WAITFORMULTIPLEOBJECTS
    
    Set m_tParams.cHostObject = Me
    
    ' // Run wait
    vbWaitForMultipleObjects = SetEvent(m_tParams.hControl)
    
End Function

' // Check if waiting is active
Public Property Get IsActive() As Boolean
    
    Const FUNCTION_NAME As String = "IsActive_get"
    
    If m_bIsInit Then
        ' // If thread is in waiting state (on hControl event) then dwWaitActive = 0
        IsActive = m_tParams.dwWaitActive
    Else
        Err.Raise E_NOT_VALID_STATE, MODULE_NAME & "::" & FUNCTION_NAME, ERR_NOT_INITIALIZED
    End If
    
End Property

' // Abort
Public Function Abort( _
                Optional ByVal bWait As Boolean = True) As Boolean
                
    Const FUNCTION_NAME As String = "Abort"
    
    If m_bIsInit Then
        
        ' // A wait operation returns WAIT_IO_COMPLETION
        Abort = QueueUserAPC(m_tParams.pfnAPCAbort, m_tParams.hThread, 0)
        
        If bWait Then
            
            ' // Wait until thread is in waiting state on hControl event
            Do While m_tParams.dwWaitActive
                SleepEx 200, 1
            Loop
            
        End If
        
        Set m_tParams.cHostObject = Nothing
        
    Else
        Err.Raise E_NOT_VALID_STATE, MODULE_NAME & "::" & FUNCTION_NAME, ERR_NOT_INITIALIZED
    End If
    
End Function

Private Sub Class_Initialize()

    Const FUNCTION_NAME As String = "Class_Initialize"
    
    Dim hMapping    As Long
    Dim pCode       As Long
    Dim lTemp       As Long
    Dim cOp(124)    As Currency
    Dim lCodeSize   As Long
    Dim hUser32     As Long
    Dim hKernel32   As Long
    Dim hMSVBVM60   As Long
    Dim hr          As Long
    Dim vResult     As Variant
    Dim lIndex      As Long
    Dim sAPIName    As String
    
    hKernel32 = GetModuleHandle(StrPtr("kernel32"))
    hUser32 = GetModuleHandle(StrPtr("user32"))
    hMSVBVM60 = GetModuleHandle(StrPtr("msvbvm60"))
    
    ' //
    ' // The code structure
    ' //
    ' // +--------+-----------------------------------+
    ' // | offset |           description             |
    ' // +--------+-----------------------------------+
    ' // |  0x00  | Reference counter                 |
    ' // |  0x04  | Mapping address of first instance |
    ' // |  0x08  | The asm code starts here          |
    ' // |  ....  |                                   |
    ' // +--------+-----------------------------------+
    ' //
    
    lCodeSize = (UBound(cOp) + 1) * 8 + 8
    
    cOp(0) = 629327571972663.1253@:    cOp(1) = 7128988583017.0054@:      cOp(2) = 91698883875572.6731@:     cOp(3) = -405888263236301.6192@
    cOp(4) = 508315660442093.4725@:    cOp(5) = 100019506.614@:           cOp(6) = 300445632.896@:           cOp(7) = 100333680461.5561@
    cOp(8) = 596623767453601.2032@:    cOp(9) = 125791273827515.3663@:    cOp(10) = 605693046.5535@:         cOp(11) = -151320947157676.7116@
    cOp(12) = -7194103727798.5276@:    cOp(13) = 3011497924664.4342@:     cOp(14) = 2983680280480.9834@:     cOp(15) = 391955277.425@
    cOp(16) = -53949039968567.851@:    cOp(17) = 61379493.6064@:          cOp(18) = 763822158615576.6763@:   cOp(19) = 319502136290944.0512@
    cOp(20) = -147945554819252.4157@:  cOp(21) = 2983680300381.2099@:     cOp(22) = 590872271927374.9846@:   cOp(23) = -899246690272844.175@
    cOp(24) = 28241716805612.364@:     cOp(25) = 52724825746.8041@:       cOp(26) = 23996572049332.864@:     cOp(27) = 55201293.9264@
    cOp(28) = -875246440076060.0691@:  cOp(29) = 24953586902577.9844@:    cOp(30) = 58637267.7632@:          cOp(31) = 508259365447273.8445@
    cOp(32) = 510005228857239.4756@:   cOp(33) = -25290526657452.4152@:   cOp(34) = 30376919613.1328@:       cOp(35) = 904323023906891.3664@
    cOp(36) = 351680179031.2701@:      cOp(37) = 463617329542562.1364@:   cOp(38) = -900719925474018.7449@:  cOp(39) = 118641598510452.339@
    cOp(40) = 1789618030.7711@:        cOp(41) = 96940090422879.8464@:    cOp(42) = 100340551204.6335@:      cOp(43) = 502321932364934.8096@
    cOp(44) = 612490072787244.6461@:   cOp(45) = -408269983778105.6119@:  cOp(46) = 846986352420041.4277@:   cOp(47) = 235764.3016@
    cOp(48) = 176287668140559.5648@:   cOp(49) = -53950359382258.7137@:   cOp(50) = 461337076993316.992@:    cOp(51) = -6854905491806.6176@
    cOp(52) = -410462093834169.1306@:  cOp(53) = 633529803938922.4005@:   cOp(54) = 770999763531451.5816@:   cOp(55) = -7165921823302.9376@
    cOp(56) = -151195492785983.4794@:  cOp(57) = 226672015431525.4577@:   cOp(58) = 2983680280481.009@:      cOp(59) = -825838016208356.4147@
    cOp(60) = 616469818780017.4677@:   cOp(61) = -410449073482292.7648@:  cOp(62) = 114715126865841.2613@:   cOp(63) = 463617334696527.2299@
    cOp(64) = -843073850243597.7529@:  cOp(65) = 134007003440.6469@:      cOp(66) = 612490072788130.4405@:   cOp(67) = -410490181902199.0519@
    cOp(68) = 10189394156926.5734@:    cOp(69) = 405971254312225.9818@:   cOp(70) = 15340359863.3091@:       cOp(71) = 236129.5616@
    cOp(72) = -151320947233174.1323@:  cOp(73) = 7205759639922.3623@:     cOp(74) = 123813803.4805@:         cOp(75) = 16351368324.1472@
    cOp(76) = -7205759161001.6@:       cOp(77) = 8114261442450.031@:      cOp(78) = -773071132879644.2624@:  cOp(79) = 17230977617.1008@
    cOp(80) = -7205759161001.6@:       cOp(81) = 434794286573791.455@:    cOp(82) = -151251788198256.6339@:  cOp(83) = -7194100965418.3546@
    cOp(84) = 407380213303065.0966@:   cOp(85) = 868547335637019.8155@:   cOp(86) = 751828474743350.3492@:   cOp(87) = -6742315682181.8368@
    cOp(88) = 21136850940.3222@:       cOp(89) = -75629963.136@:          cOp(90) = -20350833609.8581@:      cOp(91) = 440409196.7231@
    cOp(92) = 680028187934801.6895@:   cOp(93) = 626156185509408.0201@:   cOp(94) = 351569300091.9895@:      cOp(95) = 146297851567774.106@
    cOp(96) = 269906302400977.5953@:   cOp(97) = 821437236207955.1537@:   cOp(98) = 577868051601282.2532@:   cOp(99) = -4756771624543.759@
    cOp(100) = 583456414262699.5286@:  cOp(101) = 521263402991163.8155@:  cOp(102) = 112705306370.4043@:     cOp(103) = 626873850920871.2448@
    cOp(104) = 27307.404@:             cOp(105) = 510011584541074.8159@:  cOp(106) = 147676176384734.0044@:  cOp(107) = 389339.2384@
    cOp(108) = 593611435084919.5864@:  cOp(109) = 90001741880205.9007@:   cOp(110) = 12385718884525.6833@:   cOp(111) = 689956851773394.1248@
    cOp(112) = 261336651617919.2158@:  cOp(113) = 576503968325255.45@:    cOp(114) = -441347375875332.3544@: cOp(115) = 468346669130389.9148@
    cOp(116) = -843073327109606.3228@: cOp(117) = 82557081825497.1972@:   cOp(118) = -16722218076546.5776@:  cOp(119) = 34284119689586.2783@
    cOp(120) = 713381726038840.6272@:  cOp(121) = 756615402806997.2736@:  cOp(122) = 698970537005891.2768@:  cOp(123) = 331470320374385.536@
    cOp(124) = 1.2544@:
    
    With m_tParams
        
        For lIndex = 0 To UBound(.pfnAPIs)
            
            Select Case lIndex
            
            Case 0:     sAPIName = "CreateEventW"
            Case 1:     sAPIName = "GetLastError"
            Case 2:     sAPIName = "CreateThread"
            Case 3:     sAPIName = "WaitForSingleObjectEx"
            Case 4:     sAPIName = "WaitForMultipleObjectsEx"
            Case 5:     sAPIName = "CloseHandle"
            Case 6:     sAPIName = "GlobalAlloc"
            Case 7:     sAPIName = "GlobalFree"
            Case 8:     sAPIName = "QueueUserAPC"
            Case 9:     sAPIName = "SetEvent"
        
            Case 10:    sAPIName = "RegisterClassExW"
            Case 11:    sAPIName = "CreateWindowExW"
            Case 12:    sAPIName = "UnregisterClassW"
            Case 13:    sAPIName = "MsgWaitForMultipleObjects"
            Case 14:    sAPIName = "PeekMessageW"
            Case 15:    sAPIName = "TranslateMessage"
            Case 16:    sAPIName = "DispatchMessageW"
            Case 17:    sAPIName = "PostMessageW"
            Case 18:    sAPIName = "SendMessageW"
            Case 19:    sAPIName = "DestroyWindow"
        
            Case 20:    sAPIName = "__vbaRaiseEvent"
            
            End Select
            
            If lIndex < 10 Then
                .pfnAPIs(lIndex) = GetProcAddress(hKernel32, sAPIName)
            ElseIf lIndex < 20 Then
                .pfnAPIs(lIndex) = GetProcAddress(hUser32, sAPIName)
            Else
                .pfnAPIs(lIndex) = GetProcAddress(hMSVBVM60, sAPIName)
            End If
            
            If .pfnAPIs(lIndex) = 0 Then
                Err.Raise 453, MODULE_NAME & "::" & FUNCTION_NAME
            End If
            
        Next
        
        .hInstance = App.hInstance
        
    End With
    
    hMapping = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_EXECUTE_READWRITE, 0, lCodeSize, _
                                 StrPtr(MODULE_NAME & "#" & MODULE_VERSION & "_" & CStr(GetCurrentProcessId())))
    If hMapping = 0 Then
        Err.Raise 7, MODULE_NAME & "::" & FUNCTION_NAME, "CreateFileMapping failed"
    End If
    
    pCode = MapViewOfFile(hMapping, FILE_MAP_READ Or FILE_MAP_WRITE Or FILE_MAP_EXECUTE, 0, 0, 0)
    If pCode = 0 Then
        CloseHandle hMapping
        Err.Raise 7, MODULE_NAME & "::" & FUNCTION_NAME, "MapViewOfFile failed"
    End If
    
    ' // Increment ref counter
    If InterlockedIncrement(ByVal pCode) = 1 Then
        ' // Put the address for the first instance
        PutMem4 ByVal pCode + 4, pCode
    Else
        
        ' // There is already mapped region. Use previous mapping instead current one
        lTemp = pCode
        GetMem4 ByVal pCode + 4, pCode
        UnmapViewOfFile lTemp
        CloseHandle hMapping
        hMapping = 0
        
    End If
    
    ' // Copy asm code
    memcpy ByVal pCode + 8, cOp(0), lCodeSize - 4
    
    ' // Run initialization
    hr = DispCallFunc(ByVal 0&, pCode + 8, CC_STDCALL, vbLong, 1, vbLong, VarPtr(CVar(VarPtr(m_tParams))), vResult)
    
    If hr < 0 Or vResult = 0 Then
        
        CloseHandle hMapping
        
        If InterlockedDecrement(ByVal pCode) = 0 Then
            UnmapViewOfFile pCode
        End If
        
        If hr < 0 Then
            Err.Raise hr, MODULE_NAME & "::" & FUNCTION_NAME, "DispCallFunc failed"
        Else
            Err.Raise 5, MODULE_NAME & "::" & FUNCTION_NAME, "Initialization failed"
        End If
        
    End If
    
    m_pCode = pCode
    m_hMapping = hMapping
    m_bIsInit = True
    
End Sub

Private Sub Class_Terminate()

    If m_bIsInit Then
        
        Set m_tParams.cResetNotifierObject = Nothing
        
        ' // Check if it's a last instance
        If InterlockedDecrement(ByVal m_pCode) = 0 Then
            UnmapViewOfFile m_pCode
        End If
        
        If m_hMapping Then
            CloseHandle m_hMapping
        End If
        
    End If
        
End Sub
