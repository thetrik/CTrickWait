;
; CTrickWait assembly code
; By The trick 2021
; FASM compiler
;

format binary
use32

include "win32wx.inc"

WAIT_MODE_EXIT equ 0
WAIT_WAITFORSINGLEOBJECT equ 1
WAIT_WAITFORMULTIPLEOBJECTS equ 2

struct tST
    hHandle	dd ?
    dwTime	dd ?
ends

struct tMT
    pHandles	dd ?
    dwTime	dd ?
    dwCount	dd ?
    dwWaitAll	dd ?
ends


struct tThreadParams
    pResetNotifierObject	   dd ?  ; When this object ref counter reaches zero it uninitializes all
    pVtbl			   dd ?
    dwRefCounter		   dd ?

    hControl			   dd ?  ; To increase performance all the waiting operations are performed in a static thread
					 ; This event handle is used to manage waiting requests from main thread. So we don't
					 ; need to create a separate thread for each request. This reduces system resources usage
    hWnd			   dd ?
    bRegistered 		   dd ?
    hThread			   dd ?
    hInstance			   dd ?
    dwWaitActive		   dd ?
    pHostObject 		   dd ?
    pfnAPCAbort 		   dd ?

    pfnCreateEventW		   dd ?
    pfnGetLastError		   dd ?
    pfnCreateThread		   dd ?
    pfnWaitForSingleObjectEx	   dd ?
    pfnWaitForMultipleObjectsEx    dd ?
    pfnCloseHandle		   dd ?
    pfnGlobalAlloc		   dd ?
    pfnGlobalFree		   dd ?
    pfnQueueUserAPC		   dd ?
    pfnSetEvent 		   dd ?

    pfnRegisterClassEx		   dd ?
    pfnCreateWindowEx		   dd ?
    pfnUnregisterClass		   dd ?
    pfnMsgWaitForMultipleObjects   dd ?
    pfnPeekMessageW		   dd ?
    pfnTranslateMessage 	   dd ?
    pfnDispatchMessageW 	   dd ?
    pfnPostMessageW		   dd ?
    pfnSendMessageW		   dd ?
    pfnDestroyWindow		   dd ?

    pfnvbaRaiseEvent		   dd ?

    pfnQI			   dd ?  ; CResetNotifier virtual functions table
    pfnAddRef			   dd ?
    pfnRelease			   dd ?

    dwWaitMode			   dd ?
    union
       tS tST
       tM tMT
    ends
ends

virtual at 0
  call initialize   ; disable removing proc
end virtual

proc initialize uses esi edi, pParams
    locals
	tCls WNDCLASSEX
	bRegistered db 0
	bResult db 0
    endl

    mov esi, [pParams]
    lea edi, [tCls]
    mov ecx, sizeof.WNDCLASSEX / 4
    xor eax, eax
    rep stosd

    mov [tCls.cbSize], sizeof.WNDCLASSEX
    mov eax, [esi + tThreadParams.hInstance]
    mov [tCls.hInstance], eax
    call @f
    @@: pop edi
    lea eax, [edi + wnd_proc - @b]
    mov [tCls.lpfnWndProc], eax
    lea eax, [edi + WND_CLASS_NAME - @b]
    mov [tCls.lpszClassName], eax

    stdcall [esi + tThreadParams.pfnRegisterClassEx], addr tCls

    .if eax = 0
	stdcall [esi + tThreadParams.pfnGetLastError]
	.if eax <> 1410 ; ERROR_CLASS_ALREADY_EXISTS
	    jmp .exit_proc
	.endif
    .else
	mov [bRegistered], 1
    .endif

    stdcall [esi + tThreadParams.pfnCreateWindowEx], 0, addr edi + WND_CLASS_NAME - @b, 0, 0, 0, 0, 0, 0, -3, 0, [esi + tThreadParams.hInstance], 0

    .if eax = 0
	jmp .exit_proc
    .else
	mov [esi + tThreadParams.hWnd], eax
    .endif

    stdcall [esi + tThreadParams.pfnCreateEventW], 0, 0, 0, 0

    .if eax = 0
	jmp .exit_proc
    .else
	mov [esi + tThreadParams.hControl], eax
    .endif

    stdcall [esi + tThreadParams.pfnCreateThread], 0, 0, addr edi + thread_proc - @b, esi, 0, 0

    .if eax = 0
	jmp .exit_proc
    .else
	mov [esi + tThreadParams.hThread], eax
    .endif

    lea eax, [edi + abort - @b]
    mov [esi + tThreadParams.pfnAPCAbort], eax

    lea eax, [edi + CResetNotifier_QueryInterface - @b]
    mov [esi + tThreadParams.pfnQI], eax

    lea eax, [edi + CResetNotifier_AddRef - @b]
    mov [esi + tThreadParams.pfnAddRef], eax

    lea eax, [edi + CResetNotifier_Release - @b]
    mov [esi + tThreadParams.pfnRelease], eax

    lea eax, [esi + tThreadParams.pfnQI]
    mov [esi + tThreadParams.pVtbl], eax
    lea eax, [esi + tThreadParams.pVtbl]
    mov [esi + tThreadParams.pResetNotifierObject], eax

    mov [esi + tThreadParams.dwRefCounter], 1

    .if [bRegistered]
	mov [esi + tThreadParams.bRegistered], 1
    .endif

    mov [bResult], 1

  .exit_proc:

    .if ~[bResult]

	.if [esi + tThreadParams.hControl]
	    stdcall [esi + tThreadParams.pfnCloseHandle], [esi + tThreadParams.hControl]
	    mov [esi + tThreadParams.hControl], 0
	.endif

	.if [esi + tThreadParams.hWnd]
	    stdcall [esi + tThreadParams.pfnDestroyWindow], [esi + tThreadParams.hWnd]
	    mov [esi + tThreadParams.hWnd], 0
	.endif

	.if [bRegistered]
	    stdcall [esi + tThreadParams.pfnUnregisterClass], addr edi + WND_CLASS_NAME - @b, [esi + tThreadParams.hInstance]
	.endif

    .endif

    movzx eax, [bResult]

    ret

endp

; Uninitialization
proc uninitialize uses esi edi, pParams
   locals
       tMsg MSG
       hr dd 0;
   endl

   mov esi, [pParams]

   mov [esi + tThreadParams.dwWaitMode], WAIT_MODE_EXIT ; Exit thread mode

   stdcall [esi + tThreadParams.pfnQueueUserAPC], [esi + tThreadParams.pfnAPCAbort], [esi + tThreadParams.hThread], 0  ; Abort a pending operation

   .if eax = 0
       mov [hr], 0x80004005
       jmp .exit_proc
   .endif

   stdcall [esi + tThreadParams.pfnSetEvent], [esi + tThreadParams.hControl] ; Wake thread

   .if eax = 0
       mov [hr], 0x80004005
       jmp .exit_proc
   .endif

   ; Because of the window lives in current thread we need to pump the messages from the thread to finish uninitialization

   .loop:

       stdcall [esi + tThreadParams.pfnMsgWaitForMultipleObjects], 1, addr esi + tThreadParams.hThread, 0, -1, 0x5FF

       .if eax = 0
	   jmp .end_loop
       .elseif eax = 1

	   stdcall [esi + tThreadParams.pfnPeekMessageW], addr tMsg, 0, 0, 0, 1
	   stdcall [esi + tThreadParams.pfnTranslateMessage], addr tMsg
	   stdcall [esi + tThreadParams.pfnDispatchMessageW], addr tMsg

       .else
	   mov [hr], 0x80004005
	   jmp .exit_proc
       .endif

    jmp .loop

  .end_loop:

    stdcall [esi + tThreadParams.pfnCloseHandle], [esi + tThreadParams.hThread]

    mov [esi + tThreadParams.hThread], 0

  .exit_proc:

    mov eax, [hr]

    ret

endp

; When a waiting operation is active the APC requests break this. So we can break any waiting operation
proc abort pParams
    ret
endp

proc thread_proc uses esi edi, pParams

    mov esi, [pParams]			     ; Params

  .main_loop:

    mov [esi + tThreadParams.dwWaitActive], 0

    stdcall [esi + tThreadParams.pfnWaitForSingleObjectEx], [esi + tThreadParams.hControl], -1, 1

    .if eax = 0

	; New wait request
	.if [esi + tThreadParams.dwWaitMode] = WAIT_MODE_EXIT
	    ; Exit thread
	    jmp .exit_thread
	.elseif [esi + tThreadParams.dwWaitMode] = WAIT_WAITFORSINGLEOBJECT
	    ; WaitForSingleObjectEx
	    mov [esi + tThreadParams.dwWaitActive], 1
	    stdcall [esi + tThreadParams.pfnWaitForSingleObjectEx], [esi + tThreadParams.tS.hHandle], [esi + tThreadParams.tS.dwTime], 1
	.else
	    ; WaitForMultipleObjectsEx
	    mov [esi + tThreadParams.dwWaitActive], 1
	    stdcall [esi + tThreadParams.pfnWaitForMultipleObjectsEx], [esi + tThreadParams.tM.dwCount], [esi + tThreadParams.tM.pHandles], \
		    [esi + tThreadParams.tM.dwWaitAll], [esi + tThreadParams.tM.dwTime], 1
	.endif

	.if eax = 0xC0 ; WAIT_IO_COMPLETION
	    ; Abort request
	    jmp .main_loop
	.endif

	mov edi, eax

	; Allotcate memory for parameters
	stdcall [esi + tThreadParams.pfnGlobalAlloc], 0, 8

	.if eax
	    mov [eax], edi
	    mov edi, [esi + tThreadParams.tS.hHandle]
	    mov [eax + 4], edi
	.else
	    ;error
	    jmp .exit_thread
	.endif

	stdcall [esi + tThreadParams.pfnPostMessageW], [esi + tThreadParams.hWnd], 0x400, esi, eax

    .elseif eax = 0xC0	; WAIT_IO_COMPLETION
	; Abort request
	jmp .main_loop
    .else
	; Error
	jmp .exit_thread
    .endif

  .end_loop:

    jmp .main_loop

  .exit_thread:

    ; Destroy window request in main thread
    stdcall [esi + tThreadParams.pfnSendMessageW], [esi + tThreadParams.hWnd], 0x401, esi, esi

    ret

endp

proc wnd_proc uses esi edi, hWnd, uMsg, wParam, lParam

    mov esi, [wParam]

    .if [uMsg] = 0x400

	mov ecx, [lParam]

	push ecx

	; Because an Abort request can be placed after wait operation has requested pHostObject can contains NULL
	; So we need to check this condition. There is no race condition because this procedure is called in main thread
	mov edi, [esi + tThreadParams.pHostObject]

	.if edi

	    ; Check if aborted
	    xor eax, eax

	    ; Zero object before event because an event handler can modify this field
	    mov [esi + tThreadParams.pHostObject], eax

	    push eax
	    push dword [ecx + 4]
	    push eax
	    push 3
	    push eax
	    push dword [ecx]
	    push eax
	    push 3
	    push 2
	    push 1
	    push edi
	    call [esi + tThreadParams.pfnvbaRaiseEvent]
	    add esp, 44

	    ; Release
	    mov eax, edi
	    push eax
	    mov eax, [eax]
	    call dword [eax + 8]

	.endif

	call [esi + tThreadParams.pfnGlobalFree]

    .elseif [uMsg] = 0x401

	; This is called from Class_terminate event
	stdcall [esi + tThreadParams.pfnDestroyWindow], [hWnd]
	mov [esi + tThreadParams.hWnd], 0

	stdcall [esi + tThreadParams.pfnCloseHandle], [esi + tThreadParams.hControl]
	mov [esi + tThreadParams.hControl], 0

	.if [esi + tThreadParams.bRegistered]
	    call @f
	    @@: pop eax
	    stdcall [esi + tThreadParams.pfnUnregisterClass], addr eax + WND_CLASS_NAME - @b, [esi + tThreadParams.hInstance]
	.endif

    .elseif [uMsg] = 0x81      ; WM_NCCREATE
	mov eax, 1
    .else
	xor eax, eax
    .endif

    ret
endp

CResetNotifier_QueryInterface:

    mov eax, [esp + 0x04]
    mov ecx, [esp + 0x0c]
    mov [ecx], eax
    stdcall CResetNotifier_AddRef, eax
    xor eax, eax
    ret 0x0c

CResetNotifier_AddRef:

    mov eax, [esp + 0x04]
    inc dword [eax + 0x04]
    mov eax, [eax + 0x04]
    ret 0x04

CResetNotifier_Release:

    mov eax, [esp + 0x04]
    dec dword [eax + 0x04]

    .if ZERO?

	push eax
	sub eax, 4
	stdcall uninitialize, eax
	pop eax

    .endif

    mov eax, [eax + 0x04]

    ret 0x04

WND_CLASS_NAME du "TrickWaitClass1.1", 0
