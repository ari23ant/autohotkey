#UseHook  ; 無限ループ防ぐ

GroupAdd Terminal, ahk_class VTWin32  ; Tera Term
GroupAdd Terminal, ahk_class VirtualConsoleClass  ; Cmder
GroupAdd Terminal, ahk_class ConsoleWindowClass  ; Ubuntu
;GroupAdd Terminal, ahk_class CASCADIA_HOSTING_WINDOW_CLASS  ; Windows Terminal

GroupAdd MOffice, ahk_class XLMAIN  ; excel
GroupAdd MOffice, ahk_class OpusApp  ; word
GroupAdd MOffice, ahk_class PPTFrameClass ; powerpoint


; includeする場合は、デバッグコードをコメントアウトする
;#include IME.ahk  ; IME制御関数群

; Tera TermとCmderがアクティブのとき以下のスクリプトが走る
#IfWinActive, ahk_group Terminal
    ; ESCキーが押下されたとき
    Esc::
        if(IME_GET()){
            ; IME有効時 ECS→スリープ→IME無効
            Send, {ESC}
            Sleep 1
            IME_SET(0)
        } else {
            ; IME無効時 ECS
            Send, {Esc}
        }
        return

    ; Ctrl + [ が押下されたとき
    ^[::
        if(IME_GET()){
            ; IME有効時 ECS→スリープ→IME無効
            Send, {ESC}
            Sleep 1
            IME_SET(0)
        } else {
            ; IME無効時 ECS
            Send, {Esc}
        }
        return

#IfWinActive

; エクセルがアクティブのとき
;#IfWinActive, ahk_class XLMAIN
#IfWinActive, ahk_group MOffice
    ; Ctrl + Space が押下されたとき
    ^Space::
        if(IME_GET()){
            ; IME有効時→IME無効
            IME_SET(0)
        } else {
            ; IME無効時→IME有効
            IME_SET(1)
        }
        return

#IfWinActive

; https://w.atwiki.jp/eamat/pages/17.html からIME.ahkをダウンロード
; IME_GET()とIME_SET()を以下にコピペする

;-----------------------------------------------------------
; IMEの状態の取得
;   WinTitle="A"    対象Window
;   戻り値          1:ON / 0:OFF
;-----------------------------------------------------------
IME_GET(WinTitle="A")  {
	ControlGet,hwnd,HWND,,,%WinTitle%
	if	(WinActive(WinTitle))	{
		ptrSize := !A_PtrSize ? 4 : A_PtrSize
	    VarSetCapacity(stGTI, cbSize:=4+4+(PtrSize*6)+16, 0)
	    NumPut(cbSize, stGTI,  0, "UInt")   ;	DWORD   cbSize;
		hwnd := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
	             ? NumGet(stGTI,8+PtrSize,"UInt") : hwnd
	}

    return DllCall("SendMessage"
          , UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwnd)
          , UInt, 0x0283  ;Message : WM_IME_CONTROL
          ,  Int, 0x0005  ;wParam  : IMC_GETOPENSTATUS
          ,  Int, 0)      ;lParam  : 0
}

;-----------------------------------------------------------
; IMEの状態をセット
;   SetSts          1:ON / 0:OFF
;   WinTitle="A"    対象Window
;   戻り値          0:成功 / 0以外:失敗
;-----------------------------------------------------------
IME_SET(SetSts, WinTitle="A")    {
	ControlGet,hwnd,HWND,,,%WinTitle%
	if	(WinActive(WinTitle))	{
		ptrSize := !A_PtrSize ? 4 : A_PtrSize
	    VarSetCapacity(stGTI, cbSize:=4+4+(PtrSize*6)+16, 0)
	    NumPut(cbSize, stGTI,  0, "UInt")   ;	DWORD   cbSize;
		hwnd := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
	             ? NumGet(stGTI,8+PtrSize,"UInt") : hwnd
	}

    return DllCall("SendMessage"
          , UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwnd)
          , UInt, 0x0283  ;Message : WM_IME_CONTROL
          ,  Int, 0x006   ;wParam  : IMC_SETOPENSTATUS
          ,  Int, SetSts) ;lParam  : 0 or 1
}
