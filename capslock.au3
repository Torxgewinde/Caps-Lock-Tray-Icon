#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=capslock_unknown.ico
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Run_Au3Stripper=y

#AutoIt3Wrapper_Res_Icon_Add=capslock_on.ico 		;201
#AutoIt3Wrapper_Res_Icon_Add=capslock_off.ico		;202
#AutoIt3Wrapper_Res_Icon_Add=capslock_unknown.ico	;203

#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <misc.au3>
#include <TrayConstants.au3>

Dim $tooglecaps = TrayCreateItem("Toggle Caps Lock")
Dim $capson = TrayCreateItem("Caps Lock On")
Dim $capsoff = TrayCreateItem("Caps Lock Off")
Dim $capsexit = TrayCreateItem("Exit")

Global Const $VK_CAPITAL = 0x14
Dim $WshShell = ObjCreate("WScript.Shell")
Dim $hUser32DLL = DllOpen('User32.dll')

Dim $CapsLockLastState
Dim $FirstRun = True

Func HandleTrayMessage($trayMsg)
	Switch $trayMsg
		Case $tooglecaps
			ConsoleWrite("toggle caps" & @CRLF)
			$WshShell.SendKeys("{CAPSLOCK}")

		Case $capsoff
			ConsoleWrite("caps off" & @CRLF)

			If _Key_Is_On($VK_CAPITAL) Then
			   $WshShell.SendKeys("{CAPSLOCK}")
			EndIf

		Case $capson
			ConsoleWrite("caps on" & @CRLF)

			If Not _Key_Is_On($VK_CAPITAL) Then
			   $WshShell.SendKeys("{CAPSLOCK}")
			EndIf

		Case $capsexit
			ConsoleWrite("caps exit" & @CRLF)
			Exit

		 Case Else

	EndSwitch
EndFunc

Func OnAutoItExit()
   ConsoleWrite("leaving now, cleaning up..." & @CRLF)
   $WshShell = 0
   DllClose($hUser32DLL)
 EndFunc

Func _Key_Is_On($nVK_KEY)
    Local $a_Ret = DllCall($hUser32DLL, "short", "GetKeyState", "int", $nVK_KEY)
    Return Not @error And BitAND($a_Ret[0], 0xFF) = 1
EndFunc

OnAutoItExitRegister("OnAutoItExit")
Opt("TrayMenuMode", 3)
TraySetToolTip("Caps Lock")
TraySetIcon(@ScriptFullPath, 203)

While True
   HandleTrayMessage(TrayGetMsg())

   Local $CapsLockState = _Key_Is_On($VK_CAPITAL)

   If $CapsLockState And (Not $CapsLockLastState Or $FirstRun) Then
	  ConsoleWrite("Capslock changed to ON" & @CRLF)
	  TraySetIcon(@ScriptFullPath, 201)
	  $CapsLockLastState = True
   EndIf

   If Not $CapsLockState And ($CapsLockLastState Or $FirstRun) Then
	  ConsoleWrite("Capslock changed to OFF" & @CRLF)
	  TraySetIcon(@ScriptFullPath, 202)
	  $CapsLockLastState = False
   EndIf

   Sleep(100)
   $FirstRun = False
WEnd
