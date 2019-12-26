#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoIt_Main_v11_256x256_RGB-A.ico
#AutoIt3Wrapper_Res_Comment=Вставка текущей/выбранной даты в спецформате
#AutoIt3Wrapper_Res_Language=1049
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Alex Perfiliev

 Script Function:


#ce ----------------------------------------------------------------------------

#include <Misc.au3>
#Include <Constants.au3>
#include <Date.au3>

#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

Local $hDLL = DllOpen("user32.dll")

Opt("WinTitleMatchMode", 2)
Opt("SendKeyDelay", 50)

Opt("TrayMenuMode",1)
TraySetState()
;TraySetToolTip("Alt+Средняя клавиша мыши - вставить текущую дату в формате 'yyyymmdd-01'")
TraySetToolTip("Вставка текущей даты")

$helpitem = TrayCreateItem("Помощь")
TrayCreateItem("")
$exititem = TrayCreateItem("Выход")

While 1

	;Alt+Средняя клавиша мыши - выводит в активное текстовое поле текущую дату в формате + 'yyyymmdd-01' (пример: 20191108-01)

	If _IsPressed("04", $hDLL) And _IsPressed("12", $hDLL) Then

		While _IsPressed("04", $hDLL) Or _IsPressed("12", $hDLL)
			Sleep(10)
		WEnd

		AltMButton()

	EndIf

	;Ctrl+Средняя клавиша мыши - выбирает и выводит в активное текстовое поле текущую дату в формате 'dd.mm.yyyy, wd' (пример: 08.11.2019, Пт)

	If _IsPressed("04", $hDLL) And _IsPressed("11", $hDLL) Then

		While _IsPressed("04", $hDLL) Or _IsPressed("11", $hDLL)
			Sleep(10)
		WEnd

		CtrlMButton()

	EndIf

	;Shift+Средняя клавиша мыши - выводит в активное текстовое поле текущую дату и время (пример: 08.11.2019, 13:37:56)

	If _IsPressed("04", $hDLL) And _IsPressed("10", $hDLL) Then

		While _IsPressed("04", $hDLL) Or _IsPressed("10", $hDLL)
			Sleep(10)
		WEnd

		ShiftMButton()

	EndIf

;~  	Sleep(100)

    $msg = TrayGetMsg()
    Select
        Case $msg = 0
            ContinueLoop
        Case $msg = $helpitem

			ShowHelpWindows()

        Case $msg = $exititem
            ExitLoop
    EndSelect

WEnd

Exit

Func ShowHelpWindows()

	#Region ### START Koda GUI section ### Form=
	$Form1 = GUICreate("Вставка даты в разных форматах:", 665, 87, -1, -1, -1, BitOR($WS_EX_TOOLWINDOW,$WS_EX_WINDOWEDGE))
	GUISetFont(10, 400, 0, "MS Sans Serif")
	$Label1 = GUICtrlCreateLabel("Alt + Средняя клавиша мыши", 16, 8, 210, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x008000)
	$Label2 = GUICtrlCreateLabel("Ctrl + Средняя клавиша мыши", 16, 32, 215, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x008000)
	$Label3 = GUICtrlCreateLabel("Shift + Средняя клавиша мыши", 16, 56, 222, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x008000)
	$Label4 = GUICtrlCreateLabel("- текущая дата, формат 'yyyymmdd-01' (пример: 20191108-01)", 248, 8, 500, 20)
	$Label5 = GUICtrlCreateLabel("- выбрать дату, формат 'dd.mm.yyyy, wd' (пример: 08.11.2019, Пт)", 248, 32, 500, 20)
	$Label6 = GUICtrlCreateLabel("- текущая дата, дата и время (пример: 08.11.2019, 13:37:56)", 248, 56, 500, 20)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete();
				ExitLoop

		EndSwitch
	WEnd

EndFunc

Func ShiftMButton()

	Opt("SendKeyDelay", 1)
	_SendEx(@MDAY & "." & @MON & "." & @YEAR & ", " & @HOUR & ":" & @MIN & ":" & @SEC)
	Opt("SendKeyDelay", 50)

EndFunc

Func CtrlMButton()

	Local $aMyDate, $aMyTime

	$hWnd = WinGetHandle("[ACTIVE]")

	$PreselectDate = @YEAR & "/" & @MON & "/" & @MDAY

	$pos = MouseGetPos()

	#Region ### START Koda GUI section ### Form=D:\AutoIt\InsDate\FormCalendar.kxf

	$Form1_1 = GUICreate("Выбери", 182, 179, $pos[0], $pos[1], BitOR($GUI_SS_DEFAULT_GUI,$DS_MODALFRAME), BitOR($WS_EX_TOOLWINDOW,$WS_EX_TOPMOST,$WS_EX_WINDOWEDGE))
	$MonthCal1 = GUICtrlCreateMonthCal($PreselectDate, 8, 8, 166, 164)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				$CloseCal = True
				ExitLoop

			Case $MonthCal1
				$CloseCal = False
				ExitLoop
		EndSwitch
	WEnd

	$SelDate = GUICtrlRead($MonthCal1)
	$rez = _DateTimeFormat($SelDate, 2)

	GUIDelete();

	If $CloseCal Then
		Return
	EndIf

	_DateTimeSplit($SelDate, $aMyDate, $aMyTime)

	$DOW = _DateToDayOfWeek($aMyDate[1], $aMyDate[2], $aMyDate[3])

	Select
		Case $DOW = 1
			$sShortDayName = "Вс"
		Case $DOW = 2
			$sShortDayName = "Пн"
		Case $DOW = 3
			$sShortDayName = "Вт"
		Case $DOW = 4
			$sShortDayName = "Ср"
		Case $DOW = 5
			$sShortDayName = "Чт"
		Case $DOW = 6
			$sShortDayName = "Пт"
		Case $DOW = 7
			$sShortDayName = "Сб"
	EndSelect

	WinActivate($hWnd)

	Opt("SendKeyDelay", 1)
	_SendEx($rez & ", " & $sShortDayName)
	Opt("SendKeyDelay", 50)

EndFunc

Func AltMButton()

	Opt("SendKeyDelay", 1)
	Send(@YEAR & @MON & @MDAY & "-01")
	Opt("SendKeyDelay", 50)

EndFunc

;Автор: CreatoR
;Интерпритация на функцию Send(), только с использованием б.обмена - обход проблемы с кодировками
Func _SendEx($sString)
	Local $sOld_Clip = ClipGet()

	ClipPut($sString)
	Sleep(10)
	Send("+{INSERT}")

	ClipPut($sOld_Clip)
EndFunc
