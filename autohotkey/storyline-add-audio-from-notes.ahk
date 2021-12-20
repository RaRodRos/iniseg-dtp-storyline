#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
CoordMode, Mouse, Screen
CoordMode, Tooltip, Screen
SetMouseDelay, 0

Loop {
	InputBox
		,NumDiapos
		,Número de diapositivas
		,¿Cuántas diapositivas tiene el documento?
		,,Width
		,Height
		,X
		,Y
		,Locale
		,Timeout
	If ErrorLevel
		ExitApp
	If NumDiapos is Integer
		Break
}

Loop, %NumDiapos% {
	ToolTip, % A_Index . "/" . NumDiapos . "|" 
		. Floor(((NumDiapos - A_Index + 1) * 9.75)/60) . "min & "
		. Round(Mod((NumDiapos - A_Index + 1) * 9.75,60)) . " seg"
		, 1, 1

	; Click ribbon Insertar
	Sleep, 100
	Click, 160 42

	; Click botón Audio
	Sleep, 100
	Click, 656 114

	; Click Texto a Voz
	Sleep, 100
	Click, 655 198

	; Click Copiar de la Nota
	Sleep, 1000
	Click, 951 570

	; Enter para Aceptar
	Sleep, 350
	Send, {Enter}

	; Esc por si no hay nota
	Sleep, 100
	Send, {Esc}

	; Flecha abajo para ir a siguiente diapositiva
	Sleep, 8000
	Send, {Down}
}

ExitApp

^Esc:: ExitApp
Pause:: Pause