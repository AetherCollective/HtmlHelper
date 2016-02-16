#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Ensures only 1 copy of the script can run at any given time.
#include "Misc.au3"
_Singleton(StringTrimRight(@ScriptName, 4))

;Sets the delay at which keys are pressed and held. You may need to play with this on a per app basis. Chrome can be set as low as 1.
Opt("sendkeydelay", 1)
Opt("sendkeydowndelay", 1)

;Binds Hotkeys to Tag() function. ^ = Control, + = Shift, ! = Alt, # = Windows. See 1st reference for help.
HotKeySet("^+a", "Tag")
HotKeySet("^+b", "Tag")
HotKeySet("^+c", "Tag")
HotKeySet("^+i", "Tag")
HotKeySet("^+l", "Tag")
HotKeySet("^+o", "Tag")
HotKeySet("^+q", "Tag")
HotKeySet("^+s", "Tag")
HotKeySet("^+u", "Tag")
HotKeySet("^+z", "Tag")
HotKeySet("^!+a", "Tag")
HotKeySet("^!+c", "Tag")

;Put the script in Idle Mode.
idle()
Func idle();Idle Mode simply loops a sleep command until you press a hotkey. Once the function is done, Idle mode resumes.
	While 1
		Sleep(1000)
	WEnd
EndFunc   ;==>idle
Func Tag()
	;Store what is currently in the clipboard for later use.
	Local $Clipboard = ClipGet()
	;Empties the Clipboard; Does not check for any errors.
	Local $iRet = ClipPut("")
	;Uses multiple methods to capture selected text.
	Send("{CTRLDOWN}c{CTRLUP}")
	Send("^{x 5}")

	Local $SelectedText = ClipGet()
	TagManagement($SelectedText, $Clipboard)
EndFunc   ;==>Tag
Func TagManagement($SelectedText, $ClipboardFlag)
	;Detects Which hotkey was pressed with @hotkeypressed and prepares varibales for TagWork()

	;If @hotkeypressed = Control + Shift + A then
	If @HotKeyPressed = "^+a" Then

		;Set Variable for A Tag

		;Sets the Tag to "a"
		Local $tag = "a"

		;Sets the ClipboardFlagName to ' href="' (Note: I used single quote intentionally; If you use Double Quotes in a string, you must enclose the string with single quotes so the interpretor doesnt get confused.)
		Local $ClipboardFlagName = ' href="'

		;Sets the ClipboardFlagEnd to '"' (Note: I used single quote intentionally; If you use Double Quotes in a string, you must enclose the string with single quotes so the interpretor doesnt get confused.)
		Local $ClipboardFlagEnd = '"'

		;If there are other hotkeys
	ElseIf @HotKeyPressed = "^!+a" Then

		;same as before, skipping to reduce redundancy.
		Local $tag = "a"
		Local $ClipboardFlagName = ' href="'
		Local $ClipboardFlagEnd = '" target="_blank"'
	ElseIf @HotKeyPressed = "^+b" Then
		Local $tag = "b"
	ElseIf @HotKeyPressed = "^+c" Then
		Local $tag = "code"
	ElseIf @HotKeyPressed = "^!+c" Then
		Local $tag = "cited"
	ElseIf @HotKeyPressed = "^+i" Then
		Local $tag = "i"
	ElseIf @HotKeyPressed = "^+l" Then
		Local $tag = "li"
	ElseIf @HotKeyPressed = "^+o" Then
		Local $tag = "ol"
	ElseIf @HotKeyPressed = "^+q" Then
		Local $tag = "blockquote"
	ElseIf @HotKeyPressed = "^+s" Then
		Local $tag = "strong"
	ElseIf @HotKeyPressed = "^+u" Then
		Local $tag = "u"
	ElseIf @HotKeyPressed = "^+z" Then
		;Empties the Clipboard
		Local $iRet = ClipPut("")

		;If Clipboard could not be Emptied, Show an Error Message.
		If Not $iRet Then MsgBox(0, StringTrimRight(@ScriptName, 4), "Clipboard could not be emptied.")

		;Gets Mouse Position and stores it to a variable to be later called in a ToolTip
		Local $mousepos = MouseGetPos()

		;Creates a Tooltip at an offset to the cursor location informing the user the Clipboard has been emptied.
		ToolTip("Clipboard Emptied", $mousepos[0] + 2, $mousepos[1] - 9, "", "", 4)

		;Wait 500 ms
		Sleep(500)

		;Removes the previously created tooltip.
		ToolTip("")

		;Put the script in Idle Mode.
		idle()
	EndIf

	;If $tag OR $ClipboardFlagName OR $ClipboardFlagEnd are not defined, defines them with a blank string.
	If Not IsDeclared("tag") Then Local $tag = ""
	If Not IsDeclared("ClipboardFlagName") Then Local $ClipboardFlagName = ""
	If Not IsDeclared("ClipboardFlagEnd") Then Local $ClipboardFlagEnd = ""

	;;Formats the Tag you requested and pastes it at your cursor location.
	TagWork($SelectedText, $tag, $ClipboardFlagName, $ClipboardFlag, $ClipboardFlagEnd)

	;If you need multiple flags, See the following Example
	;If Not IsDeclared("ClipboardFlagName2") Then Local $ClipboardFlagName2 = ""    ;;This is an Example for using 2 Flags
	;If Not IsDeclared("ClipboardFlag2") Then $ClipboardFlag2 = ""    ;;This is an Example for using 2 Flags
	;If Not IsDeclared("ClipboardFlagEnd2") Then Local $ClipboardFlagEnd2 = ""    ;;This is an Example for using 2 Flags
	;TagWork($SelectedText, $tag, $ClipboardFlagName, $ClipboardFlag, $ClipboardFlagEnd,$ClipboardFlagName2, $ClipboardFlag2, $ClipboardFlagEnd2)
EndFunc   ;==>TagManagement
Func TagWork($SelectedText, $tag, $ClipboardFlagName, $ClipboardFlag, $ClipboardFlagEnd)
	;Formats the Tag you requested and pastes it at your cursor location.

	;If you need multiple flags, you will need to modify your the above line so it can make use of your flags. See the Example Below.
	;Func TagWork($SelectedText, $tag, $ClipboardFlagName, $ClipboardFlag, $ClipboardFlagEnd,$ClipboardFlagName2, $ClipboardFlag2, $ClipboardFlagEnd2)


	;Creates & Prepares the $ClipboardStorage Variable for use.
	;An example of how this would looks if A Tag was used & $selectedtext="Hello World" & $clipboardFlag="http://example.com/"
	;<a href="http://example.com/">Hello World</a>
	;An example of how this would looks if B Tag was used & $selectedtext="Hello World":
	;<b>Hello World</b>
	Local $ClipboardStorage = '<' & $tag & $ClipboardFlagName & $ClipboardFlag & $ClipboardFlagEnd & '>' & $SelectedText & "</" & $tag & ">"
	;; This is an Example for using 2 Flags.
	;Local $ClipboardStorage = '<' & $tag & $ClipboardFlagName & $ClipboardFlag & $ClipboardFlagEnd & $ClipboardFlagName2 & $ClipboardFlag2 & $ClipboardFlagEnd2 & '>' & $SelectedText & "</" & $tag & ">"

	;Overwrite the Clipboard's contents with the contents from variable $ClipboardStorage
	ClipPut($ClipboardStorage)

	;Pastes the Contents of the clipboard at the cursor location.
	Send("^v")

	;If no text was selected, then move cursor to the middle of the tags. <b>{HERE}</b>
	If $SelectedText = "" Then Send("{shiftdown}{shiftup}^{left}{left}{left}")

	;Empties the Clipboard; Stores the results in $iRet
	Local $iRet = ClipPut("")

	;If Clipboard could not be Emptied, Show an Error Message.
	If $iRet = False Then MsgBox(0, StringTrimRight(@ScriptName, 4), "Clipboard could not be emptied.")
EndFunc   ;==>TagWork

;;Manuals to Reference:
;https://www.autoitscript.com/autoit3/docs/functions/Send.htm
;https://www.autoitscript.com/autoit3/docs/intro/lang_operators.htm
;https://www.autoitscript.com/autoit3/docs/intro/lang_variables.htm
