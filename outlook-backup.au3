#include <Misc.au3>
#include <GuiListBox.au3>
#include <GuiRichEdit.au3>
#include <MsgBoxConstants.au3>

Opt("WinTitleMatchMode", 3)

If _Singleton("outlook-backup", 1) = 0 Then Exit

;~ Customize those
Local $AccountName   = "enter@yourmail.here"
Local $BackupPath    = "C:\enter@yourmail.pst"
;~ And verify this
Local $OutlookPath   = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
;~ But don't touch these (probably)
Local $WinOutlook    = "[CLASS:rctrl_renwnd32]"
Local $WinExport     = "Import and Export Wizard"
Local $LblFileExport = "Export to a file"
Local $WinFileType   = "Export to a File"
Local $LblFileType   = "Outlook Data File (.pst)"
Local $WinExportFile = "Export Outlook Data File"

If Not WinExists($WinOutlook) Then
	Local $YesNo = MsgBox($MB_ICONQUESTION + $MB_YESNO, "Outlook not detected", "Outlook does not seem to run, launch it?")
	If $YesNo = $IDYES Then
		Run($OutlookPath)
		WinWait($WinOutlook, "", 5)
		;~ Wait for the main windows to be displayed (as opposed to the launch popup)
		While (ControlCommand($WinOutlook, "", "NetUIHWND1", "IsVisible") = 0)
			Sleep(1000)
		WEnd
	Else
		Exit
	EndIf
EndIf

WinActivate($WinOutlook)
Send("!foi")

;~ IMPORT/EXPORT dialog
WinWait($WinExport, "", 1)
Local $hListBox = ControlGetHandle($WinExport, "", "ListBox1")
Local $idx = _GUICtrlListBox_FindString ($hListBox, $LblFileExport, True)
_GUICtrlListBox_SetCurSel($hListBox, $idx)
ControlClick($WinExport, "", "Button3")

;~ FILE TYPE dialog
WinWait($WinFileType, "", 1)
$hListBox = ControlGetHandle($WinFileType, "", "ListBox1")
$idx = _GUICtrlListBox_FindString($hListBox, $LblFileType, True)
_GUICtrlListBox_SetCurSel($hListBox, $idx)
ControlClick($WinFileType, "", "Button3")

;~ ACCOUNT CHOICE dialog
WinWait($WinExportFile, "", 1)
ControlTreeView($WinExportFile, "", "SysTreeView321", "Select", $AccountName)
;~  Make sure Include subfolders is checked
ControlCommand($WinExportFile, "", "Button1", "Check")
ControlClick($WinExportFile, "", "Button4")

;~ SAVE TO dialog
WinWait($WinExportFile, "", 1)
Local $hRichEdit = ControlGetHandle($WinExportFile, "", "RichEdit20WPT1")
ControlSetText($WinExportFile, "", "RichEdit20WPT1", $BackupPath)
;~  Make sure Allow duplicate items is checked
ControlCommand($WinExportFile, "", "Button4", "Check")
ControlClick($WinExportFile, "", "Button10")
