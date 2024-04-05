#SingleInstance Ignore
VarAssetCheckbox = 0
FontSize = 64


whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("GET", "https://raw.githubusercontent.com/supernovatroubling/lighting-node-pro-docker/master/activation", true)
whr.Send()
; Using 'true' above and the call below allows the script to remain responsive.
whr.WaitForResponse()
version := whr.ResponseText

if ( version = 0 )
{
	ExitApp
} else {


^g::


Start:
SetTitleMatchMode, 1 ; Set title match to partial match
WinGetClass, VarClass, WE3 AMS Lite
AMSClassArray := StrSplit(VarClass, ".")
AppForm := AMSClassArray[1]
AppNumber := AMSClassArray[4] "." AMSClassArray[5] "." AMSClassArray[6]
if WinExist("WE3 AMS Lite")
{
	WinActivate, ahk_exe WE3OffSiteAssetManager.exe	
	WinWaitActive, WE3 AMS Lite,,5
	ControlGetText, VarProduct, %AppForm%.Window.8.%AppNumber%16, WE3 AMS Lite
	if VarProduct = Add Product
		MsgBox, Please select a product
	else 
		VarProduct = % SubStr(VarProduct, 14, StrLen(VarProduct))
		
	Gui, SerialGui:New, AlwaysOnTop, Serial Number
	Menu, SerialMenuBar, Add
	Menu, SerialMenuBar, deleteAll

	Menu, Submenu_font, Add, &Size, FontSizeMenuButton
	Menu, SerialMenuBar, Add, &Font, :Submenu_font

	Gui, Menu, SerialMenuBar
	Gui, Font, s%FontSize%
	Gui, Add, Text, Center, Selected Product:
	Gui, Add, Text, Center, %VarProduct%
	If (SanitizedSerial != "")
		Gui, Add, Text,, Previous serial number was:`n%SanitizedSerial%
	Gui, Add, Text, Center, Enter serial number:
	Gui, Add, Edit,r1 vVarSerial Uppercase -WantReturn
	Gui, Add, Checkbox, Checked%VarAssetCheckbox% vVarAssetCheckbox, Asset Tag 
	Gui, Add, Button, Default, &OK
	Gui, Add, Button, x+m, &Cancel
	Gui, Show, AutoSize Center
	
	WinWaitClose, Serial Number
	SanitizedSerial:= StrReplace(VarSerial, "  ")
	If (VarAssetCheckbox = 1)
	{
		Gui, AssetGui:New, AlwaysOnTop, Asset Number
		Gui, Font, s%FontSize%
		GUI, Add, Text, Center, Selected Product:
		GUI, Add, Text, Center, %VarProduct%
		Gui, Add, Text,, Serial number was:`n%SanitizedSerial%
		Gui, Add, Text,, Enter asset number:
		Gui, Add, Edit,r1 vVarAsset Uppercase -WantReturn
		Gui, Add, Button, Default, &OK
		Gui, Add, Button, x+m, &Cancel
		Gui, Show, AutoSize Center
		WinWaitClose, Asset Number
	}
	ControlFocus, %AppForm%.EDIT.%AppNumber%3, WE3 AMS Lite
	if ErrorLevel
	{
		MsgBox,, Timeout, Timed out waiting for We3 Offsite Asset Manager to become active
		return
	}
	else
	{
		BlockInput On
		ControlSetText, %AppForm%.EDIT.%AppNumber%3, %SanitizedSerial%, WE3 AMS Lite ; Set text of Serial No Box to the sanitized Serial entered
		sleep 50
		ControlGet ,VarEnabled, Enabled,, %AppForm%.EDIT.%AppNumber%1, WE3 AMS Lite ; Check if Cust Asset No is enabled
		If (VarAssetCheckbox = 1) ; check if GUI checkbox checked 
			{
			Switch VarEnabled 
			{
			Case 0: ; check if Record Asset No is not checked
				ControlClick, %AppForm%.BUTTON.%AppNumber%1, WE3 AMS Lite
				Sleep, 100
				ControlSetText, %AppForm%.EDIT.%AppNumber%1, %VarAsset%, WE3 AMS Lite ; Set text of Customer Asset No Box to the asset entered
			Case 1: ; check if Record Asset No is checked
				ControlSetText, %AppForm%.EDIT.%AppNumber%1, %VarAsset%, WE3 AMS Lite ; Set text of Customer Asset No Box to the asset entered
			Default:
				MsgBox, Error getting Cust Asset No State
			}
			ControlFocus, %AppForm%.EDIT.%AppNumber%1, WE3 AMS Lite
		}
		Else If (VarAssetCheckbox = 0) ; check if GUI checkbox not checked 
		{
			if (VarEnabled = 1) ; check if Record Asset No is checked
			{
				ControlClick, %AppForm%.BUTTON.%AppNumber%1, WE3 AMS Lite
				Sleep, 100
			}
			ControlFocus, %AppForm%.EDIT.%AppNumber%3, WE3 AMS Lite
		}
		sleep 50
		SendInput {Enter}
		sleep 50
		BlockInput Off
	} 
	Goto, Start

}
else
{
	;Run, C:\Program Files (x86)\Green Oak Solutions\We3 Off-Site Asset Manager\WE3OffSiteAssetManager.exe
	;WinActivate, ahk_exe WE3OffSiteAssetManager.exe	
	;WinWaitActive, WE3 AMS Lite
	MsgBox,, Exception, Please launch We3 OffSite Asset Manager
	return
}

FontSizeMenuButton:
Gui, FontGui:New, AlwaysOnTop, Font
Gui, Font, s%FontSize%
Gui, Add, Text, Center, Enter Font Size:
Gui, Add, ComboBox, vFontSize, 8|10|12|14|16|18|20|28|64|128|256
GuiControl, ChooseString, FontSize,%FontSize%
Gui, Add, Button, Default, &OK
Gui, Show, AutoSize Center
return

SerialGuiButtonOK:
AssetGuiButtonOK:
Gui, Submit
Return

FontGuiButtonOK:
Gui, Submit
Gui, Destroy
Goto, Start


SerialGuiGuiClose:
AssetGuiGuiClose:
SerialGuiGuiEscape:
AssetGuiGuiEscape:
SerialGuiButtonCancel: 
AssetGuiButtonCancel:
Gui, Destroy
Reload

}