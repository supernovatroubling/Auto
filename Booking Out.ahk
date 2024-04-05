#SingleInstance Ignore

FontSize = 28
VarProdCorrectYes = 1
VarProdCorrectNo = 0

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


^h::
Start:
if WinExist("AMS Processor Client ")
{
	SetTitleMatchMode, 1 ; Set title match to partial match
	WinGetClass, VarClass, AMS Processor Client
	AMSClassArray := StrSplit(VarClass, ".")
	AppForm := AMSClassArray[1]
	AppNumber := AMSClassArray[4]
	
	Loop
	{
		WinActivate, ahk_exe AMSProcessor.exe
		WinWaitActive, AMS Processor Client,,5
		ControlGetText, VarScanStatus, %AppForm%.STATIC.%AppNumber%6, AMS Processor Client ; Get Barcode Scanning status

		If WinActive("AMS Processor Client ")
		{
			Switch VarScanStatus
			{
				Case "Waiting For Scan", "ID Not Valid, Waiting For Scan":
					if WinExist("Serial Number")
					{
						WinGet, VarSerialGUIState, MinMax, Serial Number
						if (VarSerialGUIState = -1)
						{
							WinRestore, Serial Number
						}						
					}
					Else
					{
						Gui, SerialGui:New, AlwaysOnTop, Serial Number
						Menu, SerialMenuBar, Add
						Menu, SerialMenuBar, deleteAll

						Menu, Submenu_font, Add, &Size, FontSizeMenuButton
						Menu, SerialMenuBar, Add, &Font, :Submenu_font

						Gui, Menu, SerialMenuBar
						Gui, Font, s%FontSize%
						GUI, Add, Text, Center, Scan Barcode
						Gui, Add, Text, Center, Enter serial number:
						Gui, Add, Edit,r1 vVarSerial Uppercase -WantReturn
						Gui, Add, Button, Default, &OK
						Gui, Add, Button, x+m, &Cancel
						Gui, Show, AutoSize Center
						
					}
					WinWaitClose, Serial Number
					SanitizedSerial:= StrReplace(VarSerial, "  ")

					ControlFocus, %AppForm%.EDIT.%AppNumber%5, AMS Processor Client ; Switch to AMS
					BlockInput On
					ControlSetText, %AppForm%.EDIT.%AppNumber%5, %SanitizedSerial%, AMS Processor Client ; Enter Serial Number intotext box
					SetControlDelay -1
					ControlClick,  %AppForm%.BUTTON.%AppNumber%2,AMS Processor Client ,,,, NA ; Click Enter Button
					BlockInput Off
					Sleep, 200
				Case "Does this item contain a Hard Drive (HDD/SSD)":
					ControlGetText, VarProductName, %AppForm%.EDIT.%AppNumber%4, AMS Processor Client ; Get Product Name
					Gui, AttributesGUI:New, AlwaysOnTop, Attributes
					Menu, AttributesGUI, Add
					Menu, AttributesGUI, deleteAll

					Menu, Submenu_font, Add, &Size, FontSizeMenuButton
					Menu, MenuBar, Add, &Font, :Submenu_font

					Gui, Menu, MenuBar
					Gui, Font, s%FontSize%
					Gui, Add, Text, Center, Product Name: %VarProductName%
					Gui, Add, Text, Center, Enter Make:
					Gui, Add, Edit,r1 vVarMake x+m w420, %VarMake%
					Gui, Add, Text, Center xm, Enter Model:
					Gui, Add, Edit,r1 vVarModel x+m w420, %VarModel%
					Gui, Add, Text, Center xm, Enter HDD Action:
					Gui, Add, DropDownList, vVarHDDAction x+m Sort, Destroy|No HDD||Wipe
					GuiControl, ChooseString, VarHDDAction, ||%VarHDDAction% ; Sets DropDownList to previously entered option
					Gui, Add, Text, Center xm, Is the product name correct?
					Gui, Add, Radio, Checked%VarProdCorrectYes% vVarProdCorrect, Yes
					Gui, Add, Radio, Checked%VarProdCorrectNo%, No
					Gui, Add, Text, Center, Enter HDD Serial:
					Gui, Add, Edit,r2 Uppercase WantReturn -Wrap vVarHDDSerial x+m
					Gui, Add, Button, Default xm, &OK
					Gui, Add, Button, x+m, &Cancel
					Gui, Show, AutoSize Center
					WinWaitClose, Attributes

					BlockInput On
					Loop
					{
						ControlGetText, VarWorkflow, %AppForm%.STATIC.%AppNumber%8, AMS Processor Client ; Get workflow name
						Switch VarWorkflow
						{
							Case "WorkFlow": ; Does this unit contain a HDD menu
								SetControlDelay -1
								If (VarHDDAction = "No HDD")
									ControlClick, %AppForm%.BUTTON.%AppNumber%5,AMS Processor Client ,,,, NA ; Click No Button
								Else If (VarHDDAction = "Destroy" || "Wipe")
									ControlClick, %AppForm%.BUTTON.%AppNumber%4,AMS Processor Client ,,,, NA ; Click Yes Button								
								Gosub, WorkflowWait					
							Case "This screen is used to capture certificate information about hard drives wiped and belonging to this unit.": ; HDD Serial Menu
								ControlSetText, %AppForm%.EDIT.%AppNumber%6, %VarHDDSerial%, AMS Processor Client ; Enter HDD Serial Number into box
								ControlSetText, %AppForm%.EDIT.%AppNumber%5, %VarHDDSerial%, AMS Processor Client ; Enter HDD Serial Number into box
								Sleep, 500
								SetControlDelay -1
								ControlClick, %AppForm%.BUTTON.%AppNumber%2,AMS Processor Client ,,,, NA ; Click Next Button
								Gosub, WorkflowWait
							Case "Buttons": ; Is the product name correct menu
								Switch VarProdCorrect
								{
									Case "1":
										SetControlDelay -1
										ControlClick, %AppForm%.BUTTON.%AppNumber%5,AMS Processor Client ,,,, NA ; Click Yes Button
										VarProdCorrectYes = 1
										VarProdCorrectNo = 0
									Case "2":
										SetControlDelay -1
										ControlClick, %AppForm%.BUTTON.%AppNumber%4,AMS Processor Client ,,,, NA ; Click No Button
										VarProdCorrectYes = 0
										VarProdCorrectNo = 1
								}								
								Gosub, WorkflowWait								
							Case "New Model Number": ; Product change menu
								BlockInput Off
								Gosub, WorkflowWait
								BlockInput On
							Case "Use the available fields below to record extra information that may be required against this product", "Capture Attribute Data": ; Product Data menu
								SetControlDelay -1
								if ( VarMake != "") ; Check to see if user has entered text for make and proceed is they have
								{
									ControlSetText, %AppForm%.EDIT.%AppNumber%5, %VarMake%, AMS Processor Client ; Enter make into text box
									ControlSetText, %AppForm%.EDIT.%AppNumber%6, %VarModel%, AMS Processor Client ; Enter model into text box
									Control, ChooseString, %VarHDDAction%, %AppForm%.COMBOBOX.%AppNumber%1, AMS Processor Client ; Set HDD action
								}
								Else
								{
									MsgBox, Please fill out fields, then click OK
								}
								ControlClick, %AppForm%.BUTTON.%AppNumber%3,AMS Processor Client ,,,, NA ; Click Next button
								Gosub, WorkflowWait
							Case "Current Unit ID / Serial No": ; Check if there are no attributes to capture
								ControlGetText, VarNoAttrib, %AppForm%.STATIC.%AppNumber%5, AMS Processor Client ; Get capture attributes data
								if (VarNoAttrib = "No data fields exist. Click on NEXT to continue")
								{
									ControlClick, %AppForm%.BUTTON.%AppNumber%3,AMS Processor Client ,,,, NA ; Click Next button
								}
								Gosub, WorkflowWait
							Case "Add Unit To Inventory":
								SetControlDelay -1
								ControlClick, %AppForm%.BUTTON.%AppNumber%3, AMS Processor Client, Print Inventory Label ; Uncheck print label box
								ControlClick, %AppForm%.BUTTON.%AppNumber%5, AMS Processor Client ; Click Add button
								WinWait, Inventory Add,, 5 
								If WinExist("Inventory Add")
									WinActivate ; Use the window found by WinExist.
									Send {Space}
									Break
							Case "":
								MsgBox, 16, Timed Out, Timed out waiting for AMS Processor Client
								Reload
							Default:
								MsgBox, Unknown Workflow (%VarWorkflow%)
								Break
						}
						Sleep, 500
					}
					BlockInput Off
				Case "This unit is not to be processed at this workstation! Review the workflow from the details below and move accordingly. Contact your supervisor for assistance if necessary.":
					SetControlDelay -1
					ControlClick, %AppForm%.BUTTON.%AppNumber%3,AMS Processor Client ,,,, NA ; Click Reset Button
					Sleep, 200
				Case "Unit Found Ready To Process":
					Sleep, 500
				Default:
					MsgBox, Unknown scan status (%VarScanStatus%)
			}

		}
		Sleep, 200
	}
}
Return

WorkflowWait:
Loop
{
	ControlGetText, VarWorkflowDone, %AppForm%.STATIC.%AppNumber%8, AMS Processor Client ; Get new workflow name after finishing case
	Sleep, 200
} 
Until VarWorkflow != VarWorkflowDone	
Return

=::
Input, VarInput, T0.3 L1, {enter}
If (ErrorLevel = "EndKey:Enter")
{	
	If WinActive("Attributes")
	{
		ControlFocus, Edit3, Attributes
	}
}
Else if (ErrorLevel = Max)
{
	SendInput {text}=%VarInput%
}
Else
{
	
	SendInput {text}=%VarInput%
}
Return

>::
If WinActive("Attributes")
{
	SetControlDelay -1
	ControlClick,
	ControlClick, Button3, Attributes, OK ; Click ok button on attributes GUI
}
Else
{
	Send {Text}>
}
Return

<::
If WinActive("AMS Processor Client ")
{
	Send !r
	Goto, Start
}
Else
{
	Send {Text}<
}
Return

= & Enter::

Return

FontSizeMenuButton:
Gui, FontGui:New, AlwaysOnTop, Font
Gui, Font, s%FontSize%
Gui, Add, Text, Center, Enter Font Size:
Gui, Add, ComboBox, vFontSize, 8|10|12|14|16|18|20|28|64|128|256
GuiControl, ChooseString, FontSize,%FontSize%
Gui, Add, Button, Default, &OK
Gui, Show, AutoSize Center
Return

SerialGuiButtonOK:
AttributesGuiButtonOK:
Gui, Submit
Return

FontGuiButtonOK:
Gui, Submit
Gui, Destroy
Goto, Start


SerialGuiGuiClose:
SerialGuiGuiEscape:
SerialGuiButtonCancel:
AttributesGuiButtonCancel:
AttributesGuiGuiEscape:
AttributesGuiGuiClose:
Gui, Destroy
Reload

}