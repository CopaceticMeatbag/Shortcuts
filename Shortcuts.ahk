#SingleInstance force
#NoEnv
#Include WinClipAPI.ahk												; WinClip is a set of functions that enable easy access to clipboard. 
#Include WinClip.ahk												; Important to note we are using an internal clipboard (commands prefaced by "i" e.g. WinClip.iSetHTML) not windows clipboard. Windows clipboard is too slow to respond
#Include GetNestedTag.ahk											; Used to find the Font Type and Size in use when pasting in replacement text

SetBatchLines, -1
ListLines, Off
SendMode Input  													; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  										; Ensures a consistent starting directory.
Menu, tray, add, View Hotkeys, LaunchGui							; Adds the "View Hotkeys" menulet to the tray icon
;#############Setup initial folder and files#####################

IfNotExist,%A_appdata%\Shortcuts									; Checks for the existence of AppData\Shortcuts folder, creates if not found
	FileCreateDir, %A_appdata%\Shortcuts
FileInstall, MasterList.txt, %A_appdata%\Shortcuts\List.txt, 0	; MasterList.txt in Netlogon is added to exe when script is compiled, then copied to Shortcuts\list.txt - not overwriting if file exists already

GoSub,RefreshDataMaster												; Call to RefreshDataMaster subroutine, returns on complete
GoSub,RefreshData													; Call to RefreshData subroutine, returns on complete

;################Hotkey - press Alt 3 to activate################

!3::																;Hotkey function, !=Alt, for a full list of  symbols go to Mouse and Keyboard > Hotkeys and Hotstrings folder in AutoHotKey Help
WinClip.Clear()
Send, {shift Down}^{Left}{shift Up}^c								; Selects the word directly to the left of the cursor
Sleep, 100															; Wait to give windows time to copy to clipboard
vClip=%clipboard%													; Set vClip variable to contents of clipboard (will be hotkey e.g "ps1")
vTest := vArray[vClip]												; Sets vTest variable to item contained for dictionary key vClip in dictionary vArray - blank if key does not exist 
if (vTest)															; Tests if vTest is blank (i.e hotkey doesn't exist)
{
	wc := new WinClip												; Define a new instance of WinClip
	vFontType := GetFont()
	SubPat := 														; SubPat will be used in regex expression to extract a sub pattern
	vUrlArray := Object()											; Define new dictionary, will be used to regex a string in the following format: "Text with possible *url=http://test.com* inside"
	vUrlArrayCount=0
	FoundPos=1
	vUrlCheck = 0
	While FoundPos := RegExMatch(vTest,"(?<=\*).*?(?=\*)", SubPat)		; Loops through text stored in vTest variable until it encounters a *, saves all content until next *
	{																	; The .*? in the middle means "match all data non greedy". Omitting the ? will result in matching all data until the last *, instead of just the second * encountered.
		vUrlCheck=1														; FoundPos will be the character count inside the searched string where the first match was found. e.g for "abc123abc" FoundPos for "123" will be 4
		vUrlArrayCount += 1												; Output matched data to SubPat variable
		StringLeft,vLeft,vTest,FoundPos-2								; Get text from the left of the entire vTest variable, store in vLeft variable, start from FoundPos-2 (one back so we're behind first detected char, one to account for *)
		vUrlArray[vUrlArrayCount] := vLeft								; vLeft will be all text before the first detected URL. Save in vUrlArray dictionary with key being count of URLs found
		StringTrimLeft,vTest,vTest,FoundPos-1							; Deletes all characters from variable vTest to the left of character # FoundPos (minus 1 to account for first character in URL string)
		
		vUrlArrayCount += 1												; Add 1 to count of URLs found (1 is first string before URL, 2 is the URL, 3 will be next string found before next URL, 4 will be next URL, etc)
		vUrlArray[vUrlArrayCount] := SubPat								; Save matched string (variable SubPat) which is URL
		StringTrimLeft,vTest,vTest,StrLen(SubPat)+1						; Deletes all characters from left of variable vTest starting from length of matched URL string + 1 to account for *
	}
	if (vUrlCheck = 1)													; If URL was found
	{
		vUrlArrayCount += 1											
		vUrlArray[vUrlArrayCount] := vTest								; Add final entry to vUrlArray list, in case there was further text after the last detected URL
		vHTMLString = %vFontType%										;Format initial condition to match user's font type and size
		vRTFString = 
		Loop %vURLArrayCount%											; Loop to the count of vUrlArrayCount variable
		{
			if a_index in 1,3,5,7,9,11,13,15,17,19,21					; A_Index holds current loop data, if odd will be text, if even will be URL string.
			{
				vHTMLString .=  vURLArray[a_index]						; .= is the same as "vHTMLString := vHTMLString . vURLArray[A_Index]"
				vRTFString .=  vURLArray[a_index]
			}
			else														; Must be URL link if a_index isn't in the odd list
			{
				vTemp :=  vURLArray[a_index]							; := means assign vTemp the result of an expression (in this case, ask dictionary vUrlArray for data under key a_index)
				StringSplit,vUrl,vTemp,=								; Split the string into two parts at delimiter "=", access by calling vUrl1 or vUrl2 respectively
				if vUrl3												; If there's more than two entries in "vUrl", there must be an "=" as part of the url link
				{
					Loop, %vUrl0%										; Loop to the count of the number of entries in vUrl
					{
						if a_index in 1,2								; If the current count is 1 or 2, skip (because those entries will always be correct)
							continue
						vUrl2 .= "=" . vUrl%a_index%					; If higher than 2, append to the end of vUrl2, with an "=" appended first (as it's the delimiter char it's removed, so we re-add)
					}
				}
				vTempHTMLString=<a href=%vUrl2%>%vUrl1%</a>				; Format the URL data and save in string
				vHTMLString .= vTempHTMLString							; Append to vHTMLString
				vRTFString .=  vUrl1
			}
		}
		vHTMLString .= "</span>"
		wc.iClear()														; Clear the internal WinClip clipboard. Using internal because windows clipboard is too slow to respond to changes
		wc.iSetText(vRTFString)											; Set Plain Text version of hotkey text, URL data excluded.
		wc.iSetHTML(vHTMLString)										; Set HTML version, url included and correctly formatted
		IfWinActive, ahk_class rctrl_renwnd32							; If current window is Outlook (window class is "rctrl_renwnd32" and can be checked by using WinGetClass - see AHK help for example)
			Send, ^{Del}												; Just pasting data doesn't work due to Outlook's "helpful" autoformatting - so send "Ctrl + delete" to delete currently selected text first
		Sleep,50														; Wait 50ms to give windows time
		wc.iPaste()														; Paste from the internal clipboard. The appropriate data format (HTML or Plain) will automatically be chosen by windows
	}
	else
		Send % vArray[vClip]											; If no URLs were detected in hotkey text, just use "Send" command to send data as keystrokes
}
return
;####################Subroutines#############################
RefreshData:															; This subroutine reads the local txt file in appdata into variable vList, and parses the data into the vArray dictionary
FileRead,vList,%A_appdata%\Shortcuts\List.txt						; The vArray dictionary stores the text (in variable vSplit2) under a key (contained in variable vSplit1)
Sort,vList
vArray := Object()
vArrayCount = 0
Loop,Parse,vList,`n,`r
{
	vArrayCount += 1
	StringSplit,vSplit,A_LoopField,~
	vArray[vSplit1] := vSplit2
}
FileDelete, %A_appdata%\Shortcuts\List.txt							; Remove the old text file
FileAppend, %vList%, %A_appdata%\Shortcuts\List.txt					; Write the contents of variable vList to the text file (created automatically when no file exists)
return

RefreshDataMaster:														; This subroutine is run once at launch, and reads all entries from MasterList.txt on Netlogon into variable vListNew
FileRead,vListNew,MasterList.txt										; The data is then combined with items in local list.txt then saved (overwriting) in local list.txt
vMasterArray := Object()
Loop,parse,vListNew,`n,`r												; It is parsed in the same format as RefreshData
{
	if !A_LoopField
		continue
	StringSplit,vSplitMaster,A_LoopField,~
	vMasterArray[vSplitMaster1] := vSplitMaster2
}
FileRead,vListOld,%A_appdata%\Shortcuts\List.txt						; The local txt file is then read into variable vListOld
Loop, parse, vListOld,`n,`r												; Parse variable vListOld using "`n,`r" as delimiters
{
	if !A_LoopField														; If the current line is empty, go to next loop iteration
		continue
	StringSplit,vSplitMaster,A_LoopField,~								; Otherwise, split into two strings contained in variable vSplitMaster, delimiter is "~" symbol. vSplitMaster1 will contain first half, vSplitMaster2 the other half
	vMasterTest := vMasterArray[vSplitMaster1]							; Set variable vMasterTest to contents of vMasterArray dictionary as pointed to by key contained in vSplitMaster1
	if !vMasterTest														; If the key doesn't exist it means the hotkey is unique and can be added to vListNew
	{
		vMasterTempString = `r`n%A_LoopField%							; Append each entry with windows newline sequence `r`n
		vListNew .= vMasterTempString
	}
	else
	{
		if vListNew not contains %A_LoopField%							; Check the entire string (e.g. "ps1~test text and *url=http://test.com*") against vListNew. This is to stop entries which are *exactly* the same being marked as duplicate
		{																; i.e entries that have been copied accross from MasterList previously
			vMasterTempString = `r`n%vSplitMaster1%1~%vSplitMaster2%	; Otherwise, rename the hotkey "hotkeyname1"
			vListNew .= vMasterTempString								; And append to vListNew
		}
	}
}
Sort,vListNew															; Sort alphabetically
FileDelete, %A_appdata%\Shortcuts\List.txt							; Delete and copy new data from vListNew into local list.txt, this now contains all user-added combos as well as all entries in MasterList.txt
FileAppend, %vListNew%, %A_appdata%\Shortcuts\List.txt
return

ListView1:																; Unused g-label is called when certain actions are taken in GUI on the ListView (e.g doubleclicking a row).
return

ButtonAddNewCombo:														; This subroutine is called when user clicks "Add new combo" in GUI, it adds a new section in ListView but does not save the data
vUserIn1=																; User must click "Save & Close" button to save data
vuserIn2=
Loop
{
	InputBox,vUserIn1,Add New Combo,Enter new Keyword (e.g. t1),					; InputBox will ask the user to enter hotkey
	if (ErrorLevel || !vUserIn1)													; If user clicks cancel, close, or anything except for "OK", ErrorLevel is set to 1
		break
	If (vArray[vUserIn1])															; Ask vArray for data under key vUserIn1 (the text the user just entered to be new hotkey). If it returns data, we can't use that hotkey
		MsgBox, %vUserIn1% already assigned, please choose a different hotkey
	else
	{
		InputBox,vUserIn2,Add New Combo,Enter Text (e.g. This text replaces t1),	; Ask the user to enter text to action with hotkey
		if ErrorLevel
			break
		If (!vUserIn2)																; If blank, loop exits with no new hotkey created
			break
		Else
			LV_Add("",vUserIn1,vUserIn2)								; Otherwise, LV_Add is a built in function to add a new row in ListView, vUserIn1 is the hotkey, vUserIn2 is the text
			vArray[vUserIn1] := vUserIn2								; Make sure we update vArray to include new hotkey, otherwise user will be able to add duplicate hotkey before saving
			break
	}
}
return

ButtonSaveClose:														; This subroutine gets all data from ListView and appends it to variable "vNew", then writes contents of vNew to local list.txt
vNew = 
Loop % LV_GetCount()													; Loops through the count of number of rows in ListView (LV_GetCount() returns number of rows, can do columns as well)
{
LV_GetText(vRowText1,A_Index,1)
LV_GetText(vRowText2,A_Index,2)
vTotal = %vRowText1%~%vRowText2%`n										; Concatenate the hotkey and text with a return character at the end
vNew .= vTotal															; Append to variable "vNew"
}
FileDelete, %A_appdata%\Shortcuts\List.txt
FileAppend, %vNew%, %A_appdata%\Shortcuts\List.txt					; Write vNew to local txt file
GoSub, RefreshData														; Updates vArray dictionary with new entries by running "RefreshData" subroutine
Gui, Destroy															; Destroys the GUI ("Gui,Cancel" only hides the window)
return	

ButtonDeleteSelected:													; This subroutine deletes the selected row in ListView
vRowNum = 0
vRowNum := LV_GetNext(vRowNum)											; LV_GetNext searches for the next highlighted or selected row after the specified row (in this case we search from row 0)
if not vRowNum
{
	MsgBox, Please select a row first
}
else
	LV_Delete(vRowNum)
return

;####################Functions#############################
GetFont()
{
vDataGrab := WinClip.GetHTML()
vFontFamily := GetNestedTag(vDataGrab,"<span")
StringReplace, vFontFamily, vFontFamily, `r`n, , All
vFound := RegExMatch(vFontFamily,".*?>", vSpanCode)
return %vSpanCode%
}

;####################GUI Section#############################
LaunchGui:																					; This subroutine creates the GUI
Gui, Add, Text, x15 y13 w150 h20 , Action Hotkey =  Alt + 3			
Gui, Add, Button, x150 y10 w100 h20 , Add New Combo											; Clicking triggers "ButtonAddNewCombo"
Gui, Add, Button, x250 y10 w100 h20 , Delete Selected										; Clicking triggers "ButtonDeleteSelected"
;Gui, Add, Button, x382 y10 w100 h20 , Refresh												; Clicking triggers "ButtonRefresh"
Gui, Add, Button, x482 y10 w100 h20 , Save && Close											; Double Ampersand means one literal ampersand will be shown, when clicked triggers "ButtonSaveClose"
vListHeight := 80 + vArrayCount * 10														; Dynamically adjust size of ListView by 
If vListHeight > 1000
vListHeight = 1000
Gui, Add, ListView, x15 r%vArrayCount% h%vListHeight% w800 gListView1 vListView1, Keyword|Text		; Adds ListView to GUI, r%vArrayCount% creates rows to the count of vArray entries
for index, element in vArray																; Cycles through the entire vArray dictionary, variable "index" contains hotkey, variable "element" contains text
{
	LV_Add("",index,element)
}
Gui, Show,Autosize, Hotkeys												; Show the GUI window, Autosize to content upon creation
Return

GuiClose:																; This subroutine destroys the GUI and closes the window
MsgBox,4, Save Changes, Close without saving?							; Asks the user to click "Yes" or "No" to confirm exiting without clicking "Save & Close"
IfMsgBox,Yes
 Gui, Destroy
return