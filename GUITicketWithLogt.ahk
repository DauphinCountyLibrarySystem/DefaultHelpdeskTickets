/*
						***SPICEWORKS AUTO-TICKET APP***
	App for the automatic completion and submission of common helpdesk tickets.
	Author: Christopher Roth
	*NEW VERSION! developing internal tracking of ticket types and logs them to a file.
*/
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;==========VARIABLE DECLARATION==========
SplitPath, A_ScriptName, , , , ScriptBasename ; for Log fucntion
StringReplace, AppTitle, ScriptBasename, _, %A_SPACE%, All ; For Log function
aBranches := {1: "ESA", 2: "KL", 3: "MOM", 4: "MRL", 5: "AFL", 6: "JOH", 7: "EV", 8: "ND"} ;Array containg library branch locations
vCopierTally := 0
vEmailTally := 0
vEreaderTally := 0
vLoginTally := 0
vPrintTally := 0
vScannerTally := 0
vScarewareTally := 0
vSoftwareAssistTally := 0
vTimerExtendTally := 0
vWebAssistTally := 0
vWifiConnectTally := 0
vWifiPrintTally := 0

;==========INITIALIZING==========
InputBox,vUser, Log In, Please enter your full user name: ;Recieves HDA name for final ticket, allows for multiple users.
Log("Starting ticket tracking for user: "vUser)

;==========GUI GENERATION==========
Gui, New, , Common Helpdesk Tickets
;This section creates a toggle for Library locations.
Gui, Font, Bold s10
Gui, Add, GroupBox, r8, Select Branch:
Gui, Font, Norm
Gui, Add, Radio, altsubmit vBranch xp+10 yp+20, East Shore
Gui, Add, Radio, altsubmit, Kline Library
Gui, Add, Radio, altsubmit, Madeline Olewine
Gui, Add, Radio, altsubmit, McCormick Riverfront
Gui, Add, Radio, altsubmit, Alexander Family
Gui, Add, Radio, altsubmit, Johnson Memorial
Gui, Add, Radio, altsubmit, Elizabethville
Gui, Add, Radio, altsubmit, Northern Dauphin
;This section creates buttons for common tickets.
Gui, Font, Bold
Gui, Add, Text, ym, Select Ticket Type:
Gui, Font, Norm
Gui, Add, Button, gBprnt xp+10 yp+20 w100, Printing
Gui, Add, Button, gBlogn w100, E-Ware Login
Gui, Add, Button, gBwifi w100, Wi-Fi
Gui, Add, Button, gBwebs w100, Web Assist
Gui, Add, Button, gBscar w100, Scareware
Gui, Add, Button, gBtime w100, Extend Timer
Gui, Add, Button, gBsoft ym+20 w100, Software
Gui, Add, Button, gBcopy w100, Copier
Gui, Add, Button, gBmail w100, E-Mail
Gui, Add, Button, gBread w100, E-Reader
Gui, Add, Button, gBscan w100, Scanner
Gui, Add, Button, gBwprt w100, Wi-Fi Print
Gui, Show, AutoSize
Return

Fsend: ;Final subroutine, that takes in the variable strings, checks that toggles and windows are active, and submits the result to the new ticket form.
{
	Gui, Submit, NoHide
	if (Branch == 0) ;Check if branch location was properly toggled on.
	{	
		MsgBox, 48, No Library, Please select your library branch.
		Return
	}
	else
	{
		vMyLoc := aBranches[Branch]
		WinActivate Spiceworks
		Sleep 500
		IfWinActive Spiceworks ;Check if Spiceworks widow is open.
		{ 
			;Take inputs from the variables in the subrotines, and appends them to a final Send string that fills the web form, using tabs to skip fields.
			SendInput n ;Spiceworks shortcut for new ticket window.
			Sleep 1000
			SendInput %vSuma% {Tab} %vDesc% {Tab} %vUser% {Tab 7} %vCata% {Tab 3} %vMyLoc% {Tab 3} %vSubc% {Tab 3} On-Site {Tab 3} Patron {Tab 7}
			Return
		}
		else
		{	
			MsgBox, 48, Spiceworks Offline, Please open a new Spiceworks web interface window.
			Return
		}
	}
	Return
}
Return
Bprnt: ;Subroutine for Printing button.
	{	
		vPrintTally++
		Log("Printing ticket.")
		vSuma := "Release a Print"
		vDesc := "Showed a patron how to print a document and release it from the Print Release Station."
		vCata := "Training"
		vSubc := "LPTOne"
		Gosub, Fsend
		Return
	}
Blogn: ;Subroutine for Login button.
	{
		vLoginTally++
		Log("Envisionware login ticket.")
		vSuma := "Login to Envisionware"
		vDesc := "Helped a patron with logging into Envisionware."
		vCata := "Software"
		vSubc := "Envisionware"
		Gosub, Fsend
		Return
	}
Bwifi: ;Subroutine for Wi-Fi button.
	{
		vWifiConnectTally++
		Log("Wireless connections ticket.")
		vSuma := "Connect to Wi-Fi"
		vDesc := "Helped a patron bypass the certificate error and connect to public wi-fi."
		vCata := "Network"
		vSubc := "Troubleshooting"
		Gosub, Fsend
		Return
	}
Bwebs: ;Subroutine for website assistance button.
	{
		vWebAssistTally++
		Log("Website assistance ticket.")
		vSuma := "Website Assistance"
		vDesc := "Assisted a patron with navigating a web interface."
		vCata := "Training"
		vSubc := "Website Assistance"
		Gosub, Fsend
		Return
	}
Bsoft: ;Subroutine for Software button.
	{
		vSoftwareAssistTally++
		Log("Software assistance ticket.")
		vSuma := "Default Software"
		vDesc := "Showed a patron how to use some of the more advanced features of our default software."
		vCata := "Training"
		vSubc := "Software Assistance"
		Gosub, Fsend
		Return
	}
Bcopy: ;Subroutine for copier button.
	{
		vCopierTally++
		Log("Copier assistance ticket.")
		vSuma := "Copier Assistance"
		vDesc := "Helped a patron with copier functions."
		vCata := "Hardware"
		vSubc := "Copier"
		Gosub, Fsend
		Return
	}	
Bmail: ;Subroutine for e-mail button.
	{
		vEmailTally++
		Log("Email functions ticket.")
		vSuma := "E-Mail Assistance"
		vDesc := "Helped a patron with e-mail functions."
		vCata := "Training"
		vSubc := "Email"
		Gosub, Fsend
		Return
	}
Bread: ;Subroutine for e-reader button.
	{
		vEreaderTally++
		Log("e-Reader assisance ticket.")
		vSuma := "E-Reader Assistance"
		vDesc := "Helped a patron with quesions about e-books and their e-Reader device."
		vCata := "eReader"
		vSubc := "Software Assistance"
		Gosub, Fsend
		Return
	}
Bscar: ;Subroutine for Scareware button
	{
		vScarewareTally++
		Log("Scareware ticket.")
		vSuma := "Scareware"
		vDesc := "Helped clear a scareware prompt."
		vCata := "Software"
		vSubc := "Troubleshooting"
		Gosub, Fsend
		Return
	}
Bscan: ;Subroutine for Scanner functions
	{	
		vScannerTally++
		Log("Scanner assistance ticket.")
		vSuma := "Scan a Document"
		vDesc := "Showed a patron how to use the Copier as a scanner."
		vCata := "Hardware"
		vSubc := "Copier"
		Gosub, Fsend
		Return
	}
Btime: ;Subroutine for session timer button.
	{	
		vTimerExtendTally++
		Log("Time extention ticket.")
		vSuma := "Extended Patron Time"
		vDesc := "Helped a patron with questions about extending their session time"
		vCata := "Software"
		vSubc := "Envsionware"
		Gosub, Fsend
		Return
	}
Bwprt: ;Subroutine for print from anywhere.
	{
		vWifiPrintTally++
		Log("Print from Anywhere ticket.")
		vSuma := "Print From Anywhere"
		vDesc := "Showed a patron how to print from their wireless device"
		vCata := "Training"
		vSubc := "LPTOne"
		Gosub, Fsend
		Return
	}
Log(msg) ;Log Function
{
	global ScriptBasename, AppTitle
	FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%.%A_MSec%%A_Tab%%msg%`n, %ScriptBasename%.log
	Sleep 50 ; Hopefully gives the filesystem time to write the file before logging again
	Return
}
Total(a,b,c,d,e,f,g,h,i,j,k,l)
{
	Return a + b + c + d + e + f + g + h + i + j + k + l
}
GuiClose:
{
	vGrandTotal := Total(vCopierTally, vEmailTally, vEreaderTally, vLoginTally, vPrintTally, vScannerTally, vScarewareTally, vSoftwareAssistTally, vTimerExtendTally, vWebAssistTally, vWifiConnectTally, vWifiPrintTally)
}