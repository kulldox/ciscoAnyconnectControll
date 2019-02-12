;============================== Start Auto-Execution Section ==============================
; Always run as admin
if not A_IsAdmin
{
   Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
   ExitApp
}
;============================== Main Script ==============================
; https://autohotkey.com/docs/Variables.htm#ComSpec

DetectHiddenWindows, On

msgboxDelay := 2
; Keeps script permanently running
#Persistent

#NoEnv
; Ensures that there is only a single instance of this script running.
#SingleInstance, Force

; sets title matching to search for "containing" instead of "exact"
SetTitleMatchMode, 2


; GroupAdd, saveReload, %A_ScriptName%

; return

;============================== Save Reload / Quick Stop ==============================
; #IfWinActive, ahk_group saveReload

; Use Control+S to save your script and reload it at the same time.
; ~^s::
	; TrayTip, Reloading updated script, %A_ScriptName%
	; SetTimer, RemoveTrayTip, 1500
	; Sleep, 1750
	; Reload
; return

; Removes any popped up tray tips.
; RemoveTrayTip:
	; SetTimer, RemoveTrayTip, Off 
	; TrayTip 
; return 

; Hard exit that just closes the script
; ^Esc::
; ExitApp


; #IfWinActive
;============================== Main Script ==============================

; Your main code here.

; VPNNAME := "Cisco AnyConnect Secure Mobility Agent"
global iniFileFullPath := A_ScriptDir . "\anyconnectControll.ini"
IniRead, NRFAILEDRETRIES, %iniFileFullPath%, settings, NRFAILEDRETRIES

global VPNSERVICENAME := "vpnagent"
global NRFAILEDRETRIES := 50
global PAUSETIME := 1
global CHECKPROFILE := "yes"
global RESTARTSERVICEFLAG := 0

global VPNPROFILEBACKUP := "C:\bkp\anyconnect\asOfAug72017\AnyConnect-General-Client-mct-custom-Profile.xml"
global APPVIPPATH := "C:\Program Files (x86)\Symantec\VIP Access Client\"
global VPNPROFILE := "C:\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client\Profile\AnyConnect-General-Client-Profile.xml"
global APPPATH := "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\"

; Loading all the Settings
ini_getsectiontext(iniFileFullPath,"settings")

currentStatusWindow := ""
guiLogWindowId := ""

ini_getsectiontext(iniFile, theSection) {

	IniRead, INIall, %iniFile%, %theSection%
	MsgBox % INIall
	tmpArr := StrSplit(INIall, "`n")
	for index, element in tmpArr
	{
		If element contains = 
			{ 
				Field := StrSplit(element, =) 
				varName := Field[1]
				varValue := Field[2]
				; MsgBox % index "==> " element " -> " varName ": " varValue
				if ( varName != "" )
				{
					%varName% := "%varValue%"
				}
				else {
					log( A_ThisFunc . ":" . varName . " is empty" . varValue)
				}
				; log("Reading from INI: " . index . "==> " . element . " -> " . varName . ": " . varValue)
			}
	}
	return
}

ini_profiles(iniFile, theSection) {
	profileList := []
	IniRead, INIall, %iniFile%
	tmpArr := StrSplit(INIall, "`n")
	; MsgBox %	INIall
	for index, element in tmpArr
	{
		RegExMatch(element,"`ami)(" . theSection . ".*)", sectionText)
		if ( sectionText != "" ) {
			; MsgBox %	index element "-" sectionText
			profileList.Push(sectionText)
		}
	}
	; MsgBox % A_ThisFunc ":"  profileList.Length()
	return	profileList
}

global iniExtractedProfiles := ini_profiles(iniFileFullPath,"profile_")

Gui, +LastFound
initialYval := -29
if ( iniExtractedProfiles.Length() > 0 ) {
	for index, element in iniExtractedProfiles
	{
		vpnProfileName := SubStr(element, 9)
		yVal := initialYval + 32 * index

		IniRead, TITLENAME, %iniFileFullPath%, %element%, TITLENAME
		; MsgBox , , Debug Log, % element
		; MsgBox , , Debug Log, vConnect2VPN%index%#%element%
		Gui, Add, Text, x12 y%yVal% w170 h20 , % TITLENAME " (" index ")"
		Gui, Add, Button, vConnect2VPN%index%#%element% gExec x192 y%yVal% w110 h20, Connect
		Gui, Add, Button, x312 y%yVal% w110 h20 gDisconnect, Disconnect (Ctrl+0)
	}
} else {
	Gui, Add, Text, x12 y19 w170 , % "There are no profiles to load in " iniFileFullPath
}

Gui, Add, Edit, x12 y109 w409 h200 vStatusWindow,
Gui, Add, StatusBar,,
; Generated using SmartGUI Creator 4.0
Gui, Show, x194 y107 h330 w439, Cisco Anyconnect Manager

return

; TheSubs
Exec:
	; MsgBox % A_GuiControl A_ThisLabel
    StartFunction( A_GuiControl )
Return


log(appendText, stillLoading := 0) {
	tempValue := ""
	tempTimestamp := ""
	FormatTime, tempTimestamp,,M/d/yyyy H:mm:ss
	ControlGetText, tempValue, Edit1, Cisco Anyconnect Manager
	if ( stillLoading = 0 ) {
		currentStatusWindow := % tempValue . "`r`n" . tempTimestamp . " " . appendText
	} else {
		currentStatusWindow := % tempValue . "."
	}
	ControlSetText, Edit1, %currentStatusWindow%, Cisco Anyconnect Manager
	; auto-scroll the the end of the editBox
	ControlSend,Edit1,^{End}, Cisco Anyconnect Manager
}

; generic VPN Disconnect Routine
; for some reason I'm getting BSDs when I have this service running. Thus I'm always stopping it on Disconnect.
; and starting it before Connecting.
Disconnect:
	SB_SetText("Disconnecting from VPN.")
	log("Start: disconnecting")
	log("Start: net stop " . VPNSERVICENAME)
	RunWait, %comspec% /c "net stop %VPNSERVICENAME%",,Hide,
	Sleep, 1000  ; Wait X seconds for the service to stop (Just to be sure)
	log("End: net stop " . VPNSERVICENAME)

	log("Start: taskkill.exe /IM VIPUIManager.exe")
	RunWait, taskkill.exe /IM VIPUIManager.exe,,Hide,currentStatusWindow
	log("End: taskkill.exe /IM VIPUIManager.exe")
	log("Start: taskkill.exe /IM vpnui.exe")
	RunWait, taskkill.exe /IM vpnui.exe,,Hide,currentStatusWindow
	log("End: taskkill.exe /IM vpnui.exe")
	log("Finish: disconnecting")
	; MsgBox , , Cisco Anyconnect VPN Client connection, Disconnected from VPN, %msgboxDelay%
	SB_SetText("Disconnected from VPN.")
return

; ShrewSoft VPN Client Connect
STmolab1VPN:
	Run C:\cmd\appstarter\stmolabVPN.bat
	MsgBox , , ShrewSoft VPN Client connection, Connected to TMO NVQALab1 VPN, %msgboxDelay%
return

; ShrewSoft VPN Client Disconnect
STmolab1VPNDisconnect:
	Run taskkill.exe /IM ipsecc.exe
	MsgBox , , ShrewSoft VPN Client connection, Disconnected from TMO NVQALab1 VPN, %msgboxDelay%
return

; generic CiscoAnyConnect Service Restart
RestartService:
	if ( CHECKPROFILE = "yes" ) {
		; some VPNs will overwrite the profile
		; so, I'm restoring it from my backup
		log("Start: Restore the custom profile from " . VPNPROFILEBACKUP . " to " . VPNPROFILE )
		RunWait, %comspec% /c "xcopy /y %VPNPROFILEBACKUP% %VPNPROFILE%",, Hide
		log("End: Restore the custom profile from " . VPNPROFILEBACKUP )
	}

	if ( RESTARTSERVICEFLAG = 0 ) {
		log("Start: net stop " . VPNSERVICENAME )
		RunWait, %comspec% /c "net stop %VPNSERVICENAME%",, Hide
		log("Ens: net stop " . VPNSERVICENAME )
		log("Start: net start " . VPNSERVICENAME)
		RunWait, %comspec% /c "net start %VPNSERVICENAME%",, Hide
		log("End: net start " . VPNSERVICENAME )
		Sleep, 1000
	}
return

GuiClose:
ExitApp

; this is the function called from the GUI
; just a wrapper for the main Connection Function. this way, it easier to call it from GUI
Connect2VPN(l_pName){
	; MsgBox %	A_ThisFunc ": "  iniFileFullPath " => " l_pName
	if ( l_pName = "") {
		MsgBox % A_ThisFunc ": Empty/invalid VPN profile name given '" l_pName "'"
		return
	}
	IniRead, i_vpnUser, %iniFileFullPath%, %l_pName%, vpnUser
	IniRead, i_vpnPass, %iniFileFullPath%, %l_pName%, vpnPass
	IniRead, i_PROFILENAME, %iniFileFullPath%, %l_pName%, PROFILENAME
	IniRead, i_TITLENAME, %iniFileFullPath%, %l_pName%, TITLENAME
	IniRead, i_hasBanner, %iniFileFullPath%, %l_pName%, hasBanner
	IniRead, i_hasVIP, %iniFileFullPath%, %l_pName%, hasVIP
	CiscoAnyConnectAutoConnect(i_vpnUser, i_vpnPass, i_PROFILENAME, i_TITLENAME, i_hasBanner, i_hasVIP)
}

; this is the function that is actually connecting to the VPN
CiscoAnyConnectAutoConnect(i_vpnUser, i_vpnPass, i_vpnProfileNname, i_vpnWindowTitleNname, hasBanner = 0, hasVIP = 0){
	anyconnectWindowTitleName := "Cisco AnyConnect Secure Mobility Client"

	SB_SetText("Connecting to " . i_vpnProfileNname . "...")
	log("Start: connecting to " . i_vpnProfileNname . "")
	; Run C:\cmd\appstarter\anyconnectVPNswitcher.bat production
	GoSub, RestartService
	if ( hasVIP != 0 ) {
		log("Start: starting VIPUIManager.exe")
		Run, %comspec% /c "%APPVIPPATH%VIPUIManager.exe",,Hide,
		log("Ends: starting VIPUIManager.exe")
	}
	Sleep, 500
	log("Start: starting vpnui.exe")
	Run, %comspec% /c "%APPPATH%vpnui.exe",,Hide,
	log("End: starting vpnui.exe")
	; MsgBox , , Cisco Anyconnect VPN Client connection, Connected to TMO Production VPN, %msgboxDelay%
	Sleep, 500
	log("Start: starting '" . i_vpnProfileNname . "' Connection")
	; SetTimer, ProductionVPNConnect, 100

	; MsgBox, The value of vpnUser/vpnPass is %i_vpnUser%/%i_vpnPass%.
	; return
    ; ############ Step 1
    ; select VPN Profile from the ComboBox1
	WinWaitActive, ahk_exe vpnui.exe
	; WinWaitActive, %anyconnectWindowTitleName%
	SB_SetText("Select VPN Profile '" . i_vpnProfileNname . "'")
	Sleep, 500
	Control, ChooseString , %i_vpnProfileNname%, ComboBox1, %anyconnectWindowTitleName%
	ControlSetText, Edit1, %i_vpnProfileNname%, %anyconnectWindowTitleName%

    ; ############ Step 2
	; Enter VPN username and password
	Sleep, 500
	SetControlDelay -1
	WinWaitActive, %i_vpnWindowTitleNname%
	SB_SetText("Provide user/pass for '" . i_vpnUser . "'")
	; Use the next line only for DEBUG
	log("user/pass: '" . i_vpnUser . "'/'" . i_vpnPass . "'")
	ControlSetText, Edit1, %i_vpnUser%, %i_vpnWindowTitleNname%
	ControlSetText, Edit2, %i_vpnPass%, %i_vpnWindowTitleNname%
	Sleep, 1000
	SB_SetText("Click Connect button.")
	SetControlDelay -1
	ControlClick , Button1, %i_vpnWindowTitleNname%,,,, NA
    if ( hasVIP != 0 ) {
	    ; ############ Step 3
	    ; Enter the Symactec VIP token
		Sleep, 1000
		SB_SetText("Fill in the SymantecVIPAccess Token.")
		log("Fill in the SymantecVIPAccess Token.")
		WinWaitActive, %i_vpnProfileNname%
		; ControlGetText, VIPToken , Static6, ahk_exe VIPUIManager.exe
		ControlGetText, VIPToken , Static6, VIP Access
		ControlSetText, Edit1, %VIPToken%, %i_vpnWindowTitleNname%
		SetControlDelay -1
		ControlClick , Button1, %i_vpnWindowTitleNname%,,,, NA
		log("SymantecVIPAccess Token filled in.")
		SB_SetText("SymantecVIPAccess Token filled in.")
	}
    if ( hasBanner != 0 ) {
	    ; ############ Step 4
	    ; Click Accept button for the banner
		Sleep, 1000
	    log("'Accept' the Banner.")
	    SB_SetText("'Accept' the Banner.")
		WinWaitActive, Cisco AnyConnect
		SetControlDelay -1
		ControlClick , Button1, Accept,,,, NA
		ControlClick , Button1, Cisco AnyConnect,,,, NA
	    log("Accepted Banner.")
	    SB_SetText("Accepted Banner.")
	}

	Sleep, 500
	WinWaitActive, %anyconnectWindowTitleName%
	successfullyConnected := "no"
	tempLoginStatus := ""
	Loop {
		SB_SetText("Check the LoginStatus.")
		; SB_SetText("Check the LoginStatus. Attempt: " . A_Index . "/" . NRFAILEDRETRIES)
		; log("Check the LoginStatus. Attempt: " . A_Index . "/" . NRFAILEDRETRIES)
		ControlGetText, LoginStatus, Static2, %anyconnectWindowTitleName%
		if ( tempLoginStatus !=  LoginStatus) {
			log("get LoginStatus " . LoginStatus)
		} else {
			log("get LoginStatus " . LoginStatus, 1)
		}
	    if (LoginStatus = "Login failed.")
	    {
	    	MsgBox , , VPN connection Status, ERROR: CiscoVPN returned "%LoginStatus%". Failed to connect to %i_vpnWindowTitleNname%!!!,
	    	SB_SetText("ERROR: CiscoVPN returned '" . LoginStatus . "'. Failed to connect to " . i_vpnWindowTitleNname . "!!!")
	    	break
	    }
	    if ( LoginStatus = "Connected to " . i_vpnProfileNname . "." ) {
	    	successfullyConnected := "yes"
	    	SB_SetText("Connected to " . i_vpnProfileNname . ".")
	    	break
		}
		Sleep, 1000
		tempLoginStatus := LoginStatus
		if ( NRFAILEDRETRIES <= A_Index ) {
			log("Reached number of MAX Retries " . NRFAILEDRETRIES . ". Exit.")
			SB_SetText("Reached number of MAX Retries " . NRFAILEDRETRIES . ". Exit.")
			break
		}
	}
    if ( successfullyConnected = "yes" ) {
	    log("All done, VPN is ready for use.")
	    SB_SetText("All done, VPN is ready for use.")
    } else {
	    log("Failed to connect to VPN '" . i_vpnProfileNname . "'.")
	    SB_SetText("Failed to connect to VPN '" . i_vpnProfileNname . "'.")
    }

	log("End: starting " . i_vpnProfileNname . "Connect")
	log("Finish: connecting to " . i_vpnProfileNname . "")
	SB_SetText("Connected to " . i_vpnProfileNname . ".")
}

; Function starter
; a hack for the GUI limitation (unable to call a function on Button click)
StartFunction( CtrlVarName )
{
    global
    local Param0,Param1,Param2,Param3,Param4
    ; MsgBox %	A_ThisFunc ":"  CtrlVarName
    ; Extract function name and parameter
    If ( RegExMatch( CtrlVarName, "([^\#].*)#", function ) 
      && RegExMatch( CtrlVarName, "#([\w#]+)", reMatch ) )
        StringSplit,Param,reMatch1,#
    
    ; Execute
    functionName := SubStr(function, 1, StrLen(function)-2)
    log("End: triggering " . functionName . " with " . Param1 . " param")
    ; MsgBox %	A_ThisFunc ":"  CtrlVarName "-" functionName ":> " Param1
    If ( IsFunc(functionName) ) {
    	; MsgBox %functionName%( %Param1%, %Param2%, %Param3%, %Param4% )
        return %functionName%( Param1, Param2, Param3, Param4 )
    }
    Else
        return false
}


#IfWinActive, Cisco Anyconnect Manager
; Ctrl + 0
	~^0:: Gosub, Disconnect
	return

; Ctrl + Shift + 0
	~^+0:: Gosub, STmolab1VPNDisconnect
	return

/*
TODO: I actually need to find a way to dinamically assing key bindings for each loaded VPN profile
currently, this doesn't work
 */

; for index, element in iniExtractedProfiles
; {
; 	vpnProfileName := SubStr(element, 9)

; Ctrl + 1
	; ~^{%index%}:: Gosub, %vpnProfileName%
	; return
	
; }


#IfWinActive


;============================== Program 1 ==============================
; Evertything between here and the next #IfWinActive will ONLY work in someProgram.exe
; This is called being "context sensitive"
; #IfWinActive, ahk_exe someProgram.exe



; #IfWinActive
;============================== ini Section ==============================
; Do not remove /* or */ from this section. Only modify if you're
; storing values to this file that need to be permanantly saved.
/*
[SavedVariables]
Key=Value
*/
;============================== GroggyOtter ==============================
;============================== End Script ==============================
