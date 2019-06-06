; This is a simple and pretty generic example of an AutoHotkey script to run a
; program when you press a keyboard shortcut. Add as many of these as you want
; to a .ahk file, and set that to be run at startup.

; See the Hotkeys reference [1] for details of the modifiers and keys available.

; [1]: http://www.autohotkey.com/docs/Hotkeys.htm


; Win+G - Open Gmail in Chrome
$#g::
    Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --app="https://mail.google.com/mail/"
    Return

; Win+Shift+Break - Edit this file
F7::
    if WinExist( ● AutoHotkey.ahk - Visual Studio Code)
    {
        WinActive("● AutoHotkey.ahk - Visual Studio Code")
        send ^s
        Reload
        WinActivate, ● AutoHotkey.ahk - Visual Studio Code
        return

    }
    if !WinExist( ● AutoHotkey.ahk - Visual Studio Code)
    {
        Run "C:\Program Files\Microsoft VS Code\code.exe" "C:\Users\ross.fiamingo\Documents\AutoHotkey.ahk"
        Return
    }


$#F7::
    Send, ^s ; To save a changed script
    Sleep, 300 ; give it time to save the script
    Reload
    Return
	
	; Win+Shift+Break - Open Nav 
F8::
    Run "C:\Program Files (x86)\Microsoft Dynamics NAV\80\RoleTailored Client\Microsoft.Dynamics.Nav.Client.exe" 
    Return
	

F3:: ;Open Jira case number
	InputBox, jobno, Enter Jira Number, Please Enter SHCA number
	run, chrome.exe http://jira.wencomine.com:8080/browse/SHCA-%jobno%
	return 
	
F10:: ; Print to pdf with proper naming scheme
	if WinActive("ahk_class WindowsForms10.Window.8.app.0.265601d_r9_ad1")
	{	
		ControlGetText,customer,WindowsForms10.EDIT.app.0.265601d_r9_ad121
		EndPos := InStr(customer, "-") -1
		custmr := SubStr(customer,1,EndPos)
		WinGetActiveTitle, heading
		serviceitemno := SubStr(heading, 23,9)
		FormatTime, CurrentDateTime,, dd-MM-yy
		Click, 723,681
		Click, 395, 10
		send, {LAlt}
		send, h
		send, p
		WinWaitActive,Edit - Service Quotation (Wenco)
		send,{Ctrl down}{Shift down}p{Ctrl up}{Shift up}
		send, {Right} 
		send, {Enter}
		WinWaitActive,Export File
		send,{raw}Repair Quote # %custmr%-%serviceitemno% %CurrentDateTime%
	}
	return	
	
$#S::

    ; Putty copies selected text to the clipboard so you don't need to copy it
    ; doing Ctrl-Insert throws away what you already have in the Clipboard.
    ; Can't use Ctrl-C in putty, because it sends that to your session as ^C
    WinGet, Active_ID, ID, A
    WinGet, Active_Process, ProcessName, ahk_id %Active_ID%
    if ( Active_Process ="putty.exe" )
    {

        host = %Clipboard% p

		
    }
    else
    {

        host_tmp := SelectedViaClipboard()
        host := ValidateHostname(host_tmp)

    }

    title_msg := "Please enter the hostname"
    prompt_msg := "Didn't get host from selected text:"
 
    if (!host)
    {
        InputBox, new_host, %title_msg%, %prompt_msg%,, 220, 111, , , , , %host_tmp%
        if (ErrorLevel)
        {
            return
        }
        else
        {
            host := ValidateHostname(new_host)
        }
    }

    if (host)
    {
        ; MsgBox, Doing ssh %host%
        Run "C:\Users\ross.fiamingo\Downloads\PuTTYPortable\PuTTYPortable\PuTTYPortable\PuTTYPortable.exe" "-ssh" "%host%"
        ; NB: You could add your username as "user@%host%"
    }
    else
    {
        MsgBox, %prompt_msg% [%host_tmp%]
    }
    return

;; utility functions

;; Get selected text without clobbering Clipboard
SelectedViaClipboard()
{
    old_clipboard = %ClipboardAll% ; save current Clipboard
    Clipboard := "" ; clears Clipboard
    Send, ^{Insert}
    selection = %Clipboard% ; save the content of the clipboard
    Clipboard = %old_clipboard% ; restore old content of the clipboard
    return selection

}

;; Capture a hostname out of some selected text
ValidateHostname(str)
{
    if (RegExMatch(str, "^([ \'\[@]*)([\w\-\.]+)([ \'\]:/]*)$", match))
        return match2
    return ""
}


F12:: ;Copy service item number to excel document and copy excel contents to clipboard
    if WinActive("ahk_class WindowsForms10.Window.8.app.0.265601d_r9_ad1")
    {
    ControlGetText,customer,WindowsForms10.EDIT.app.0.265601d_r9_ad121 

    fn := "C:\Users\ross.fiamingo\Desktop\Nav price templates.xlsx"
    oExcel := ComObjCreate("Excel.Application")
    oWorkbook := oExcel.Workbooks.Open(fn)
    oWorkbook.Sheets(1).Range("A2:A7").Value := customer
    oWorkbook.Save()
    oWorkbook.Sheets(1).Range("A2:F7").copy
    oExcel.Quit()

    }
    return

F9:: ; input ip address of unit and select folder to save mac address and license backups
    InputBox, IP, IP Address, IP of external Computer?
    run mac address.exe, C:\Licence\Mac Address Getter
    WinActivate C:\Licence\Mac Address Getter\mac address.exe
    WinWaitActive C:\Licence\Mac Address Getter\mac address.exe
    Send, %IP%
    SendInput, {Enter}

    sleep 1000
    FileSelectFolder, filelocation, Z:\Customer , 3
    
    FileMove, C:\Licence\requests\*.req, %filelocation%
    FileMove, C:\Licence\licences\*.lic, %filelocation%
    FileMove, C:\Licence\Mac Address Getter\*.txt, %filelocation%
    Run, %filelocation%
    return

 F4:: ;Search explorer in the customer folder from the service item number
     if WinActive("ahk_class WindowsForms10.Window.8.app.0.265601d_r9_ad1")
     {
     ControlGetText,sic,WindowsForms10.EDIT.app.0.265601d_r9_ad121
     EndPos := InStr(sic, "-") -1
 	 custmoo := SubStr(sic,1,EndPos)
     
     Iniread, Custdir,C:\Users\ross.fiamingo\Documents\customers.ini,Customernames,%custmoo%
     run, %custdir%
    
     }

      return
    