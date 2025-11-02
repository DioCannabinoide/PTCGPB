#Include %A_ScriptDir%\Logging.ahk

#SingleInstance, force
CoordMode, Mouse, Screen
SetTitleMatchMode, 3

if not A_IsAdmin
{
    ; Relaunch script with admin rights
    Run *RunAs "%A_ScriptFullPath%"
    ExitApp
}

settingsPath := A_ScriptDir "\..\..\Settings.ini"

IniRead, instanceLaunchDelay, %settingsPath%, UserSettings, instanceLaunchDelay, 5
IniRead, waitAfterBulkLaunch, %settingsPath%, UserSettings, waitAfterBulkLaunch, 40000
IniRead, Instances, %settingsPath%, UserSettings, Instances, 1
IniRead, folderPath, %settingsPath%, UserSettings, folderPath, C:\Program Files\Netease
; New keep-awake settings (optional toggle in Settings.ini under [UserSettings])
IniRead, KeepAwakeEnabled, %settingsPath%, UserSettings, KeepAwakeEnabled, 1
IniRead, KeepAwakeIntervalSeconds, %settingsPath%, UserSettings, KeepAwakeIntervalSeconds, 45
IniRead, KeepAwakeIncludeDisplay, %settingsPath%, UserSettings, KeepAwakeIncludeDisplay, 0
mumuFolder = %folderPath%\MuMuPlayerGlobal-12.0
if !FileExist(mumuFolder)
    mumuFolder = %folderPath%\MuMu Player 12
if !FileExist(mumuFolder){
    MsgBox, 16, , Can't Find MuMu, try old MuMu installer in Discord #announcements, otherwise double check your folder path setting!`nDefault path is C:\Program Files\Netease
    ExitApp
}

; --- Keep-awake feature: periodically tell Windows to stay awake using SetThreadExecutionState ---
; Uses ES_SYSTEM_REQUIRED primarily; optionally also ES_DISPLAY_REQUIRED if enabled via INI.
; Keeps actions minimal (no simulated mouse/keyboard) and logs ticks. Compatible with S3 and attempts to
; detect Modern Standby (S0) via `powercfg /a` output; falls back to ES_SYSTEM_REQUIRED if S0 is found.

; Constants for SetThreadExecutionState
ES_CONTINUOUS         := 0x80000000
ES_SYSTEM_REQUIRED    := 0x00000001
ES_DISPLAY_REQUIRED   := 0x00000002
ES_AWAYMODE_REQUIRED  := 0x00000040

; Normalize interval and start timer if enabled
if (KeepAwakeEnabled != "0") {
    interval := KeepAwakeIntervalSeconds + 0
    if (interval < 30)
        interval := 30
    if (interval > 60)
        interval := 60
    ms := interval * 1000

    if (DetectModernStandby()) {
        LogToFile("KeepAwake: Modern Standby (S0) detected; using ES_SYSTEM_REQUIRED fallback", "Monitor.txt")
    } else {
        LogToFile("KeepAwake: No Modern Standby detected; using ES_SYSTEM_REQUIRED", "Monitor.txt")
    }

    ; Start timer - KeepAwakeTick will call SetThreadExecutionState periodically
    SetTimer, KeepAwakeTick, % ms
    LogToFile("KeepAwake: Enabled. Interval (s): " . interval . ", IncludeDisplay: " . KeepAwakeIncludeDisplay, "Monitor.txt")
} else {
    LogToFile("KeepAwake: Disabled via INI.", "Monitor.txt")
}

; --- Helper: detect Modern Standby by calling `powercfg /a` and scanning output ---
DetectModernStandby()
{
    tmpFile := A_Temp "\\ahk_powercfg_out.txt"
    RunWait, %ComSpec% " /c powercfg /a > """ tmpFile """ 2>&1",, Hide
    FileRead, out, %tmpFile%
    FileDelete, %tmpFile%
    if (ErrorLevel)
        return 0

    ; Look for common Modern Standby identifiers
    if InStr(out, "Standby (S0 Low Power Idle)") || InStr(out, "S0 Low Power Idle")
        return 1

    return 0
}

; Timer callback: call SetThreadExecutionState with minimal flags to keep system alive
KeepAwakeTick()
{
    global ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED, ES_AWAYMODE_REQUIRED, KeepAwakeIncludeDisplay

    flags := ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    if (KeepAwakeIncludeDisplay != "0")
        flags |= ES_DISPLAY_REQUIRED

    ; Prefer SetThreadExecutionState (DllCall) — avoids simulated input when possible
    result := DllCall("SetThreadExecutionState", "UInt", flags, "UInt")

    ; Log result for diagnostics (keeps messages lightweight)
    LogToFile("KeepAwakeTick: Called SetThreadExecutionState with flags=0x" . Format("{:X}", flags) . ", result=0x" . Format("{:X}", result), "Monitor.txt")

    ; If SetThreadExecutionState returned 0 (failure), perform a very lightweight fallback:
    ; a single, zero-distance mouse move via SendMessage to avoid visible cursor movement.
    if (result == 0) {
        ; Try a minimal mouse_event fallback that won't move the cursor visually (mouse event with no movement but a synthetic move may be ignored); instead send a harmless VK_SHIFT key down/up with PostMessage to workspace thread.
        ; We keep this commented out by default to avoid interfering. Uncomment if you need an input fallback.
        ; Send, {Shift down}{Shift up}
        LogToFile("KeepAwakeTick: SetThreadExecutionState failed; fallback not active by default.", "Monitor.txt")
    }
}

; --- End keep-awake additions ---

Loop {
    ; Loop through each instance, check if it's started, and start it if it's not
    launched := 0
    
    nowEpoch := A_NowUTC
    EnvSub, nowEpoch, 1970, seconds
    
    Loop %Instances% {
        instanceNum := Format("{:u}", A_Index)
        
        IniRead, LastEndEpoch, %A_ScriptDir%\..\%instanceNum%.ini, Metrics, LastEndEpoch, 0
        secondsSinceLastEnd := nowEpoch - LastEndEpoch
        if(LastEndEpoch > 0 && secondsSinceLastEnd > (11 * 60))
        {
            ; msgbox, Killing Instance %instanceNum%! Last Run Completed %secondsSinceLastEnd% Seconds Ago
            msg := "Killing Instance " . instanceNum . "! Last Run Completed " . secondsSinceLastEnd . " Seconds Ago"
            LogToFile(msg, "Monitor.txt")
            
            scriptName := instanceNum . ".ahk"
            
            killedAHK := killAHK(scriptName)
            killedInstance := killInstance(instanceNum)
            Sleep, 3000
            
            cntAHK := checkAHK(scriptName)
            pID := checkInstance(instanceNum)
            if not pID && not cntAHK {
                ; Change the last end date to now so that we don't keep trying to restart this beast
                IniWrite, %nowEpoch%, %A_ScriptDir%\..\%instanceNum%.ini, Metrics, LastEndEpoch
                
                launchInstance(instanceNum)
                
                sleepTime := instanceLaunchDelay * 1000
                Sleep, % sleepTime
                launched := launched + 1
                
                Sleep, %waitAfterBulkLaunch%
                
                ;Command := "Scripts\" . scriptName
                ;Run, %Command%
                scriptPath := A_ScriptDir "\.." "\" scriptName
                Run, "%A_AhkPath%" /restart "%scriptPath%"
            }
        }
    }
    
    ; Check for dead instances every 30 seconds
    Sleep, 30000
}

killAHK(scriptName := "")
{
    killed := 0
    
    if(scriptName != "") {
        DetectHiddenWindows, On
        WinGet, IDList, List, ahk_class AutoHotkey
        Loop %IDList%
        {
            ID:=IDList%A_Index%
            WinGetTitle, ATitle, ahk_id %ID%
            if InStr(ATitle, "\" . scriptName) {
                ; MsgBox, Killing: %ATitle%
                WinKill, ahk_id %ID% ;kill
                ; WinClose, %fullScriptPath% ahk_class AutoHotkey
                killed := killed + 1
            }
        }
    }
    
    return killed
}

checkAHK(scriptName := "")
{
    cnt := 0
    
    if(scriptName != "") {
        DetectHiddenWindows, On
        WinGet, IDList, List, ahk_class AutoHotkey
        Loop %IDList%
        {
            ID:=IDList%A_Index%
            WinGetTitle, ATitle, ahk_id %ID%
            if InStr(ATitle, "\" . scriptName) {
                cnt := cnt + 1
            }
        }
    }
    
    return cnt
}

killInstance(instanceNum := "")
{
    killed := 0
    
    pID := checkInstance(instanceNum)
    if pID {
        Process, Close, %pID%
        killed := killed + 1
    }
    
    return killed
}

checkInstance(instanceNum := "")
{
    ret := WinExist(instanceNum)
    if(ret)
    {
        WinGet, temp_pid, PID, ahk_id %ret%
        return temp_pid
    }
    
    return ""
}

launchInstance(instanceNum := "")
{
    global mumuFolder
    
    if(instanceNum != "") {
        mumuNum := getMumuInstanceNumFromPlayerName(instanceNum)
        if(mumuNum != "") {
            ; Run, %mumuFolder%\shell\MuMuPlayer.exe -v %mumuNum%
            Run_(mumuFolder . "\shell\MuMuPlayer.exe", "-v " . mumuNum)
        }
    }
}

getMumuInstanceNumFromPlayerName(scriptName := "") {
    global mumuFolder
    
    if(scriptName == "") {
        return ""
    }
    
    ; Loop through all directories in the base folder
    Loop, Files, %mumuFolder%\vms\*, D ; D flag to include directories only
    {
        folder := A_LoopFileFullPath
        configFolder := folder "\configs" ; The config folder inside each directory
        
        ; Check if config folder exists
        IfExist, %configFolder%
        {
            ; Define paths to vm_config.json and extra_config.json
            extraConfigFile := configFolder "\extra_config.json"
            
            ; Check if extra_config.json exists and read playerName
            IfExist, %extraConfigFile%
            {
                FileRead, extraConfigContent, %extraConfigFile%
                ; Parse the JSON for playerName
                RegExMatch(extraConfigContent, """playerName"":\s*""(.*?)""", playerName)
                if(playerName1 == scriptName) {
                    RegExMatch(A_LoopFileFullPath, "[^-]+$", mumuNum)
                    return mumuNum
                }
            }
        }
    }
}

; Temporary function to avoid an error in Logging.ahk
ReadFile(filename) {
    return false
}

; Function to run as a NON-adminstrator, since MuMu has issues if run as Administrator
; See: https://www.reddit.com/r/AutoHotkey/comments/bfd6o1/how_to_run_without_administrator_privileges/
/*
  ShellRun by Lexikos
    requires: AutoHotkey v1.1
    license: http://creativecommons.org/publicdomain/zero/1.0/
  
  Credit for explaining this method goes to BrandonLive:
  http://brandonlive.com/2008/04/27/getting-the-shell-to-run-an-application-for-you-part-2-how/
  
  Shell.ShellExecute(File [, Arguments, Directory, Operation, Show])
  http://msdn.microsoft.com/en-us/library/windows/desktop/gg537745
*/
Run_(target, args:="", workdir:="") {
    try
    ShellRun(target, args, workdir)
    catch e
        Run % args="" ? target : target " " args, % workdir
}
ShellRun(prms*)
{
    shellWindows := ComObjCreate("Shell.Application").Windows
    VarSetCapacity(_hwnd, 4, 0)
    desktop := shellWindows.FindWindowSW(0, "", 8, ComObj(0x4003, &_hwnd), 1)
    
    ; Retrieve top-level browser object.
    if ptlb := ComObjQuery(desktop
        , "{4C96BE40-915C-11CF-99D3-00AA004AE837}" ; SID_STopLevelBrowser
        , "{000214E2-0000-0000-C000-000000000046}") ; IID_IShellBrowser
    {
        ; IShellBrowser.QueryActiveShellView -> IShellView
        if DllCall(NumGet(NumGet(ptlb+0)+15*A_PtrSize), "ptr", ptlb, "ptr*", psv:=0) = 0
        {
            ; Define IID_IDispatch.
            VarSetCapacity(IID_IDispatch, 16)
            NumPut(0x46000000000000C0, NumPut(0x20400, IID_IDispatch, "int64"), "int64")
            
            ; IShellView.GetItemObject -> IDispatch (object which implements IShellFolderViewDual)
            DllCall(NumGet(NumGet(psv+0)+15*A_PtrSize), "ptr", psv
                , "uint", 0, "ptr", &IID_IDispatch, "ptr*", pdisp:=0)
            
            ; Get Shell object.
            shell := ComObj(9,pdisp,1).Application
            
            ; IShellDispatch2.ShellExecute
            shell.ShellExecute(prms*)
            
            ObjRelease(psv)
        }
        ObjRelease(ptlb)
    }
}

~+F7::ExitApp
