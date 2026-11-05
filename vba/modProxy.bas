Attribute VB_Name = "modProxy"
Option Explicit
' The proxy of the Jira site when the Config sheet leaves ProxyServer empty, found as the browser finds it: from the
' Windows settings of the user (Internet Options), in this order: the automatic detection (WPAD) and the proxy script
' (PAC, AutoConfigURL), then the fixed proxy, else direct. WinHTTP does not do it alone: without SetProxy it uses its
' own setting (netsh winhttp), most often "direct". modJira.JiraConfigFromSheet calls it once per run.
' The strings that WinHTTP allocates are copied, then freed with GlobalFree, as its documentation asks.
'
' The WinHTTP structures are byte buffers with the offsets of the C structures (a pointer is 8 bytes in 64-bit Office,
' 4 in 32-bit; a pointer after a 4-byte field starts at the next multiple of its size). A VBA Type did not match the C
' layout of WINHTTP_CURRENT_USER_IE_PROXY_CONFIG in 64-bit Excel: the call wrote past it and Excel crashed.

#If Win64 Then
Private Const PTR_SIZE As Long = 8
#Else
Private Const PTR_SIZE As Long = 4
#End If

Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "winhttp.dll" (ByVal config As LongPtr) As Long
Private Declare PtrSafe Function WinHttpOpen Lib "winhttp.dll" (ByVal agent As LongPtr, ByVal accessType As Long, _
    ByVal proxyName As LongPtr, ByVal proxyBypass As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpGetProxyForUrl Lib "winhttp.dll" (ByVal session As LongPtr, _
    ByVal url As LongPtr, ByVal options As LongPtr, ByVal info As LongPtr) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp.dll" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal target As LongPtr, ByVal source As LongPtr, _
    ByVal bytes As LongPtr)

Private Const AUTOPROXY_AUTO_DETECT As Long = 1
Private Const AUTOPROXY_CONFIG_URL As Long = 2
Private Const AUTO_DETECT_TYPE_DHCP As Long = 1
Private Const AUTO_DETECT_TYPE_DNS_A As Long = 2
Private Const ACCESS_TYPE_NO_PROXY As Long = 1

Public gProxyFake As Object         ' tests: the Windows settings to use (FakeProxySettings), instead of the real ones

Public Function FakeProxySettings(ByVal autoDetect As Boolean, ByVal autoConfigUrl As String, ByVal proxy As String, _
                                  ByVal bypass As String) As Object
    ' The four Windows settings, as WindowsProxySettings returns them (also the form the tests give to gProxyFake).
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("AutoDetect") = autoDetect: d("AutoConfigUrl") = autoConfigUrl: d("Proxy") = proxy: d("Bypass") = bypass
    Set FakeProxySettings = d
End Function

Public Function WindowsProxySettings() As Object
    ' The proxy settings of the user in Windows (Internet Options), or the fake of the tests. All empty when Windows
    ' has none for this user.
    ' WINHTTP_CURRENT_USER_IE_PROXY_CONFIG: fAutoDetect (4 bytes), then 3 pointers from offset PTR_SIZE.
    Cov "PXY01"
    Dim config() As Byte
    If Not gProxyFake Is Nothing Then Cov "PXY02": Set WindowsProxySettings = gProxyFake: Exit Function
    Set WindowsProxySettings = FakeProxySettings(False, "", "", "")
    ReDim config(0 To 4 * PTR_SIZE - 1)
    If WinHttpGetIEProxyConfigForCurrentUser(VarPtr(config(0))) <> 0 Then
        Cov "PXY03"
        Set WindowsProxySettings = FakeProxySettings(LongAt(config, 0) <> 0, TakeString(PtrAt(config, PTR_SIZE)), _
                                                     TakeString(PtrAt(config, 2 * PTR_SIZE)), _
                                                     TakeString(PtrAt(config, 3 * PTR_SIZE)))
    End If
End Function

Public Function ScriptProxy(ByVal url As String, ByVal autoDetect As Boolean, ByVal autoConfigUrl As String) As Object
    ' The proxy that the automatic detection (WPAD) or the proxy script (PAC) gives for url, as WinHTTP finds it.
    ' Ok = False when no script was found or it could not run (Error = the Windows error number). Proxy = "" = direct.
    ' WINHTTP_AUTOPROXY_OPTIONS: dwFlags (0), dwAutoDetectFlags (4), lpszAutoConfigUrl (8), lpvReserved (8 + PTR_SIZE),
    ' dwReserved (8 + 2 PTR_SIZE), fAutoLogonIfChallenged (12 + 2 PTR_SIZE). WINHTTP_PROXY_INFO: dwAccessType (0),
    ' lpszProxy (PTR_SIZE), lpszProxyBypass (2 PTR_SIZE).
    Cov "PXY04"
    Dim options() As Byte, info() As Byte, flags As Long, session As LongPtr, agent As String, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Ok") = False: d("Proxy") = "": d("Bypass") = "": d("Error") = 0
    ReDim options(0 To 16 + 2 * PTR_SIZE - 1)
    ReDim info(0 To 3 * PTR_SIZE - 1)
    If autoDetect Then
        flags = AUTOPROXY_AUTO_DETECT
        PutLong options, 4, AUTO_DETECT_TYPE_DHCP Or AUTO_DETECT_TYPE_DNS_A
    End If
    If Len(autoConfigUrl) > 0 Then
        flags = flags Or AUTOPROXY_CONFIG_URL
        PutPtr options, 8, StrPtr(autoConfigUrl)                 ' the ByVal copy lives until the end of this function
    End If
    PutLong options, 0, flags
    PutLong options, 12 + 2 * PTR_SIZE, 1                        ' a script behind a proxy that wants the Windows session
    agent = "jira2dashboard"
    session = WinHttpOpen(StrPtr(agent), ACCESS_TYPE_NO_PROXY, 0, 0, 0)
    If session = 0 Then d("Error") = Err.LastDllError: Set ScriptProxy = d: Exit Function
    If WinHttpGetProxyForUrl(session, StrPtr(url), VarPtr(options(0)), VarPtr(info(0))) <> 0 Then
        Cov "PXY05"
        d("Ok") = True
        d("Proxy") = TakeString(PtrAt(info, PTR_SIZE))           ' empty when the script says DIRECT
        d("Bypass") = TakeString(PtrAt(info, 2 * PTR_SIZE))
    Else
        Cov "PXY06"
        d("Error") = Err.LastDllError
    End If
    WinHttpCloseHandle session
    Set ScriptProxy = d
End Function

Public Function DetectProxy(ByVal url As String) As Object
    ' The proxy for url from the Windows settings. Server = "" means direct. Mode says how it was found, for the log:
    ' it names no address.
    Cov "PXY07"
    Dim ie As Object, script As Object, d As Object, how As String
    Set d = CreateObject("Scripting.Dictionary")
    d("Server") = "": d("Bypass") = ""
    Set ie = WindowsProxySettings()
    If ie("AutoDetect") Or Len(ie("AutoConfigUrl")) > 0 Then
        Cov "PXY08"
        how = IIf(Len(ie("AutoConfigUrl")) > 0, "proxy script (PAC)", "automatic detection (WPAD)")
        Set script = ScriptProxy(url, ie("AutoDetect"), ie("AutoConfigUrl"))
        If script("Ok") Then
            Cov "PXY09"
            d("Server") = script("Proxy"): d("Bypass") = script("Bypass")
            d("Mode") = how & IIf(Len(script("Proxy")) > 0, ": a proxy", ": direct")
            Set DetectProxy = d: Exit Function
        End If
        how = how & " found nothing (Windows error " & script("Error") & "), then "
    End If
    If Len(ie("Proxy")) > 0 Then
        Cov "PXY10"
        d("Server") = ie("Proxy"): d("Bypass") = ie("Bypass")
        d("Mode") = how & "the fixed proxy of the Windows settings"
    Else
        Cov "PXY11"
        d("Mode") = how & "direct (no proxy in the Windows settings)"
    End If
    Set DetectProxy = d
End Function

Private Function TakeString(ByVal p As LongPtr) As String
    ' A string that WinHTTP allocated: copied into a VBA string, then freed.
    Dim n As Long, s As String
    If p = 0 Then Exit Function
    n = lstrlenW(p)
    If n > 0 Then
        s = String$(n, vbNullChar)
        RtlMoveMemory StrPtr(s), p, CLngPtr(n) * 2
    End If
    GlobalFree p
    TakeString = s
End Function

Private Function PtrAt(ByRef buffer() As Byte, ByVal offset As Long) As LongPtr
    Dim p As LongPtr
    RtlMoveMemory VarPtr(p), VarPtr(buffer(offset)), PTR_SIZE
    PtrAt = p
End Function

Private Function LongAt(ByRef buffer() As Byte, ByVal offset As Long) As Long
    Dim v As Long
    RtlMoveMemory VarPtr(v), VarPtr(buffer(offset)), 4
    LongAt = v
End Function

Private Sub PutLong(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As Long)
    RtlMoveMemory VarPtr(buffer(offset)), VarPtr(value), 4
End Sub

Private Sub PutPtr(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As LongPtr)
    RtlMoveMemory VarPtr(buffer(offset)), VarPtr(value), PTR_SIZE
End Sub
