Option Explicit
' HTTP with WinHttp: named proxy, Windows (integrated) auth to the proxy, asynchronous wait that
' keeps Excel responsive (DoEvents), retry on 429 / 503 with Retry-After and exponential backoff.
' Test hook: when gHttpFake is set, no network call is made and queued fake responses are returned.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public gHttpFake As Object      ' Dictionary: "METHOD url" -> Collection of Array(status, body, retryAfter)
Public gHttpLog As Collection   ' one Array(method, url, body) per attempt

Private Const PROXY_SETTING_PROXY As Long = 2      ' HTTPREQUEST_PROXYSETTING_PROXY
Private Const AUTOLOGON_ALWAYS As Long = 0         ' AutoLogonPolicy_Always: send Windows credentials
Private Const AUTOLOGON_NEVER As Long = 2          ' AutoLogonPolicy_Never: do not send them (ProxyWindowsAuth = No)

Public Function NewHttpConfig() As Object
    Cov "HTP01"
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c("ProxyServer") = ""            ' "proxy.client.local:8080"; empty = WinHTTP default (no proxy)
    c("ProxyBypass") = ""            ' "*.client.local;<local>"
    c("ProxyWindowsAuth") = True     ' send the Windows session credentials to the proxy
    c("ResolveMs") = 10000: c("ConnectMs") = 60000: c("SendMs") = 60000: c("ReceiveMs") = 60000   ' 60 s like the PowerShell script
    c("MaxRetries") = 4: c("RetryBaseMs") = 2000: c("RetryMaxMs") = 30000
    c("Authorization") = ""
    Set NewHttpConfig = c
End Function

Public Function HttpRequest(ByVal cfg As Object, ByVal method As String, ByVal url As String, _
                            Optional ByVal body As String = "") As Object
    ' Returns a Dictionary: Status (0 = network error), Text, RetryAfter, Error, Attempts.
    Cov "HTP02"
    Dim attempt As Long, resp As Object
    For attempt = 0 To cfg("MaxRetries")
        Set resp = HttpSendOnce(cfg, method, url, body)
        resp("Attempts") = attempt + 1
        If Not IsRetryable(resp) Then Exit For
        If attempt = cfg("MaxRetries") Then Cov "HTP03": Exit For
        Cov "HTP04"
        WaitMs RetryDelayMs(cfg, attempt, resp("RetryAfter"))
    Next attempt
    Set HttpRequest = resp
End Function

Private Function IsRetryable(ByVal resp As Object) As Boolean
    ' Like the PowerShell extraction script: 429 and every server error (5xx) are retried, 4xx are not.
    ' Only call HttpRequest with POST for read-only endpoints (search, bulkfetch).
    If resp("Status") = 429 Then IsRetryable = True: Exit Function
    If resp("Status") >= 500 Then IsRetryable = True
End Function

Public Function RetryDelayMs(ByVal cfg As Object, ByVal attempt As Long, ByVal retryAfter As String) As Double
    Cov "HTP05"
    Dim backoff As Double, fromHeader As Double
    backoff = cfg("RetryBaseMs") * 2 ^ attempt
    If backoff > cfg("RetryMaxMs") Then backoff = cfg("RetryMaxMs")
    backoff = backoff * Rnd()                                   ' full jitter
    If Len(retryAfter) > 0 Then fromHeader = Val(retryAfter) * 1000# + Rnd() * 500#
    If fromHeader > 120000# Then fromHeader = 120000#           ' never wait more than two minutes on a Retry-After
    RetryDelayMs = IIf(fromHeader > backoff, fromHeader, backoff)
End Function

Private Function HttpSendOnce(ByVal cfg As Object, ByVal method As String, ByVal url As String, _
                              ByVal body As String) As Object
    Dim r As Object, req As Object, raw As Variant
    Set r = CreateObject("Scripting.Dictionary")
    r("Status") = 0: r("Text") = "": r("RetryAfter") = "": r("Error") = ""
    If gHttpLog Is Nothing Then Set gHttpLog = New Collection
    gHttpLog.Add Array(method, url, body)
    If Not gHttpFake Is Nothing Then Cov "HTP06": Set HttpSendOnce = FakeResponse(r, method, url): Exit Function

    On Error GoTo NetError
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open method, url, True                                  ' True = asynchronous
    If Len(cfg("ProxyServer")) > 0 Then req.SetProxy PROXY_SETTING_PROXY, cfg("ProxyServer"), cfg("ProxyBypass")
    ' Windows credentials are meant for the proxy: sent to any server that asks only when a proxy is set; never when No.
    If Not cfg("ProxyWindowsAuth") Then
        req.SetAutoLogonPolicy AUTOLOGON_NEVER
    ElseIf Len(cfg("ProxyServer")) > 0 Then
        req.SetAutoLogonPolicy AUTOLOGON_ALWAYS
    End If                                                      ' else: the WinHTTP default (intranet only)
    req.SetTimeouts cfg("ResolveMs"), cfg("ConnectMs"), cfg("SendMs"), cfg("ReceiveMs")
    req.SetRequestHeader "Accept", "application/json"
    req.SetRequestHeader "Accept-Charset", "utf-8"
    If Len(cfg("Authorization")) > 0 Then req.SetRequestHeader "Authorization", cfg("Authorization")
    If Len(body) > 0 Then
        req.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        req.Send Utf8Encode(body)
    Else
        req.Send
    End If
    Do While Not req.WaitForResponse(0)                         ' returns at once; keeps Excel alive
        DoEvents
        Sleep 20
    Loop
    r("Status") = req.Status
    raw = req.ResponseBody
    If IsArray(raw) Then r("Text") = Utf8Decode(raw)   ' always decode as UTF-8, whatever the header says
    r("RetryAfter") = SafeHeader(req, "Retry-After")
    Set HttpSendOnce = r
    Exit Function
NetError:
    Cov "HTP07"
    r("Status") = 0
    r("Error") = Err.Number & " " & Err.Description
    Set HttpSendOnce = r
End Function

Private Function SafeHeader(ByVal req As Object, ByVal name As String) As String
    On Error Resume Next
    SafeHeader = req.GetResponseHeader(name)
    Err.Clear                                   ' a missing header is normal: leave no stale error
End Function

Public Sub WaitMs(ByVal ms As Double)
    ' Waits without freezing Excel. Timer wraps at midnight: handled.
    Dim t0 As Double, elapsed As Double
    t0 = Timer
    Do
        DoEvents
        Sleep 20
        elapsed = (Timer - t0) * 1000#
        If elapsed < 0 Then elapsed = elapsed + 86400000#
    Loop While elapsed < ms
End Sub

' ---------- Test hook ----------

Public Sub FakeReset()
    Set gHttpFake = CreateObject("Scripting.Dictionary")
    Set gHttpLog = New Collection
End Sub

Public Sub FakeAdd(ByVal method As String, ByVal url As String, ByVal status As Long, _
                   ByVal body As String, Optional ByVal retryAfter As String = "")
    Dim key As String
    key = method & " " & url
    If gHttpFake Is Nothing Then FakeReset
    If Not gHttpFake.Exists(key) Then gHttpFake.Add key, New Collection
    gHttpFake(key).Add Array(status, body, retryAfter)
End Sub

Public Sub FakeOff()
    Set gHttpFake = Nothing
End Sub

Private Function FakeResponse(ByVal r As Object, ByVal method As String, ByVal url As String) As Object
    Dim key As String, q As Collection, item As Variant
    key = method & " " & url
    If gHttpFake.Exists(key) Then Set q = gHttpFake(key)
    If q Is Nothing Then
        r("Status") = 404: r("Text") = "{""fake"":""no response queued""}"
    ElseIf q.Count = 0 Then
        r("Status") = 404: r("Text") = "{""fake"":""queue empty""}"
    Else
        item = q(1): q.Remove 1
        r("Status") = item(0): r("Text") = item(1): r("RetryAfter") = item(2)
    End If
    Set FakeResponse = r
End Function
