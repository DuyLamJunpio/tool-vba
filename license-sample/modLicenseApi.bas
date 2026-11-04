Attribute VB_Name = "modLicenseApi"
Option Explicit

Public Function Api_PostVerify( _
    ByVal rawToken As String, _
    ByVal deviceHash As String, _
    ByRef httpStatus As Long, _
    ByRef responseBody As String) As Boolean

    Dim url As String
    Dim requestJson As String
    Dim requestBytes() As Byte
    Dim http As Object

    On Error GoTo Failed

    httpStatus = 0
    responseBody = vbNullString

    url = Cfg_VerifyUrl()
    If LCase$(Left$(url, 8)) <> "https://" Then Exit Function

    requestJson = Contract_BuildVerifyJson(rawToken, deviceHash)
    If Len(requestJson) = 0 Then Exit Function

    requestBytes = Crypto_Utf8Bytes(requestJson)

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts _
        Cfg_ResolveTimeoutMs(), _
        Cfg_ConnectTimeoutMs(), _
        Cfg_SendTimeoutMs(), _
        Cfg_ReceiveTimeoutMs()

    http.Open "POST", url, False
    http.Option(6) = False  ' Disable automatic redirects.
    http.Option(4) = 0      ' Do not ignore TLS certificate errors.
    http.SetRequestHeader "Content-Type", _
                          "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send requestBytes

    httpStatus = CLng(http.Status)
    responseBody = CStr(http.ResponseText)
    Api_PostVerify = True
    Exit Function

Failed:
    httpStatus = 0
    responseBody = vbNullString
    Api_PostVerify = False
End Function
