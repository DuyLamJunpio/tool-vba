Attribute VB_Name = "modLicenseApiV1"
Option Explicit

' One standard module for the active /api/v1 activation contract.
' Import this file alone. All licensing decisions fail closed until an
' audited Ed25519 verifier and deployment public key are configured.

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type GUID_VALUE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type

Private Type IE_PROXY_CONFIG
    autoDetect As Long
    autoConfigUrl As LongPtr
    proxyServer As LongPtr
    proxyBypass As LongPtr
End Type

Private Type AUTO_PROXY_OPTIONS
    flags As Long
    detectFlags As Long
    configUrl As LongPtr
    reservedPointer As LongPtr
    reservedValue As Long
    autoLogon As Long
End Type

Private Type PROXY_INFO
    accessType As Long
    proxyServer As LongPtr
    proxyBypass As LongPtr
End Type

Private Declare PtrSafe Function ClipOpen Lib "user32" Alias "OpenClipboard" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ClipEmpty Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare PtrSafe Function ClipClose Lib "user32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function ClipSetData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ClipGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function ClipGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ClipGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SYSTEMTIME)
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef value As GUID_VALUE) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, _
    ByVal wideText As LongPtr, ByVal wideCount As Long, _
    ByVal resultData As LongPtr, ByVal outputSize As Long, _
    ByVal defaultChar As LongPtr, ByVal usedDefault As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, _
    ByVal sourceData As LongPtr, ByVal inputSize As Long, _
    ByVal resultData As LongPtr, ByVal outputSize As Long) As Long
Private Declare PtrSafe Function CryptStringToBinaryW Lib "crypt32.dll" ( _
    ByVal source As LongPtr, ByVal sourceLength As Long, _
    ByVal flags As Long, ByVal resultData As LongPtr, _
    ByRef outputLength As Long, ByVal skip As LongPtr, _
    ByVal usedFlags As LongPtr) As Long
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef algorithm As LongPtr, ByVal algorithmName As LongPtr, _
    ByVal implementation As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptHash Lib "bcrypt.dll" ( _
    ByVal algorithm As LongPtr, ByVal secret As LongPtr, _
    ByVal secretLength As Long, ByVal sourceData As LongPtr, _
    ByVal inputLength As Long, ByVal resultData As LongPtr, _
    ByVal outputLength As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal algorithm As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptProtectData Lib "crypt32.dll" ( _
    ByRef sourceData As DATA_BLOB, ByVal description As LongPtr, _
    ByVal entropy As LongPtr, ByVal reserved As LongPtr, _
    ByVal prompt As LongPtr, ByVal flags As Long, _
    ByRef resultData As DATA_BLOB) As Long
Private Declare PtrSafe Function CryptUnprotectData Lib "crypt32.dll" ( _
    ByRef sourceData As DATA_BLOB, ByVal description As LongPtr, _
    ByVal entropy As LongPtr, ByVal reserved As LongPtr, _
    ByVal prompt As LongPtr, ByVal flags As Long, _
    ByRef resultData As DATA_BLOB) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal address As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal address As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal address As LongPtr) As Long
Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "winhttp.dll" ( _
    ByRef config As IE_PROXY_CONFIG) As Long
Private Declare PtrSafe Function WinHttpOpen Lib "winhttp.dll" ( _
    ByVal agent As LongPtr, ByVal accessType As Long, ByVal proxy As LongPtr, _
    ByVal bypass As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function WinHttpSetTimeouts Lib "winhttp.dll" ( _
    ByVal session As LongPtr, ByVal resolveTimeout As Long, ByVal connectTimeout As Long, _
    ByVal sendTimeout As Long, ByVal receiveTimeout As Long) As Long
Private Declare PtrSafe Function WinHttpGetProxyForUrl Lib "winhttp.dll" ( _
    ByVal session As LongPtr, ByVal url As LongPtr, _
    ByRef options As AUTO_PROXY_OPTIONS, ByRef proxy As PROXY_INFO) As Long
Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp.dll" (ByVal session As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal destination As LongPtr, ByVal source As LongPtr, _
    ByVal byteCount As LongPtr)

Private Const CP_UTF8 As Long = 65001
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = 1
Private Const CACHE_APP As String = "DigitalLicenseVbaV1"
Private Const CACHE_SECTION As String = "Activation"

' Ed25519 point in extended coordinates; each field element is 16 limbs of 16 bits.
Private Type EdPoint
    X(0 To 15) As Double
    Y(0 To 15) As Double
    Z(0 To 15) As Double
    T(0 To 15) As Double
End Type

Private mEdReady As Boolean
Private mEdD(0 To 15) As Double
Private mEdD2(0 To 15) As Double
Private mEdBaseX(0 To 15) As Double
Private mEdBaseY(0 To 15) As Double
Private mEdSqrtMinusOne(0 To 15) As Double
Private mEdOrder(0 To 31) As Long

Private mChecked As Boolean
Private mAllowedUntilUnix As Double
Private mLastValidatedUnix As Double
Private mExpiresUnix As Double
Private mRevalidateSeconds As Double
Private mOfflineHours As Double
Private mToken As String
Private mFingerprint As String
Private mFeatures As String
Private mErrorCode As String
Private mCacheClearedForActivation As Boolean
Private mMaxSeenUnix As Double
Private mSavedMaxSeenUnix As Double
Private mIssuedUnix As Double
Private mRefreshCycleBase As Double
Private mRefreshAttemptCycle As Double
Private mRefreshInProgress As Boolean
Private mHardDenied As Boolean
Private mClosing As Boolean
Private mHadAllowedSession As Boolean

' One cache slot per product: products sharing a slot would overwrite each other's activation.
Private Function CacheSection() As String
    CacheSection = CACHE_SECTION & "_" & Cfg_ProductKey()
End Function

' Puts the one-time user code on the clipboard so the customer can paste it (Ctrl+V).
' Only the short user code is copied; never a device code, token or fingerprint.
Private Function CopyTextToClipboard(ByVal text As String) As Boolean
    Const CF_UNICODETEXT As Long = 13
    Const GHND As Long = &H42
    Dim hMem As LongPtr
    Dim pMem As LongPtr
    If Len(text) = 0 Then Exit Function
    hMem = ClipGlobalAlloc(GHND, (Len(text) + 1) * 2)
    If hMem = 0 Then Exit Function
    pMem = ClipGlobalLock(hMem)
    If pMem = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    CopyMemory pMem, StrPtr(text), Len(text) * 2
    ClipGlobalUnlock hMem
    If ClipOpen(0) = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    ClipEmpty
    If ClipSetData(CF_UNICODETEXT, hMem) = 0 Then
        GlobalFree hMem
    Else
        CopyTextToClipboard = True
    End If
    ClipClose
End Function

Public Function Cfg_ApiBaseUrl() As String
    ' Example shape after deployment: https://host.example/api/v1
    Cfg_ApiBaseUrl = "https://shop.sunprime.vn/api/v1"
End Function

Public Function Cfg_ProductKey() As String
    ' Must equal products.product_key for the target workbook.
    Cfg_ProductKey = vbNullString
End Function

Public Function Cfg_ProductVersion() As String
    ' Exact product version of the workbook, for example 1.2.3.
    Cfg_ProductVersion = vbNullString
End Function

Public Function Cfg_ActivationKeyIds() As Variant
    ' Active/previous key IDs. Configure the second pair only during rotation.
    Cfg_ActivationKeyIds = Array("k-d7b49d36a6111497", vbNullString)
End Function

Public Function Cfg_ActivationPublicKeysBase64() As Variant
    ' Raw 32-byte public key as standard Base64; never a private key.
    Cfg_ActivationPublicKeysBase64 = Array( _
        "BdUaYauU+Te7xcYAb7ShcTlNUrQ47uRAsJqKboQm63s=", vbNullString)
End Function

Public Function Cfg_ProtectedSheetNames() As Variant
    ' Example: Array("Premium", "Reports"). Keep a separate guide sheet visible.
    Cfg_ProtectedSheetNames = Array()
End Function

Public Function Cfg_ProxyServer() As String
    ' Empty uses Windows user/PAC settings, then the WinHTTP system default.
    ' Optional override: proxy.example:8080. Never put credentials here.
    Cfg_ProxyServer = vbNullString
End Function

Public Function Cfg_ActivationWebUrl() As String
    ' Approved customer page for entering a one-time user code.
    ' Example shape only: https://shop.example/account/activate
    ' Never add a token, fingerprint or device code to this URL.
    Cfg_ActivationWebUrl = "https://shop.sunprime.vn/account/activate"
End Function

Public Function Cfg_ResolveTimeoutMs() As Long
    Cfg_ResolveTimeoutMs = 5000
End Function

Public Function Cfg_ConnectTimeoutMs() As Long
    Cfg_ConnectTimeoutMs = 5000
End Function

Public Function Cfg_SendTimeoutMs() As Long
    Cfg_SendTimeoutMs = 5000
End Function

Public Function Cfg_ReceiveTimeoutMs() As Long
    Cfg_ReceiveTimeoutMs = 10000
End Function

Public Sub Auto_Open()
    On Error GoTo CloseUnlicensed
    License_OnOpen
    If License_PremiumAllowed() Then Exit Sub
    If NeedsDeviceActivation(mErrorCode, mCacheClearedForActivation) And _
       ConfigReady() And _
       IsHttpsUrl(Cfg_ActivationWebUrl()) Then
        If License_ActivateWithDeviceCode() Then Exit Sub
    End If
    License_ShowGuidance
CloseUnlicensed:
    ' Close only the workbook containing this module. Never quit Excel or save
    ' changes to an unlicensed workbook during the startup gate.
    If Not License_PremiumAllowed() Then
        License_OnClose
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

Public Sub Auto_Close()
    ' Excel runs Auto_Close before its "Save changes?" prompt. With unsaved changes, wait:
    ' Save hides the sheets in Workbook_BeforeSave, Don't Save keeps the saved (hidden) file,
    ' and Cancel keeps the workbook open with its access (test CL-1, DEC-115 addendum 3).
    If ThisWorkbook.Saved Then License_OnClose
End Sub

Public Function License_OnClose() As Boolean
    ' Also call from ThisWorkbook.Workbook_BeforeClose for programmatic closes.
    mClosing = True
    mAllowedUntilUnix = 0
    License_OnClose = ApplyProtectedSheetVisibility(False)
    If Not License_OnClose Then mErrorCode = "SHEET_CONFIG_INVALID"
End Function

Public Function License_ActivateWithDeviceCode() As Boolean
    Dim fingerprint As String
    Dim responseText As String
    Dim statusCode As Long
    Dim deviceCode As String
    Dim userCode As String
    Dim expiresIn As Double
    Dim expiresAt As Double
    Dim body As String
    Dim issuedToken As String
    Dim previousMaxSeen As Double
    Dim cachedToken As String
    Dim cachedFingerprint As String
    Dim cachedTime As Double
    Dim cachedMaxSeen As Double

    On Error GoTo Failed
    previousMaxSeen = mMaxSeenUnix
    If CacheLoad(cachedToken, cachedFingerprint, cachedTime, cachedMaxSeen) Then _
        previousMaxSeen = MaxDouble(previousMaxSeen, cachedMaxSeen)
    ResetLicenseState
    mMaxSeenUnix = previousMaxSeen
    mChecked = True
    If Not ApplyProtectedSheetVisibility(False) Then
        mErrorCode = "SHEET_CONFIG_INVALID"
        Exit Function
    End If
    If Not ConfigReady() Or Not IsHttpsUrl(Cfg_ActivationWebUrl()) Then
        mErrorCode = "CONFIG_REQUIRED"
        Exit Function
    End If
    If Not ClockCanUse(UtcUnixNow(), mMaxSeenUnix, 0) Then
        mErrorCode = "CLOCK_ROLLBACK"
        Exit Function
    End If
    fingerprint = DeviceFingerprintHash()
    If Not IsHex64(fingerprint) Then
        mErrorCode = "DEVICE_UNAVAILABLE"
        Exit Function
    End If

    body = BuildActivateJson(Cfg_ProductKey(), Cfg_ProductVersion(), fingerprint, _
                             NewRequestId(), UtcUnixNow())
    If Not ApiPost("/auth/device-code", body, vbNullString, _
                   statusCode, responseText) Then
        mErrorCode = "NETWORK_UNAVAILABLE"
        Exit Function
    End If
    If statusCode < 200 Or statusCode >= 300 Then
        mErrorCode = JsonSimpleString(responseText, "code")
        If Len(mErrorCode) = 0 Then mErrorCode = "SERVER_REJECTED"
        Exit Function
    End If
    deviceCode = JsonSimpleString(responseText, "device_code")
    userCode = JsonSimpleString(responseText, "user_code")
    expiresIn = JsonNumber(responseText, "expires_in")
    If Not IsDeviceCode(deviceCode) Or Not IsUserCode(userCode) Or _
       expiresIn < 30# Or expiresIn > 1800# Then
        mErrorCode = "SERVER_RESPONSE_INVALID"
        Exit Function
    End If
    expiresAt = UtcUnixNow() + expiresIn

    Dim codeHint As String
    If CopyTextToClipboard(userCode) Then
        codeHint = "Ma da duoc sao chep san: tren trang web chi can bam Ctrl+V."
    Else
        codeHint = "Hay nhap ma nay tren trang web."
    End If
    If MsgBox("Ma kich hoat cho may nay: " & userCode & vbCrLf & codeHint & vbCrLf & vbCrLf & _
              "Chon Yes de mo trang dang nhap va xac nhan ma." & vbCrLf & _
              "Chon No de dong workbook.", _
              vbYesNo + vbInformation, "Kich hoat giay phep") <> vbYes Then
        mErrorCode = "ACTIVATION_CANCELLED"
        Exit Function
    End If
    ThisWorkbook.FollowHyperlink Address:=Cfg_ActivationWebUrl()

    Do While UtcUnixNow() < expiresAt
        If MsgBox("Sau khi xac nhan ma tren trang web, chon OK de kiem tra." & vbCrLf & _
                  "Chon Cancel hoac X de dong workbook.", _
                  vbOKCancel + vbInformation, "Cho xac nhan kich hoat") <> vbOK Then
            mErrorCode = "ACTIVATION_CANCELLED"
            Exit Function
        End If

        body = BuildDeviceCodeRedeemJson(deviceCode, fingerprint, _
                                         NewRequestId(), UtcUnixNow(), _
                                         SafeDeviceName(Environ$("COMPUTERNAME")))
        If Not ApiPost("/auth/device-code/redeem", body, vbNullString, _
                       statusCode, responseText) Then
            mErrorCode = "NETWORK_UNAVAILABLE"
            Exit Function
        End If
        If statusCode = 200 Then
            issuedToken = JsonSimpleString(responseText, "token")
            If Len(issuedToken) = 0 Then
                mErrorCode = "SERVER_RESPONSE_INVALID"
                Exit Function
            End If
            License_ActivateWithDeviceCode = AcceptServerToken(issuedToken, fingerprint)
            Exit Function
        End If
        mErrorCode = JsonSimpleString(responseText, "code")
        If mErrorCode <> "AUTHORIZATION_PENDING" Then
            If Len(mErrorCode) = 0 Then mErrorCode = "SERVER_REJECTED"
            Exit Function
        End If
    Loop
    mErrorCode = "ACTIVATION_EXPIRED"
    Exit Function
Failed:
    ResetLicenseState
    mChecked = True
    ApplyProtectedSheetVisibility False
    mErrorCode = "LOCAL_ERROR"
End Function

Public Sub License_OnOpen()
    Dim cachedToken As String
    Dim cachedFingerprint As String
    Dim cachedTime As Double
    Dim currentFingerprint As String
    Dim expiry As Double
    Dim interval As Double
    Dim graceHours As Double
    Dim featuresJson As String
    Dim nowUnix As Double
    Dim canUseCached As Boolean
    Dim cachedMaxSeen As Double
    Dim issuedUnix As Double

    On Error GoTo Failed
    ResetLicenseState
    mChecked = True
    If Not ApplyProtectedSheetVisibility(False) Then
        mErrorCode = "SHEET_CONFIG_INVALID"
        Exit Sub
    End If
    If Not ConfigReady() Then
        mErrorCode = "CONFIG_REQUIRED"
        Exit Sub
    End If
    currentFingerprint = DeviceFingerprintHash()
    If Len(currentFingerprint) <> 64 Then
        mErrorCode = "DEVICE_UNAVAILABLE"
        Exit Sub
    End If
    If Not CacheLoad(cachedToken, cachedFingerprint, cachedTime, cachedMaxSeen) Then
        mErrorCode = "ACTIVATION_REQUIRED"
        Exit Sub
    End If
    If StrComp(currentFingerprint, cachedFingerprint, vbBinaryCompare) <> 0 Then
        InvalidateCachedActivation "DEVICE_MISMATCH"
        Exit Sub
    End If
    nowUnix = UtcUnixNow()
    If Not VerifyTokenClaims(cachedToken, nowUnix, expiry, interval, _
                             graceHours, featuresJson, issuedUnix, True) Then
        InvalidateCachedActivation "TOKEN_INVALID"
        Exit Sub
    End If
    If Not ClockCanUse(nowUnix, cachedMaxSeen, issuedUnix) Then
        mErrorCode = "CLOCK_ROLLBACK"
        Exit Sub
    End If
    StageVerifiedCredential cachedToken, currentFingerprint, cachedTime, _
                            expiry, interval, graceHours, featuresJson, _
                            issuedUnix, cachedMaxSeen
    mSavedMaxSeenUnix = cachedMaxSeen
    If Not ObserveClock(nowUnix) Then Exit Sub
    If PolicyRequiresOnline(nowUnix, cachedTime, expiry, interval, graceHours) Then
        If AttemptRevalidation(nowUnix) Then Exit Sub
        If mErrorCode <> "NETWORK_UNAVAILABLE" Then Exit Sub
    End If
    nowUnix = UtcUnixNow()
    If Not ObserveClock(nowUnix) Then Exit Sub
    canUseCached = PolicyCanUse(nowUnix, cachedTime, expiry, graceHours * 3600#)
    If Not canUseCached Then
        mErrorCode = "REVALIDATION_REQUIRED"
        Exit Sub
    End If
    AcceptVerifiedState cachedToken, currentFingerprint, cachedTime, _
                        expiry, interval, graceHours, featuresJson, _
                        issuedUnix, mMaxSeenUnix, True
    Exit Sub

Failed:
    ResetLicenseState
    mChecked = True
    mErrorCode = "LOCAL_ERROR"
End Sub

Public Function License_PremiumAllowed() As Boolean
    Dim currentUnix As Double
    Dim currentCycle As Double
    On Error GoTo Denied
    If mClosing Or mRefreshInProgress Then Exit Function
    If Not mChecked Then License_OnOpen
    currentUnix = UtcUnixNow()
    If Len(mToken) > 0 Then
        If Not ObserveClock(currentUnix) Then GoTo Denied
        If currentUnix >= mLastValidatedUnix + mRevalidateSeconds Or _
           Not VerifiedStateCanUse(currentUnix) Then
            currentCycle = RevalidationCycle(currentUnix, mLastValidatedUnix, mRevalidateSeconds)
            If CanAttemptRefresh(currentCycle, mRefreshAttemptCycle, _
                                 mLastValidatedUnix, mRefreshCycleBase, mRefreshInProgress) Then
                If Not AttemptRevalidation(currentUnix) Then
                    currentUnix = UtcUnixNow()
                    If Not ObserveClock(currentUnix) Then GoTo Denied
                    If mErrorCode = "NETWORK_UNAVAILABLE" And Not mHardDenied Then
                        If PolicyCanUse(currentUnix, mLastValidatedUnix, mExpiresUnix, _
                                        mOfflineHours * 3600#) Then
                            AcceptVerifiedState mToken, mFingerprint, mLastValidatedUnix, _
                                                mExpiresUnix, mRevalidateSeconds, mOfflineHours, _
                                                mFeatures, mIssuedUnix, mMaxSeenUnix, True
                        Else
                            mErrorCode = "REVALIDATION_REQUIRED"
                        End If
                    End If
                End If
            End If
        End If
    End If
    currentUnix = UtcUnixNow()
    If Len(mToken) > 0 Then
        If Not ObserveClock(currentUnix) Then GoTo Denied
    End If
    License_PremiumAllowed = VerifiedStateCanUse(currentUnix)
    If Not ApplyProtectedSheetVisibility(License_PremiumAllowed) Then
        mAllowedUntilUnix = 0
        mErrorCode = "SHEET_CONFIG_INVALID"
        License_PremiumAllowed = False
    End If
    If Not License_PremiumAllowed And mHadAllowedSession And Not mClosing Then CloseUnlicensedSession
    Exit Function
Denied:
    mAllowedUntilUnix = 0
    ApplyProtectedSheetVisibility False
    If mHadAllowedSession And Not mClosing Then CloseUnlicensedSession
    License_PremiumAllowed = False
End Function

Private Sub CloseUnlicensedSession()
    On Error Resume Next
    License_OnClose
    License_ShowGuidance
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Function VerifiedStateCanUse(ByVal currentUnix As Double) As Boolean
    VerifiedStateCanUse = _
        (Not mClosing And Not mHardDenied And mAllowedUntilUnix > 0 And _
         ClockCanUse(currentUnix, mMaxSeenUnix, mIssuedUnix) And _
         currentUnix >= mLastValidatedUnix And currentUnix < mExpiresUnix And _
         currentUnix <= mAllowedUntilUnix)
End Function

Private Function AttemptRevalidation(ByVal currentUnix As Double) As Boolean
    On Error GoTo Failed
    If mClosing Or mRefreshInProgress Or Len(mToken) = 0 Then Exit Function
    mRefreshInProgress = True
    mRefreshCycleBase = mLastValidatedUnix
    mRefreshAttemptCycle = RevalidationCycle(currentUnix, mLastValidatedUnix, mRevalidateSeconds)
    mAllowedUntilUnix = 0
    AttemptRevalidation = ValidateWithDeviceToken(mToken, mFingerprint)
    If Not AttemptRevalidation And mErrorCode <> "NETWORK_UNAVAILABLE" Then mHardDenied = True
    mRefreshInProgress = False
    Exit Function
Failed:
    mRefreshInProgress = False
    mHardDenied = True
    mAllowedUntilUnix = 0
    mErrorCode = "LOCAL_ERROR"
End Function

Private Function ValidateWithDeviceToken( _
    ByVal activationToken As String, ByVal fingerprint As String) As Boolean

    Dim body As String
    Dim responseText As String
    Dim httpStatus As Long
    Dim refreshedToken As String
    On Error GoTo Failed
    body = BuildValidateJson(activationToken, fingerprint, _
                             NewRequestId(), UtcUnixNow())
    If Not ApiPost("/activations/validate-device", body, activationToken, _
                   httpStatus, responseText) Then
        mErrorCode = "NETWORK_UNAVAILABLE"
        Exit Function
    End If
    If httpStatus < 200 Or httpStatus >= 300 Then
        If httpStatus >= 500 Then
            mErrorCode = "SERVER_REJECTED"
            Exit Function
        End If
        If httpStatus = 429 Then
            mErrorCode = "RATE_LIMITED"
            Exit Function
        End If
        mErrorCode = JsonSimpleString(responseText, "code")
        If Len(mErrorCode) = 0 Then mErrorCode = "SERVER_REJECTED"
        If ShouldClearCacheForServerError(mErrorCode, httpStatus) Then _
            InvalidateCachedActivation mErrorCode
        Exit Function
    End If
    refreshedToken = JsonSimpleString(responseText, "token")
    If Len(refreshedToken) = 0 Then
        mErrorCode = "SERVER_RESPONSE_INVALID"
        Exit Function
    End If
    ValidateWithDeviceToken = AcceptServerToken(refreshedToken, fingerprint)
    Exit Function
Failed:
    mErrorCode = "LOCAL_ERROR"
End Function

Public Function License_FeatureAllowed(ByVal featureKey As String) As Boolean
    If Not License_PremiumAllowed() Then Exit Function
    If Not IsSafeIdentifier(featureKey) Then Exit Function
    License_FeatureAllowed = _
        (InStr(1, mFeatures, """" & featureKey & """", vbBinaryCompare) > 0)
End Function

Private Function AcceptServerToken( _
    ByVal issuedToken As String, ByVal fingerprint As String) As Boolean

    Dim currentUnix As Double
    Dim expiry As Double
    Dim interval As Double
    Dim graceHours As Double
    Dim featuresJson As String
    Dim issuedUnix As Double

    currentUnix = UtcUnixNow()
    If Not VerifyTokenClaims(issuedToken, currentUnix, expiry, _
                             interval, graceHours, featuresJson, issuedUnix) Then
        mErrorCode = "TOKEN_INVALID"
        Exit Function
    End If
    If Not ClockCanUse(currentUnix, mMaxSeenUnix, issuedUnix) Then
        mErrorCode = "CLOCK_ROLLBACK"
        Exit Function
    End If
    mMaxSeenUnix = MaxDouble(mMaxSeenUnix, MaxDouble(currentUnix, issuedUnix))
    If Not CacheSave(issuedToken, fingerprint, currentUnix, mMaxSeenUnix) Then
        mErrorCode = "CACHE_UNAVAILABLE"
        Exit Function
    End If
    mSavedMaxSeenUnix = mMaxSeenUnix
    AcceptVerifiedState issuedToken, fingerprint, currentUnix, _
                        expiry, interval, graceHours, featuresJson, _
                        issuedUnix, mMaxSeenUnix, False
    AcceptServerToken = (mAllowedUntilUnix > 0)
End Function

Private Sub AcceptVerifiedState( _
    ByVal token As String, ByVal fingerprint As String, _
    ByVal validatedUnix As Double, ByVal expiry As Double, _
    ByVal interval As Double, ByVal graceHours As Double, _
    ByVal featuresJson As String, ByVal issuedUnix As Double, _
    ByVal maxSeenUnix As Double, ByVal offline As Boolean)

    StageVerifiedCredential token, fingerprint, validatedUnix, expiry, interval, _
                            graceHours, featuresJson, issuedUnix, maxSeenUnix
    If offline Then
        mAllowedUntilUnix = MinDouble(expiry, validatedUnix + graceHours * 3600#)
    Else
        mAllowedUntilUnix = MinDouble(expiry, validatedUnix + interval)
        mRefreshCycleBase = validatedUnix
        mRefreshAttemptCycle = -1
        mHardDenied = False
    End If
    mErrorCode = vbNullString
    If Not ApplyProtectedSheetVisibility(True) Then
        mAllowedUntilUnix = 0
        mErrorCode = "SHEET_CONFIG_INVALID"
    Else
        mHadAllowedSession = True
    End If
End Sub

Private Sub StageVerifiedCredential( _
    ByVal token As String, ByVal fingerprint As String, _
    ByVal validatedUnix As Double, ByVal expiry As Double, _
    ByVal interval As Double, ByVal graceHours As Double, _
    ByVal featuresJson As String, ByVal issuedUnix As Double, ByVal maxSeenUnix As Double)
    mToken = token
    mFingerprint = fingerprint
    mLastValidatedUnix = validatedUnix
    mExpiresUnix = expiry
    mRevalidateSeconds = interval
    mOfflineHours = graceHours
    mFeatures = featuresJson
    mIssuedUnix = issuedUnix
    mMaxSeenUnix = MaxDouble(mMaxSeenUnix, MaxDouble(maxSeenUnix, issuedUnix))
End Sub

Private Sub ResetLicenseState()
    mAllowedUntilUnix = 0
    mLastValidatedUnix = 0
    mExpiresUnix = 0
    mRevalidateSeconds = 0
    mOfflineHours = 0
    mToken = vbNullString
    mFingerprint = vbNullString
    mFeatures = vbNullString
    mErrorCode = vbNullString
    mCacheClearedForActivation = False
    mMaxSeenUnix = 0
    mSavedMaxSeenUnix = 0
    mIssuedUnix = 0
    mRefreshCycleBase = 0
    mRefreshAttemptCycle = -1
    mRefreshInProgress = False
    mHardDenied = False
    mClosing = False
    mHadAllowedSession = False
End Sub

Private Function ConfigReady() As Boolean
    ConfigReady = _
        (LCase$(Left$(Cfg_ApiBaseUrl(), 8)) = "https://" And _
         Len(Cfg_ProductKey()) > 0 And _
         Len(Cfg_ProductVersion()) > 0 And _
         IsSafeIdentifier(Cfg_ProductKey()) And _
         IsSemVer(Cfg_ProductVersion()) And _
         KeyPairsValid(Cfg_ActivationKeyIds(), Cfg_ActivationPublicKeysBase64()))
End Function

Private Function ArrayItemCount(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    If Not IsArray(values) Then
        ArrayItemCount = -1
        Exit Function
    End If
    ArrayItemCount = UBound(values) - LBound(values) + 1
EmptyArray:
End Function

Private Function KeyPairsValid(ByVal keyIds As Variant, ByVal publicKeys As Variant) As Boolean
    Dim slot As Long
    Dim keyId As String
    Dim keyValue As String
    Dim decodedKey() As Byte
    On Error GoTo Invalid
    If ArrayItemCount(keyIds) <> 2 Or ArrayItemCount(publicKeys) <> 2 Then Exit Function
    For slot = 0 To 1
        keyId = CStr(keyIds(LBound(keyIds) + slot))
        keyValue = CStr(publicKeys(LBound(publicKeys) + slot))
        If slot = 0 Or Len(keyId) > 0 Or Len(keyValue) > 0 Then
            If Not IsSafeIdentifier(keyId) Or Len(keyValue) = 0 Then Exit Function
            decodedKey = Base64Bytes(keyValue)
            If ByteCount(decodedKey) <> 32 Then Exit Function
        End If
    Next slot
    If Len(CStr(keyIds(LBound(keyIds) + 1))) > 0 Then
        If StrComp(CStr(keyIds(LBound(keyIds))), CStr(keyIds(LBound(keyIds) + 1)), vbBinaryCompare) = 0 Then Exit Function
    End If
    KeyPairsValid = True
Invalid:
End Function

Private Function FindPublicKey(ByVal keyId As String, ByVal keyIds As Variant, _
                               ByVal publicKeys As Variant) As String
    Dim slot As Long
    If Len(keyId) = 0 Then Exit Function
    If Not KeyPairsValid(keyIds, publicKeys) Then Exit Function
    For slot = 0 To 1
        If StrComp(keyId, CStr(keyIds(LBound(keyIds) + slot)), vbBinaryCompare) = 0 Then
            FindPublicKey = CStr(publicKeys(LBound(publicKeys) + slot))
            Exit Function
        End If
    Next slot
End Function

Private Function PublicKeyForKid(ByVal keyId As String) As String
    PublicKeyForKid = FindPublicKey(keyId, Cfg_ActivationKeyIds(), Cfg_ActivationPublicKeysBase64())
End Function

Private Function TokenKeyIdsMatch(ByVal headerKid As String, ByVal payloadKid As String) As Boolean
    TokenKeyIdsMatch = (Len(headerKid) > 0 And StrComp(headerKid, payloadKid, vbBinaryCompare) = 0)
End Function

Private Function IsProtectedSheetName(ByVal sheetName As String, ByVal names As Variant) As Boolean
    Dim candidate As Variant
    For Each candidate In names
        If StrComp(sheetName, CStr(candidate), vbTextCompare) = 0 Then
            IsProtectedSheetName = True
            Exit Function
        End If
    Next candidate
End Function

Private Function ApplyProtectedSheetVisibility(ByVal allowed As Boolean) As Boolean
    Dim names As Variant
    Dim sheetName As Variant
    Dim targetSheet As Object
    Dim workbookSheet As Object
    Dim guideVisible As Boolean
    On Error GoTo Failed
    names = Cfg_ProtectedSheetNames()
    If ArrayItemCount(names) < 0 Then Exit Function
    If ArrayItemCount(names) = 0 Then
        ApplyProtectedSheetVisibility = True
        Exit Function
    End If
    ' Validate all targets and a visible non-premium guide before hiding any.
    For Each sheetName In names
        If Len(CStr(sheetName)) = 0 Then Exit Function
        Set targetSheet = ThisWorkbook.Worksheets(CStr(sheetName))
    Next sheetName
    If Not allowed Then
        For Each workbookSheet In ThisWorkbook.Sheets
            If Not IsProtectedSheetName(CStr(workbookSheet.Name), names) Then
                If workbookSheet.Visible = xlSheetVisible Then guideVisible = True
            End If
        Next workbookSheet
        If Not guideVisible Then Exit Function
    End If
    For Each sheetName In names
        Set targetSheet = ThisWorkbook.Worksheets(CStr(sheetName))
        If allowed Then
            targetSheet.Visible = xlSheetVisible
        Else
            targetSheet.Visible = xlSheetVeryHidden
        End If
    Next sheetName
    ApplyProtectedSheetVisibility = True
    Exit Function
Failed:
    ApplyProtectedSheetVisibility = False
End Function

Public Function License_BeforeSave() As Boolean
    License_BeforeSave = ApplyProtectedSheetVisibility(False)
    If Not License_BeforeSave Then mErrorCode = "SHEET_CONFIG_INVALID"
End Function

Public Sub License_AfterSave()
    If Not mClosing Then License_PremiumAllowed
End Sub

Private Function ClockCanUse(ByVal currentUnix As Double, ByVal maxSeenUnix As Double, _
                             ByVal issuedUnix As Double) As Boolean
    ClockCanUse = (currentUnix > 0 And currentUnix >= maxSeenUnix And currentUnix >= issuedUnix)
End Function

Private Function ObserveClock(ByVal currentUnix As Double) As Boolean
    If Not ClockCanUse(currentUnix, mMaxSeenUnix, mIssuedUnix) Then
        mErrorCode = "CLOCK_ROLLBACK"
        mAllowedUntilUnix = 0
        Exit Function
    End If
    mMaxSeenUnix = MaxDouble(mMaxSeenUnix, MaxDouble(currentUnix, mIssuedUnix))
    If Len(mToken) > 0 And mMaxSeenUnix > mSavedMaxSeenUnix Then
        If Not CacheSave(mToken, mFingerprint, mLastValidatedUnix, mMaxSeenUnix) Then
            mErrorCode = "CACHE_UNAVAILABLE"
            mAllowedUntilUnix = 0
            Exit Function
        End If
        mSavedMaxSeenUnix = mMaxSeenUnix
    End If
    ObserveClock = True
End Function

Private Function RevalidationCycle(ByVal currentUnix As Double, ByVal validatedUnix As Double, _
                                   ByVal interval As Double) As Double
    RevalidationCycle = -1
    If interval <= 0 Or validatedUnix <= 0 Or currentUnix < validatedUnix Then Exit Function
    RevalidationCycle = Fix((currentUnix - validatedUnix) / interval)
End Function

Private Function CanAttemptRefresh(ByVal currentCycle As Double, ByVal attemptedCycle As Double, _
                                   ByVal validatedUnix As Double, ByVal cycleBase As Double, _
                                   ByVal refreshing As Boolean) As Boolean
    If refreshing Or currentCycle < 0 Then Exit Function
    CanAttemptRefresh = (validatedUnix <> cycleBase Or currentCycle > attemptedCycle)
End Function

Private Function PolicyCanUse( _
    ByVal currentUnix As Double, ByVal validatedUnix As Double, _
    ByVal expiresUnix As Double, ByVal allowedSeconds As Double) As Boolean

    PolicyCanUse = _
        (validatedUnix > 0 And allowedSeconds > 0 And _
         currentUnix >= validatedUnix And _
         currentUnix < expiresUnix And _
         currentUnix <= validatedUnix + allowedSeconds)
End Function

Private Function PolicyRequiresOnline( _
    ByVal currentUnix As Double, ByVal validatedUnix As Double, _
    ByVal expiresUnix As Double, ByVal interval As Double, _
    ByVal graceHours As Double) As Boolean

    ' An expired token is only a renewal credential (DEC-104).
    ' Zero offline grace always requires a successful server check on open.
    PolicyRequiresOnline = _
        (currentUnix >= validatedUnix + interval Or _
         Not PolicyCanUse(currentUnix, validatedUnix, expiresUnix, graceHours * 3600#))
End Function

Private Function NeedsDeviceActivation( _
    ByVal status As String, ByVal cacheCleared As Boolean) As Boolean
    Select Case status
        Case "ACTIVATION_REQUIRED"
            NeedsDeviceActivation = True
        Case "TOKEN_INVALID", "DEVICE_REVOKED", "DEVICE_MISMATCH"
            NeedsDeviceActivation = cacheCleared
    End Select
End Function

Private Function ShouldClearCacheForServerError( _
    ByVal status As String, ByVal httpStatus As Long) As Boolean

    If httpStatus < 400 Or httpStatus >= 500 Or httpStatus = 429 Then Exit Function
    Select Case status
        Case "TOKEN_INVALID", "DEVICE_REVOKED", "DEVICE_MISMATCH"
            ShouldClearCacheForServerError = True
    End Select
End Function

Private Sub InvalidateCachedActivation(ByVal status As String)
    mErrorCode = status
    mCacheClearedForActivation = CacheClear()
    If Not mCacheClearedForActivation Then mErrorCode = "CACHE_UNAVAILABLE"
    mToken = vbNullString
    mFeatures = vbNullString
    mAllowedUntilUnix = 0
End Sub

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function BuildActivateJson( _
    ByVal productKey As String, ByVal version As String, _
    ByVal fingerprint As String, _
    ByVal requestId As String, ByVal timestamp As Double) As String

    If Not IsSafeIdentifier(productKey) Then Err.Raise 5
    If Not IsSemVer(version) Then Err.Raise 5
    If Not IsHex64(fingerprint) Then Err.Raise 5
    BuildActivateJson = _
        "{""product_key"":""" & productKey & _
        """,""version"":""" & version & _
        """,""fp_hash"":""" & fingerprint & _
        """,""request_id"":""" & requestId & _
        """,""timestamp"":" & Format$(timestamp, "0") & "}"
End Function

Private Function BuildValidateJson( _
    ByVal activationToken As String, ByVal fingerprint As String, _
    ByVal requestId As String, ByVal timestamp As Double) As String

    If Not IsJwtCharacters(activationToken) Then Err.Raise 5
    If Not IsHex64(fingerprint) Then Err.Raise 5
    BuildValidateJson = _
        "{""activation_token"":""" & activationToken & _
        """,""fp_hash"":""" & fingerprint & _
        """,""request_id"":""" & requestId & _
        """,""timestamp"":" & Format$(timestamp, "0") & "}"
End Function

Private Function BuildDeviceCodeRedeemJson( _
    ByVal deviceCode As String, ByVal fingerprint As String, _
    ByVal requestId As String, ByVal timestamp As Double, _
    Optional ByVal deviceName As String = vbNullString) As String

    If Not IsDeviceCode(deviceCode) Then Err.Raise 5
    If Not IsHex64(fingerprint) Then Err.Raise 5
    deviceName = SafeDeviceName(deviceName)
    BuildDeviceCodeRedeemJson = _
        "{""device_code"":""" & deviceCode & _
        """,""fp_hash"":""" & fingerprint & _
        """,""request_id"":""" & requestId & _
        """,""timestamp"":" & Format$(timestamp, "0")
    If Len(deviceName) > 0 Then
        BuildDeviceCodeRedeemJson = BuildDeviceCodeRedeemJson & _
            ",""device_name"":""" & deviceName & """"
    End If
    BuildDeviceCodeRedeemJson = BuildDeviceCodeRedeemJson & "}"
End Function

' The Windows computer name lets the customer and the admin tell computers apart on the
' store ("Thiet bi"). Only letters, digits, "-", "_" and "." are kept (so the value never needs
' JSON escaping), at most 63 characters. It is a display label, never part of the fingerprint.
Private Function SafeDeviceName(ByVal value As String) As String
    Const MAX_DEVICE_NAME As Long = 63
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then SafeDeviceName = SafeDeviceName & ch
        If Len(SafeDeviceName) >= MAX_DEVICE_NAME Then Exit For
    Next i
End Function

Private Function ApiPost( _
    ByVal path As String, ByVal body As String, _
    ByVal activationToken As String, ByRef statusCode As Long, _
    ByRef responseText As String) As Boolean

    Dim http As Object
    Dim payload() As Byte
    Dim baseUrl As String
    On Error GoTo Failed

    statusCode = 0
    responseText = vbNullString
    baseUrl = Cfg_ApiBaseUrl()
    If LCase$(Left$(baseUrl, 8)) <> "https://" Then Exit Function
    If Right$(baseUrl, 1) = "/" Then baseUrl = Left$(baseUrl, Len(baseUrl) - 1)
    If Len(activationToken) = 0 Then
        If path <> "/auth/device-code" And _
           path <> "/auth/device-code/redeem" Then Exit Function
    Else
        If path <> "/activations/validate-device" Then Exit Function
        If Not IsJwtCharacters(activationToken) Then Exit Function
    End If

    payload = Utf8Bytes(body)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts Cfg_ResolveTimeoutMs(), Cfg_ConnectTimeoutMs(), _
                     Cfg_SendTimeoutMs(), Cfg_ReceiveTimeoutMs()
    http.Option(9) = &H800 Or &H2000
    http.Open "POST", baseUrl & path, False
    If Not ConfigureHttpProxy(http, baseUrl & path) Then Exit Function
    http.Option(6) = False
    http.Option(4) = 0
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    If Len(activationToken) > 0 Then _
        http.SetRequestHeader "Authorization", "Bearer " & activationToken
    http.Send payload
    statusCode = CLng(http.Status)
    responseText = CStr(http.ResponseText)
    ApiPost = True
    Exit Function

Failed:
    statusCode = 0
    responseText = vbNullString
    ApiPost = False
End Function

Private Function NativeWideText(ByVal address As LongPtr) As String
    Dim charCount As Long
    Dim value As String
    If address = 0 Then Exit Function
    charCount = lstrlenW(address)
    If charCount < 0 Or charCount > 32767 Then Err.Raise 5
    If charCount = 0 Then Exit Function
    value = String$(charCount, vbNullChar)
    CopyMemory StrPtr(value), address, CLngPtr(charCount) * 2
    NativeWideText = value
End Function

Private Function ConfigureHttpProxy(ByVal http As Object, ByVal requestUrl As String) As Boolean
    Dim userProxy As IE_PROXY_CONFIG
    Dim autoOptions As AUTO_PROXY_OPTIONS
    Dim resolvedProxy As PROXY_INFO
    Dim sessionHandle As LongPtr
    Dim agent As String
    Dim overrideProxy As String
    Dim staticProxy As String
    Dim bypass As String
    On Error GoTo Cleanup
    overrideProxy = Cfg_ProxyServer()
    If Len(overrideProxy) > 0 Then
        If InStr(overrideProxy, vbCr) > 0 Or InStr(overrideProxy, vbLf) > 0 Or _
           InStr(overrideProxy, "@") > 0 Then Exit Function
        http.SetProxy 2, overrideProxy
        ConfigureHttpProxy = True
        Exit Function
    End If
    If WinHttpGetIEProxyConfigForCurrentUser(userProxy) <> 0 Then
        staticProxy = NativeWideText(userProxy.proxyServer)
        bypass = NativeWideText(userProxy.proxyBypass)
        If userProxy.autoDetect <> 0 Then
            autoOptions.flags = 1
            autoOptions.detectFlags = 3
        End If
        If userProxy.autoConfigUrl <> 0 Then
            autoOptions.flags = autoOptions.flags Or 2
            autoOptions.configUrl = userProxy.autoConfigUrl
        End If
        If autoOptions.flags <> 0 Then
            agent = "DigitalLicenseVbaV1/Proxy"
            sessionHandle = WinHttpOpen(StrPtr(agent), 1, 0, 0, 0)
            If sessionHandle <> 0 Then
                If WinHttpSetTimeouts(sessionHandle, Cfg_ResolveTimeoutMs(), Cfg_ConnectTimeoutMs(), _
                                      Cfg_SendTimeoutMs(), Cfg_ReceiveTimeoutMs()) = 0 Then GoTo Cleanup
                autoOptions.autoLogon = 0
                If WinHttpGetProxyForUrl(sessionHandle, StrPtr(requestUrl), autoOptions, resolvedProxy) <> 0 Then
                    If resolvedProxy.accessType = 3 Then
                        http.SetProxy 2, NativeWideText(resolvedProxy.proxyServer), NativeWideText(resolvedProxy.proxyBypass)
                    ElseIf resolvedProxy.accessType = 1 Then
                        http.SetProxy 1
                    Else
                        GoTo Cleanup
                    End If
                    ConfigureHttpProxy = True
                    GoTo Cleanup
                End If
            End If
        End If
        If Len(staticProxy) > 0 Then
            http.SetProxy 2, staticProxy, bypass
            ConfigureHttpProxy = True
            GoTo Cleanup
        End If
    End If
    http.SetProxy 0
    ConfigureHttpProxy = True
Cleanup:
    If sessionHandle <> 0 Then WinHttpCloseHandle sessionHandle
    If userProxy.autoConfigUrl <> 0 Then GlobalFree userProxy.autoConfigUrl
    If userProxy.proxyServer <> 0 Then GlobalFree userProxy.proxyServer
    If userProxy.proxyBypass <> 0 Then GlobalFree userProxy.proxyBypass
    If resolvedProxy.proxyServer <> 0 Then GlobalFree resolvedProxy.proxyServer
    If resolvedProxy.proxyBypass <> 0 Then GlobalFree resolvedProxy.proxyBypass
End Function

Private Function VerifyTokenClaims( _
    ByVal token As String, ByVal nowUnix As Double, _
    ByRef expiry As Double, ByRef interval As Double, _
    ByRef graceHours As Double, ByRef featuresJson As String, _
    ByRef issuedUnix As Double, _
    Optional ByVal allowExpiredForRenewal As Boolean = False) As Boolean

    Dim parts() As String
    Dim headerJson As String
    Dim payloadJson As String
    Dim signature() As Byte
    Dim publicKey() As Byte
    Dim message() As Byte
    Dim headerKid As String
    Dim keyValue As String
    On Error GoTo Invalid
    If Not IsJwtCharacters(token) Then Exit Function
    parts = Split(token, ".")
    If UBound(parts) <> 2 Then Exit Function
    headerJson = Utf8Text(Base64UrlBytes(parts(0)))
    If JsonSimpleString(headerJson, "alg") <> "EdDSA" Then Exit Function
    If JsonSimpleString(headerJson, "typ") <> "JWT" Then Exit Function
    headerKid = JsonSimpleString(headerJson, "kid")
    keyValue = PublicKeyForKid(headerKid)
    If Len(keyValue) = 0 Then Exit Function
    signature = Base64UrlBytes(parts(2))
    publicKey = Base64Bytes(keyValue)
    If ByteCount(signature) <> 64 Then Exit Function
    If ByteCount(publicKey) <> 32 Then Exit Function
    message = Utf8Bytes(parts(0) & "." & parts(1))
    If Not VerifyEd25519Detached(message, signature, publicKey) Then Exit Function

    payloadJson = Utf8Text(Base64UrlBytes(parts(1)))
    If Not TokenKeyIdsMatch(headerKid, JsonSimpleString(payloadJson, "kid")) Then Exit Function
    If JsonSimpleString(payloadJson, "prd") <> Cfg_ProductKey() Then Exit Function
    If JsonSimpleString(payloadJson, "ver") <> Cfg_ProductVersion() Then Exit Function
    If Len(JsonSimpleString(payloadJson, "dev")) = 0 Then Exit Function
    If Len(JsonSimpleString(payloadJson, "lic")) = 0 Then Exit Function
    If Len(JsonSimpleString(payloadJson, "jti")) = 0 Then Exit Function
    expiry = JsonNumber(payloadJson, "exp")
    interval = JsonNumber(payloadJson, "rvi")
    graceHours = JsonNumber(payloadJson, "ogh")
    issuedUnix = JsonNumber(payloadJson, "iat")
    If expiry <= 0 Or interval <= 0 Or graceHours < 0 Then Exit Function
    If expiry <= nowUnix And Not allowExpiredForRenewal Then Exit Function
    If issuedUnix <= 0 Then Exit Function
    featuresJson = JsonArraySlice(payloadJson, "feat")
    If Len(featuresJson) = 0 Then Exit Function
    VerifyTokenClaims = True
    Exit Function
Invalid:
    VerifyTokenClaims = False
End Function

Private Function VerifyEd25519Detached( _
    ByRef message() As Byte, ByRef signature() As Byte, _
    ByRef publicKey() As Byte) As Boolean
    ' RFC 8032 Ed25519 verification: [S]B must equal R + [h]A where
    ' h = SHA-512(R || A || message) mod L. Public data only, so no
    ' constant-time handling is needed. Any error means "not verified".
    Dim negA As EdPoint
    Dim pointA As EdPoint
    Dim pointB As EdPoint
    Dim packed(0 To 31) As Byte
    Dim scalarH(0 To 31) As Byte
    Dim buffer() As Byte
    Dim digest() As Byte
    Dim messageLength As Long
    Dim i As Long
    On Error GoTo Failed
    If ByteCount(signature) <> 64 Then Exit Function
    If ByteCount(publicKey) <> 32 Then Exit Function
    If LBound(signature) <> 0 Or LBound(publicKey) <> 0 Then Exit Function
    EdInit
    If Not EdScalarBelowOrder(signature, 32) Then Exit Function
    If Not EdUnpackNegated(negA, publicKey) Then Exit Function

    messageLength = ByteCount(message)
    ReDim buffer(0 To 63 + messageLength)
    For i = 0 To 31
        buffer(i) = signature(i)
        buffer(32 + i) = publicKey(i)
    Next i
    For i = 0 To messageLength - 1
        buffer(64 + i) = message(LBound(message) + i)
    Next i
    If Not Sha512Bytes(buffer, digest) Then Exit Function
    EdReduceScalar digest, scalarH

    EdScalarMult pointA, negA, scalarH, 0
    EdBasePoint pointB
    EdScalarMult pointB, pointB, signature, 32
    EdPointAdd pointA, pointB
    EdPackPoint packed, pointA
    For i = 0 To 31
        If packed(i) <> signature(i) Then Exit Function
    Next i
    VerifyEd25519Detached = True
    Exit Function
Failed:
    VerifyEd25519Detached = False
End Function

Private Function Sha512Bytes( _
    ByRef sourceBytes() As Byte, ByRef digest() As Byte) As Boolean
    Dim algorithm As LongPtr
    ReDim digest(0 To 63)
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr("SHA512"), 0, 0) <> 0 Then Exit Function
    If BCryptHash(algorithm, 0, 0, VarPtr(sourceBytes(0)), _
                  ByteCount(sourceBytes), VarPtr(digest(0)), 64) = 0 Then
        Sha512Bytes = True
    End If
    BCryptCloseAlgorithmProvider algorithm, 0
End Function

Private Sub EdInit()
    Dim orderBytes As Variant
    Dim k As Long
    If mEdReady Then Exit Sub
    EdLoadLimbs mEdD, Array(30883, 4953, 19914, 30187, 55467, 16705, 2637, 112, 59544, 30585, 16505, 36039, 65139, 11119, 27886, 20995)
    EdLoadLimbs mEdD2, Array(61785, 9906, 39828, 60374, 45398, 33411, 5274, 224, 53552, 61171, 33010, 6542, 64743, 22239, 55772, 9222)
    EdLoadLimbs mEdBaseX, Array(54554, 36645, 11616, 51542, 42930, 38181, 51040, 26924, 56412, 64982, 57905, 49316, 21502, 52590, 14035, 8553)
    EdLoadLimbs mEdBaseY, Array(26200, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214, 26214)
    EdLoadLimbs mEdSqrtMinusOne, Array(41136, 18958, 6951, 50414, 58488, 44335, 6150, 12099, 55207, 15867, 153, 11085, 57099, 20417, 9344, 11139)
    orderBytes = Array(237, 211, 245, 92, 26, 99, 18, 88, 214, 156, 247, 162, 222, 249, 222, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 16)
    For k = 0 To 31
        mEdOrder(k) = CLng(orderBytes(k))
    Next k
    mEdReady = True
End Sub

Private Sub EdLoadLimbs(ByRef target() As Double, ByVal limbValues As Variant)
    Dim k As Long
    For k = 0 To 15
        target(k) = CDbl(limbValues(k))
    Next k
End Sub

Private Function EdScalarBelowOrder( _
    ByRef scalarBytes() As Byte, ByVal offset As Long) As Boolean
    ' True when the 32-byte little-endian scalar is strictly below L.
    Dim k As Long
    For k = 31 To 0 Step -1
        If scalarBytes(offset + k) < mEdOrder(k) Then
            EdScalarBelowOrder = True
            Exit Function
        End If
        If scalarBytes(offset + k) > mEdOrder(k) Then Exit Function
    Next k
End Function

Private Sub EdCarry(ByRef o() As Double)
    Dim k As Long
    Dim c As Double
    For k = 0 To 15
        o(k) = o(k) + 65536#
        c = Int(o(k) / 65536#)
        If k < 15 Then
            o(k + 1) = o(k + 1) + c - 1#
        Else
            o(0) = o(0) + 38# * (c - 1#)
        End If
        o(k) = o(k) - c * 65536#
    Next k
End Sub

Private Sub EdAdd(ByRef o() As Double, ByRef a() As Double, ByRef b() As Double)
    Dim k As Long
    For k = 0 To 15
        o(k) = a(k) + b(k)
    Next k
End Sub

Private Sub EdSub(ByRef o() As Double, ByRef a() As Double, ByRef b() As Double)
    Dim k As Long
    For k = 0 To 15
        o(k) = a(k) - b(k)
    Next k
End Sub

Private Sub EdMul(ByRef o() As Double, ByRef a() As Double, ByRef b() As Double)
    ' o = a * b mod 2^255 - 19. 2^256 is congruent to 38, so the high half
    ' folds back in with a factor of 38. Every partial sum stays below 2^53.
    Dim t(0 To 30) As Double
    Dim i As Long
    Dim j As Long
    Dim ai As Double
    For i = 0 To 15
        ai = a(i)
        For j = 0 To 15
            t(i + j) = t(i + j) + ai * b(j)
        Next j
    Next i
    For i = 0 To 14
        t(i) = t(i) + 38# * t(i + 16)
    Next i
    For i = 0 To 15
        o(i) = t(i)
    Next i
    EdCarry o
    EdCarry o
End Sub

Private Sub EdInvert(ByRef o() As Double, ByRef a() As Double)
    Dim c(0 To 15) As Double
    Dim k As Long
    For k = 0 To 15
        c(k) = a(k)
    Next k
    For k = 253 To 0 Step -1
        EdMul c, c, c
        If k <> 2 And k <> 4 Then EdMul c, c, a
    Next k
    For k = 0 To 15
        o(k) = c(k)
    Next k
End Sub

Private Sub EdPow2523(ByRef o() As Double, ByRef a() As Double)
    Dim c(0 To 15) As Double
    Dim k As Long
    For k = 0 To 15
        c(k) = a(k)
    Next k
    For k = 250 To 0 Step -1
        EdMul c, c, c
        If k <> 1 Then EdMul c, c, a
    Next k
    For k = 0 To 15
        o(k) = c(k)
    Next k
End Sub

Private Function EdLow16(ByVal v As Double) As Double
    EdLow16 = v - 65536# * Int(v / 65536#)
End Function

Private Function EdBorrow(ByVal v As Double) As Double
    If v < 0# Then EdBorrow = 1#
End Function

Private Sub EdPack(ByRef packedBytes() As Byte, ByRef n() As Double)
    ' Fully reduces n modulo 2^255 - 19 and writes 32 little-endian bytes.
    Dim t(0 To 15) As Double
    Dim m(0 To 15) As Double
    Dim i As Long
    Dim passNo As Long
    Dim borrowOut As Double
    Dim spare As Double
    For i = 0 To 15
        t(i) = n(i)
    Next i
    EdCarry t
    EdCarry t
    EdCarry t
    For passNo = 0 To 1
        m(0) = t(0) - 65517#
        For i = 1 To 14
            m(i) = t(i) - 65535# - EdBorrow(m(i - 1))
            m(i - 1) = EdLow16(m(i - 1))
        Next i
        m(15) = t(15) - 32767# - EdBorrow(m(14))
        borrowOut = EdBorrow(m(15))
        m(14) = EdLow16(m(14))
        If borrowOut = 0# Then
            For i = 0 To 15
                spare = t(i)
                t(i) = m(i)
                m(i) = spare
            Next i
        End If
    Next passNo
    For i = 0 To 15
        packedBytes(2 * i) = CByte(EdLow16(t(i)) - 256# * Int(EdLow16(t(i)) / 256#))
        packedBytes(2 * i + 1) = CByte(Int(EdLow16(t(i)) / 256#))
    Next i
End Sub

Private Sub EdUnpack(ByRef o() As Double, ByRef sourceBytes() As Byte)
    Dim i As Long
    For i = 0 To 15
        o(i) = CDbl(sourceBytes(2 * i)) + 256# * CDbl(sourceBytes(2 * i + 1))
    Next i
    If o(15) >= 32768# Then o(15) = o(15) - 32768#
End Sub

Private Function EdParity(ByRef a() As Double) As Long
    Dim packed(0 To 31) As Byte
    EdPack packed, a
    EdParity = packed(0) And 1
End Function

Private Function EdNotEqual(ByRef a() As Double, ByRef b() As Double) As Boolean
    Dim packedA(0 To 31) As Byte
    Dim packedB(0 To 31) As Byte
    Dim k As Long
    EdPack packedA, a
    EdPack packedB, b
    For k = 0 To 31
        If packedA(k) <> packedB(k) Then
            EdNotEqual = True
            Exit Function
        End If
    Next k
End Function

Private Sub EdPointAdd(ByRef p As EdPoint, ByRef q As EdPoint)
    Dim a(0 To 15) As Double
    Dim b(0 To 15) As Double
    Dim c(0 To 15) As Double
    Dim d(0 To 15) As Double
    Dim t(0 To 15) As Double
    Dim e(0 To 15) As Double
    Dim f(0 To 15) As Double
    Dim g(0 To 15) As Double
    Dim h(0 To 15) As Double
    EdSub a, p.Y, p.X
    EdSub t, q.Y, q.X
    EdMul a, a, t
    EdAdd b, p.X, p.Y
    EdAdd t, q.X, q.Y
    EdMul b, b, t
    EdMul c, p.T, q.T
    EdMul c, c, mEdD2
    EdMul d, p.Z, q.Z
    EdAdd d, d, d
    EdSub e, b, a
    EdSub f, d, c
    EdAdd g, d, c
    EdAdd h, b, a
    EdMul p.X, e, f
    EdMul p.Y, h, g
    EdMul p.Z, g, f
    EdMul p.T, e, h
End Sub

Private Sub EdBasePoint(ByRef r As EdPoint)
    Dim k As Long
    For k = 0 To 15
        r.X(k) = mEdBaseX(k)
        r.Y(k) = mEdBaseY(k)
        r.Z(k) = 0#
    Next k
    r.Z(0) = 1#
    EdMul r.T, mEdBaseX, mEdBaseY
End Sub

Private Sub EdScalarMult( _
    ByRef result As EdPoint, ByRef pointIn As EdPoint, _
    ByRef scalarBytes() As Byte, ByVal offset As Long)
    ' Double-and-add over the 256 scalar bits, most significant first.
    Dim p As EdPoint
    Dim q As EdPoint
    Dim spare As EdPoint
    Dim i As Long
    Dim bitValue As Long
    p.Y(0) = 1#
    p.Z(0) = 1#
    q = pointIn
    For i = 255 To 0 Step -1
        bitValue = (CLng(scalarBytes(offset + (i \ 8))) \ CLng(2 ^ (i And 7))) And 1
        If bitValue = 1 Then
            spare = p
            p = q
            q = spare
        End If
        EdPointAdd q, p
        EdPointAdd p, p
        If bitValue = 1 Then
            spare = p
            p = q
            q = spare
        End If
    Next i
    result = p
End Sub

Private Sub EdPackPoint(ByRef packedBytes() As Byte, ByRef p As EdPoint)
    Dim zInverse(0 To 15) As Double
    Dim tx(0 To 15) As Double
    Dim ty(0 To 15) As Double
    EdInvert zInverse, p.Z
    EdMul tx, p.X, zInverse
    EdMul ty, p.Y, zInverse
    EdPack packedBytes, ty
    packedBytes(31) = packedBytes(31) Xor CByte(EdParity(tx) * 128)
End Sub

Private Function EdUnpackNegated(ByRef r As EdPoint, ByRef keyBytes() As Byte) As Boolean
    ' Decodes the public key and returns -A (RFC 8032 section 5.1.3).
    Dim t(0 To 15) As Double
    Dim chk(0 To 15) As Double
    Dim num(0 To 15) As Double
    Dim den(0 To 15) As Double
    Dim den2(0 To 15) As Double
    Dim den4(0 To 15) As Double
    Dim den6(0 To 15) As Double
    Dim zero(0 To 15) As Double
    Dim k As Long
    Dim yIsCanonical As Boolean

    ' Reject y >= 2^255 - 19 (non-canonical encodings).
    yIsCanonical = (keyBytes(0) < 237) Or ((keyBytes(31) And 127) <> 127)
    If Not yIsCanonical Then
        For k = 1 To 30
            If keyBytes(k) <> 255 Then
                yIsCanonical = True
                Exit For
            End If
        Next k
    End If
    If Not yIsCanonical Then Exit Function

    For k = 0 To 15
        r.Z(k) = 0#
    Next k
    r.Z(0) = 1#
    EdUnpack r.Y, keyBytes
    EdMul num, r.Y, r.Y
    EdMul den, num, mEdD
    EdSub num, num, r.Z
    EdAdd den, r.Z, den
    EdMul den2, den, den
    EdMul den4, den2, den2
    EdMul den6, den4, den2
    EdMul t, den6, num
    EdMul t, t, den
    EdPow2523 t, t
    EdMul t, t, num
    EdMul t, t, den
    EdMul t, t, den
    EdMul r.X, t, den
    EdMul chk, r.X, r.X
    EdMul chk, chk, den
    If EdNotEqual(chk, num) Then EdMul r.X, r.X, mEdSqrtMinusOne
    EdMul chk, r.X, r.X
    EdMul chk, chk, den
    If EdNotEqual(chk, num) Then Exit Function
    If EdParity(r.X) = (keyBytes(31) \ 128) Then EdSub r.X, zero, r.X
    EdMul r.T, r.X, r.Y
    EdUnpackNegated = True
End Function

Private Sub EdReduceScalar(ByRef hashBytes() As Byte, ByRef result() As Byte)
    ' Reduces the 64-byte SHA-512 output modulo the group order L.
    Dim x(0 To 64) As Double
    Dim i As Long
    Dim j As Long
    Dim carry As Double
    For i = 0 To 63
        x(i) = hashBytes(i)
    Next i
    For i = 63 To 32 Step -1
        carry = 0#
        For j = i - 32 To i - 13
            x(j) = x(j) + carry - 16# * x(i) * mEdOrder(j - (i - 32))
            carry = Int((x(j) + 128#) / 256#)
            x(j) = x(j) - carry * 256#
        Next j
        x(i - 12) = x(i - 12) + carry
        x(i) = 0#
    Next i
    carry = 0#
    For j = 0 To 31
        x(j) = x(j) + carry - Int(x(31) / 16#) * mEdOrder(j)
        carry = Int(x(j) / 256#)
        x(j) = x(j) - carry * 256#
    Next j
    For j = 0 To 31
        x(j) = x(j) - carry * mEdOrder(j)
    Next j
    For i = 0 To 31
        x(i + 1) = x(i + 1) + Int(x(i) / 256#)
        result(i) = CByte(x(i) - 256# * Int(x(i) / 256#))
    Next i
End Sub

Private Function JsonSimpleString( _
    ByVal json As String, ByVal fieldName As String) As String
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & fieldName & """\s*:\s*""([^""\\]*)"""
    re.Global = False
    If re.Test(json) Then
        Set matches = re.Execute(json)
        JsonSimpleString = matches(0).SubMatches(0)
    End If
End Function

Private Function JsonNumber( _
    ByVal json As String, ByVal fieldName As String) As Double
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & fieldName & """\s*:\s*([0-9]+)(\s*[,}])"
    re.Global = False
    If re.Test(json) Then
        Set matches = re.Execute(json)
        JsonNumber = CDbl(matches(0).SubMatches(0))
    End If
End Function

Private Function JsonArraySlice( _
    ByVal json As String, ByVal fieldName As String) As String
    Dim re As Object
    Dim matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & fieldName & """\s*:\s*(\[[^\[\]]*\])"
    re.Global = False
    If re.Test(json) Then
        Set matches = re.Execute(json)
        JsonArraySlice = matches(0).SubMatches(0)
    End If
End Function

Private Function IsSafeIdentifier(ByVal value As String) As Boolean
    Dim re As Object
    If Len(value) = 0 Or Len(value) > 128 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_.-]+$"
    IsSafeIdentifier = re.Test(value)
End Function

Private Function IsJwtCharacters(ByVal value As String) As Boolean
    Dim re As Object
    If Len(value) = 0 Or Len(value) > 8192 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$"
    IsJwtCharacters = re.Test(value)
End Function

Private Function IsHex64(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9a-f]{64}$"
    IsHex64 = re.Test(value)
End Function

Private Function IsDeviceCode(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_-]{24,256}$"
    IsDeviceCode = re.Test(value)
End Function

Private Function IsSemVer(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]+\.[0-9]+\.[0-9]+$"
    IsSemVer = re.Test(value)
End Function

Private Function IsUserCode(ByVal value As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Z0-9-]{6,16}$"
    IsUserCode = re.Test(value)
End Function

Private Function IsHttpsUrl(ByVal value As String) As Boolean
    Dim re As Object
    If Len(value) = 0 Or Len(value) > 512 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^https://[A-Za-z0-9.-]+(:[0-9]{1,5})?(/[A-Za-z0-9/_-]*)?$"
    IsHttpsUrl = re.Test(value)
End Function

Private Function NewRequestId() As String
    Dim g As GUID_VALUE
    Dim i As Long
    Dim lastPart As String
    If CoCreateGuid(g) <> 0 Then Err.Raise 5
    For i = 2 To 7
        lastPart = lastPart & Right$("0" & Hex$(g.Data4(i)), 2)
    Next i
    NewRequestId = LCase$( _
        Right$("00000000" & Hex$(g.Data1), 8) & "-" & _
        Right$("0000" & Hex$(g.Data2), 4) & "-" & _
        Right$("0000" & Hex$(g.Data3), 4) & "-" & _
        Right$("0" & Hex$(g.Data4(0)), 2) & _
        Right$("0" & Hex$(g.Data4(1)), 2) & "-" & lastPart)
End Function

Private Function UtcUnixNow() As Double
    Dim value As SYSTEMTIME
    GetSystemTime value
    UtcUnixNow = CDbl(DateDiff("d", DateSerial(1970, 1, 1), _
        DateSerial(value.wYear, value.wMonth, value.wDay))) * 86400# + _
        CDbl(value.wHour) * 3600# + CDbl(value.wMinute) * 60# + _
        CDbl(value.wSecond)
End Function

Private Function ByteCount(ByRef values() As Byte) As Long
    On Error GoTo EmptyArray
    ByteCount = UBound(values) - LBound(values) + 1
    Exit Function
EmptyArray:
    ByteCount = 0
End Function

Private Function Utf8Bytes(ByVal value As String) As Byte()
    Dim count As Long
    Dim result() As Byte
    If Len(value) = 0 Then Exit Function
    count = WideCharToMultiByte(CP_UTF8, 0, StrPtr(value), _
                                Len(value), 0, 0, 0, 0)
    If count <= 0 Then Err.Raise 5
    ReDim result(0 To count - 1)
    If WideCharToMultiByte(CP_UTF8, 0, StrPtr(value), Len(value), _
                           VarPtr(result(0)), count, 0, 0) <> count Then Err.Raise 5
    Utf8Bytes = result
End Function

Private Function Utf8Text(ByRef bytes() As Byte) As String
    Dim count As Long
    Dim chars As Long
    Dim result As String
    count = ByteCount(bytes)
    If count = 0 Then Exit Function
    chars = MultiByteToWideChar(CP_UTF8, 8, VarPtr(bytes(0)), count, 0, 0)
    If chars <= 0 Then Err.Raise 5
    result = String$(chars, vbNullChar)
    If MultiByteToWideChar(CP_UTF8, 8, VarPtr(bytes(0)), count, _
                           StrPtr(result), chars) <> chars Then Err.Raise 5
    Utf8Text = result
End Function

Private Function Base64UrlBytes(ByVal value As String) As Byte()
    Dim normalized As String
    If Len(value) = 0 Then Err.Raise 5
    normalized = Replace(Replace(value, "-", "+"), "_", "/")
    If Len(normalized) Mod 4 = 1 Then Err.Raise 5
    normalized = normalized & String$((4 - Len(normalized) Mod 4) Mod 4, "=")
    Base64UrlBytes = Base64Bytes(normalized)
End Function

Private Function Base64Bytes(ByVal value As String) As Byte()
    Dim count As Long
    Dim result() As Byte
    If Len(value) = 0 Then Err.Raise 5
    If CryptStringToBinaryW(StrPtr(value), Len(value), _
                            CRYPT_STRING_BASE64, 0, count, 0, 0) = 0 Then Err.Raise 5
    If count <= 0 Then Err.Raise 5
    ReDim result(0 To count - 1)
    If CryptStringToBinaryW(StrPtr(value), Len(value), _
                            CRYPT_STRING_BASE64, VarPtr(result(0)), _
                            count, 0, 0) = 0 Then Err.Raise 5
    Base64Bytes = result
End Function

Private Function BytesToHex(ByRef values() As Byte) As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        BytesToHex = BytesToHex & LCase$(Right$("0" & Hex$(values(i)), 2))
    Next i
End Function

Private Function HexToBytes(ByVal value As String) As Byte()
    Dim i As Long
    Dim result() As Byte
    If Len(value) = 0 Or Len(value) Mod 2 <> 0 Then Err.Raise 5
    ReDim result(0 To Len(value) \ 2 - 1)
    For i = 0 To UBound(result)
        result(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2))
    Next i
    HexToBytes = result
End Function

Private Function Sha256Hex(ByVal value As String) As String
    Dim algorithm As LongPtr
    Dim sourceData() As Byte
    Dim digest(0 To 31) As Byte
    On Error GoTo Failed
    sourceData = Utf8Bytes(value)
    If ByteCount(sourceData) = 0 Then Exit Function
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr("SHA256"), 0, 0) <> 0 Then GoTo Failed
    If BCryptHash(algorithm, 0, 0, VarPtr(sourceData(0)), _
                  ByteCount(sourceData), VarPtr(digest(0)), 32) <> 0 Then GoTo Failed
    Sha256Hex = BytesToHex(digest)
    BCryptCloseAlgorithmProvider algorithm, 0
    Exit Function
Failed:
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    Sha256Hex = vbNullString
End Function

Public Function DeviceFingerprintHash() As String
    Dim parts(0 To 4) As String
    On Error GoTo Failed
    parts(0) = RegistryValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography\MachineGuid")
    parts(1) = CStr(CreateObject("Scripting.FileSystemObject").GetDrive(SystemDriveName(Environ$("SystemDrive"))).SerialNumber)
    parts(2) = FirstWmiValue("Win32_Processor", "ProcessorId")
    parts(3) = RegistryValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductId")
    parts(4) = CurrentUserSid()
    If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Or _
       Len(parts(2)) = 0 Or Len(parts(3)) = 0 Or _
       Len(parts(4)) = 0 Then Exit Function
    DeviceFingerprintHash = Sha256Hex(Join(parts, "|"))
    Exit Function
Failed:
    DeviceFingerprintHash = vbNullString
End Function

Private Function SystemDriveName(ByVal configuredDrive As String) As String
    Dim driveName As String
    driveName = UCase$(Trim$(configuredDrive))
    If Len(driveName) = 3 And Right$(driveName, 1) = "\" Then driveName = Left$(driveName, 2)
    If Len(driveName) <> 2 Then Err.Raise 5
    If Mid$(driveName, 2, 1) <> ":" Or Left$(driveName, 1) < "A" Or Left$(driveName, 1) > "Z" Then Err.Raise 5
    SystemDriveName = driveName
End Function

Private Function RegistryValue(ByVal keyName As String) As String
    RegistryValue = CStr(CreateObject("WScript.Shell").RegRead(keyName))
End Function

Private Function FirstWmiValue( _
    ByVal className As String, ByVal propertyName As String) As String
    Dim rows As Object
    Dim row As Object
    Set rows = GetObject("winmgmts:\\.\root\cimv2").ExecQuery( _
        "SELECT " & propertyName & " FROM " & className)
    For Each row In rows
        FirstWmiValue = CStr(CallByName(row, propertyName, VbGet))
        Exit Function
    Next row
End Function

Private Function CurrentUserSid() As String
    Dim account As Object
    Dim network As Object
    Set network = CreateObject("WScript.Network")
    Set account = GetObject("winmgmts:\\.\root\cimv2:Win32_UserAccount.Name='" & _
        Replace(network.UserName, "'", "''") & "',Domain='" & _
        Replace(network.UserDomain, "'", "''") & "'")
    CurrentUserSid = CStr(account.SID)
End Function

Private Function CacheClear() As Boolean
    On Error GoTo Failed
    If Len(GetSetting(CACHE_APP, CacheSection(), "Ciphertext", vbNullString)) > 0 Then
        DeleteSetting CACHE_APP, CacheSection()
    End If
    CacheClear = True
    Exit Function
Failed:
    CacheClear = False
End Function

Private Function CacheSave( _
    ByVal token As String, ByVal fingerprint As String, _
    ByVal validatedUnix As Double, ByVal maxSeenUnix As Double) As Boolean

    Dim plain() As Byte
    Dim cipher As String
    On Error GoTo Failed
    If Not IsJwtCharacters(token) Or Not IsHex64(fingerprint) Then Exit Function
    maxSeenUnix = MaxDouble(maxSeenUnix, mMaxSeenUnix)
    If validatedUnix <= 0 Or maxSeenUnix < validatedUnix Then Exit Function
    mMaxSeenUnix = maxSeenUnix
    plain = Utf8Bytes(token & vbLf & fingerprint & vbLf & _
                      Format$(validatedUnix, "0") & vbLf & Format$(maxSeenUnix, "0"))
    cipher = DpapiProtectHex(plain)
    If Len(cipher) = 0 Then Exit Function
    SaveSetting CACHE_APP, CacheSection(), "Ciphertext", cipher
    mSavedMaxSeenUnix = maxSeenUnix
    CacheSave = True
    Exit Function
Failed:
    CacheSave = False
End Function

Private Function CacheLoad( _
    ByRef token As String, ByRef fingerprint As String, _
    ByRef validatedUnix As Double, ByRef maxSeenUnix As Double) As Boolean

    Dim cipher As String
    Dim plain() As Byte
    Dim fields() As String
    On Error GoTo Failed
    token = vbNullString
    fingerprint = vbNullString
    validatedUnix = 0
    maxSeenUnix = 0
    cipher = GetSetting(CACHE_APP, CacheSection(), "Ciphertext", vbNullString)
    If Len(cipher) = 0 Then Exit Function
    plain = DpapiUnprotectHex(cipher)
    If ByteCount(plain) = 0 Then Exit Function
    fields = Split(Utf8Text(plain), vbLf)
    If UBound(fields) <> 3 Then Exit Function
    If Not IsJwtCharacters(fields(0)) Then Exit Function
    If Not IsHex64(fields(1)) Then Exit Function
    validatedUnix = Val(fields(2))
    maxSeenUnix = Val(fields(3))
    If validatedUnix <= 0 Or maxSeenUnix < validatedUnix Then Exit Function
    token = fields(0)
    fingerprint = fields(1)
    CacheLoad = True
    Exit Function
Failed:
    token = vbNullString
    fingerprint = vbNullString
    validatedUnix = 0
    maxSeenUnix = 0
    CacheLoad = False
End Function

Private Function CacheEntropyLabel() As String
    ' Public, product-specific separation label; this is not a secret.
    CacheEntropyLabel = "DigitalLicenseVbaV1|cache-v2|" & Cfg_ProductKey()
End Function

Private Function DpapiProtectHex(ByRef plain() As Byte) As String
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    Dim resultData() As Byte
    Dim entropyBytes() As Byte
    Dim entropyBlob As DATA_BLOB
    On Error GoTo Cleanup
    inputBlob.cbData = ByteCount(plain)
    If inputBlob.cbData = 0 Then Exit Function
    inputBlob.pbData = VarPtr(plain(0))
    entropyBytes = Utf8Bytes(CacheEntropyLabel())
    entropyBlob.cbData = ByteCount(entropyBytes)
    entropyBlob.pbData = VarPtr(entropyBytes(0))
    If CryptProtectData(inputBlob, 0, VarPtr(entropyBlob), 0, 0, _
                        CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup
    If outputBlob.cbData <= 0 Then GoTo Cleanup
    ReDim resultData(0 To outputBlob.cbData - 1)
    CopyMemory VarPtr(resultData(0)), outputBlob.pbData, outputBlob.cbData
    DpapiProtectHex = BytesToHex(resultData)
Cleanup:
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
End Function

Private Function DpapiUnprotectHex(ByVal cipher As String) As Byte()
    Dim sourceData() As Byte
    Dim resultData() As Byte
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    Dim entropyBytes() As Byte
    Dim entropyBlob As DATA_BLOB
    On Error GoTo Cleanup
    sourceData = HexToBytes(cipher)
    inputBlob.cbData = ByteCount(sourceData)
    inputBlob.pbData = VarPtr(sourceData(0))
    entropyBytes = Utf8Bytes(CacheEntropyLabel())
    entropyBlob.cbData = ByteCount(entropyBytes)
    entropyBlob.pbData = VarPtr(entropyBytes(0))
    If CryptUnprotectData(inputBlob, 0, VarPtr(entropyBlob), 0, 0, _
                          CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup
    If outputBlob.cbData <= 0 Then GoTo Cleanup
    ReDim resultData(0 To outputBlob.cbData - 1)
    CopyMemory VarPtr(resultData(0)), outputBlob.pbData, outputBlob.cbData
    DpapiUnprotectHex = resultData
Cleanup:
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
End Function

Public Function License_StatusCode() As String
    If Not mChecked Then License_OnOpen
    License_StatusCode = mErrorCode
    If Len(License_StatusCode) = 0 Then
        If License_PremiumAllowed() Then
            License_StatusCode = "ACTIVE"
        Else
            License_StatusCode = "REVALIDATION_REQUIRED"
        End If
    End If
End Function

Public Sub License_ShowGuidance()
    Dim guidance As String
    Select Case License_StatusCode()
        Case "CONFIG_REQUIRED"
            guidance = "Can cau hinh URL HTTPS, product key, key id va khoa cong khai."
        Case "ACTIVATION_REQUIRED"
            guidance = "Hay dang nhap va kich hoat giay phep theo huong dan cua nha phat hanh."
        Case "NETWORK_UNAVAILABLE"
            guidance = "Khong ket noi duoc may chu cap phep. Hay thu lai khi co mang."
        Case "ACTIVATION_CANCELLED"
            guidance = "Ban da huy kich hoat tren may nay."
        Case "ACTIVATION_EXPIRED"
            guidance = "Ma kich hoat da het han. Hay mo lai file de tao ma moi."
        Case "DEVICE_LIMIT_EXCEEDED"
            guidance = "Giay phep da het suat thiet bi. Hay tat may cu trong tai khoan."
        Case "LICENSE_NOT_FOUND"
            guidance = "Tai khoan chua co giay phep dang hoat dong cho san pham nay."
        Case "LICENSE_EXPIRED"
            guidance = "Giay phep da het han. Hay gia han trong tai khoan."
        Case "LICENSE_SUSPENDED", "LICENSE_REVOKED"
            guidance = "Giay phep khong con duoc phep su dung. Hay lien he nha phat hanh."
        Case "DEVICE_REVOKED", "DEVICE_MISMATCH"
            guidance = "May da bi go hoac thong tin may da thay doi. Ban co the kich hoat lai bang ma moi." & _
                       vbCrLf & "Neu het suat thiet bi, hay go may cu trong tai khoan."
        Case "RATE_LIMITED"
            guidance = "Da thu qua nhieu lan. Hay doi mot luc roi thu lai."
        Case "DEVICE_CODE_EXPIRED"
            guidance = "Ma kich hoat da het han. Hay mo lai file de tao ma moi."
        Case "DEVICE_CODE_USED"
            guidance = "Ma kich hoat nay da duoc dung. Hay mo lai file de tao ma moi."
        Case "DEVICE_CODE_DENIED"
            guidance = "Yeu cau kich hoat da bi tu choi tren trang web. Hay mo lai file neu muon thu lai."
        Case "DEVICE_CODE_INVALID"
            guidance = "Ma kich hoat khong hop le. Hay mo lai file de tao ma moi."
        Case "PRODUCT_VERSION_UNSUPPORTED"
            guidance = "Phien ban file nay khong con duoc ho tro. Hay tai phien ban moi trong tai khoan."
        Case "TOKEN_INVALID"
            guidance = "Du lieu giay phep khong hop le. Ban co the kich hoat lai bang ma moi." & _
                       vbCrLf & "Neu het suat thiet bi, hay go may cu trong tai khoan."
        Case "CACHE_UNAVAILABLE"
            guidance = "Khong the cap nhat hoac xoa cache giay phep. Hay lien he nha phat hanh."
        Case "CLOCK_ROLLBACK"
            guidance = "Dong ho may tinh da lui so voi lan xac minh truoc. Hay chinh dung ngay gio roi mo lai file."
        Case "SHEET_CONFIG_INVALID"
            guidance = "Can cau hinh dung sheet noi dung va giu mot sheet huong dan hien thi. Hay lien he nha phat hanh."
        Case "REQUEST_TIMESTAMP_INVALID"
            guidance = "Dong ho may tinh bi lech. Hay chinh dung ngay gio roi thu lai."
        Case "REVALIDATION_REQUIRED"
            guidance = "Can ket noi mang de xac minh lai giay phep."
        Case Else
            guidance = "Giay phep chua duoc xac minh. Hay lien he nha phat hanh."
    End Select
    MsgBox guidance & vbCrLf & "Nhan OK hoac dong hop thoai de dong workbook.", _
           vbInformation, "Trang thai giay phep"
End Sub

Public Sub License_RunMockTests()
    Dim sampleFp As String
    Dim sampleId As String
    Dim body As String
    Dim countPassed As Long
    Dim sampleMessage() As Byte
    Dim sampleSignature() As Byte
    Dim sampleKey() As Byte
    Dim testKeyIds As Variant
    Dim testPublicKeys As Variant
    sampleFp = String$(64, "a")
    sampleId = NewRequestId()
    If Len(sampleId) = 36 And Mid$(sampleId, 9, 1) = "-" Then countPassed = countPassed + 1
    body = BuildActivateJson("sample_product", "1.2.3", sampleFp, sampleId, 1234567890#)
    If InStr(body, """product_key""") > 0 And _
       InStr(body, """version"":""1.2.3""") > 0 And _
       InStr(body, """fp_hash""") > 0 And _
       InStr(body, """timestamp"":1234567890") > 0 Then countPassed = countPassed + 1
    If PolicyCanUse(110#, 100#, 200#, 20#) Then countPassed = countPassed + 1
    If Not PolicyCanUse(130#, 100#, 200#, 20#) Then countPassed = countPassed + 1
    sampleMessage = Utf8Bytes("test")
    sampleSignature = Utf8Bytes("invalid")
    sampleKey = Utf8Bytes("invalid")
    If Not VerifyEd25519Detached(sampleMessage, sampleSignature, _
                                  sampleKey) Then countPassed = countPassed + 1
    If IsDeviceCode(String$(32, "A")) And _
       Not IsDeviceCode("short") Then countPassed = countPassed + 1
    body = BuildDeviceCodeRedeemJson(String$(32, "A"), sampleFp, _
                                     sampleId, 1234567890#)
    If InStr(body, """device_code""") > 0 And _
       InStr(body, """fp_hash""") > 0 And _
       InStr(body, "device_name") = 0 And _
       InStr(BuildDeviceCodeRedeemJson(String$(32, "A"), sampleFp, sampleId, 1234567890#, _
             "LAPTOP-GIANG"), ",""device_name"":""LAPTOP-GIANG""}") > 0 And _
       SafeDeviceName("Ma""y {1}" & vbTab & String$(70, "x")) = "May1" & String$(59, "x") _
       Then countPassed = countPassed + 1
    If Ed25519KnownAnswerTest() Then countPassed = countPassed + 1
    If Not PolicyCanUse(200#, 100#, 200#, 3600#) Then countPassed = countPassed + 1
    If Not PolicyCanUse(100#, 100#, 200#, 0#) Then countPassed = countPassed + 1
    If PolicyRequiresOnline(200#, 100#, 200#, 1000#, 1#) Then countPassed = countPassed + 1
    If PolicyRequiresOnline(110#, 100#, 200#, 20#, 0#) Then countPassed = countPassed + 1
    If Not PolicyRequiresOnline(110#, 100#, 200#, 20#, 1#) Then countPassed = countPassed + 1
    If PolicyRequiresOnline(120#, 100#, 200#, 20#, 1#) Then countPassed = countPassed + 1
    If PolicyRequiresOnline(110#, 100#, 200#, 20#, 0.001) Then countPassed = countPassed + 1
    If NeedsDeviceActivation("ACTIVATION_REQUIRED", False) And _
       NeedsDeviceActivation("TOKEN_INVALID", True) And _
       Not NeedsDeviceActivation("REVALIDATION_REQUIRED", True) And _
       Not NeedsDeviceActivation("NETWORK_UNAVAILABLE", True) Then countPassed = countPassed + 1
    If Not NeedsDeviceActivation("TOKEN_INVALID", False) Then countPassed = countPassed + 1
    If NeedsDeviceActivation("DEVICE_REVOKED", True) Then countPassed = countPassed + 1
    If NeedsDeviceActivation("DEVICE_MISMATCH", True) Then countPassed = countPassed + 1
    If Not NeedsDeviceActivation("DEVICE_REVOKED", False) Then countPassed = countPassed + 1
    If Not NeedsDeviceActivation("DEVICE_MISMATCH", False) Then countPassed = countPassed + 1
    If Not NeedsDeviceActivation("RATE_LIMITED", True) Then countPassed = countPassed + 1
    If ShouldClearCacheForServerError("TOKEN_INVALID", 401) Then countPassed = countPassed + 1
    If ShouldClearCacheForServerError("DEVICE_REVOKED", 403) Then countPassed = countPassed + 1
    If ShouldClearCacheForServerError("DEVICE_MISMATCH", 403) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("NETWORK_UNAVAILABLE", 0) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("RATE_LIMITED", 429) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("REQUEST_TIMESTAMP_INVALID", 400) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("TOKEN_INVALID", 500) And _
       Not ShouldClearCacheForServerError("DEVICE_REVOKED", 503) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("LICENSE_EXPIRED", 403) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("LICENSE_SUSPENDED", 403) Then countPassed = countPassed + 1
    If Not ShouldClearCacheForServerError("TOKEN_INVALID", 429) Then countPassed = countPassed + 1
    If ClockCanUse(110#, 100#, 105#) Then countPassed = countPassed + 1
    If Not ClockCanUse(99#, 100#, 90#) Then countPassed = countPassed + 1
    If Not ClockCanUse(100#, 90#, 101#) Then countPassed = countPassed + 1
    If ClockCanUse(100#, 100#, 100#) Then countPassed = countPassed + 1
    If Not ClockCanUse(0#, 0#, 0#) Then countPassed = countPassed + 1
    If MaxDouble(100#, 90#) = 100# Then countPassed = countPassed + 1
    If RevalidationCycle(110#, 100#, 20#) = 0 Then countPassed = countPassed + 1
    If RevalidationCycle(120#, 100#, 20#) = 1 Then countPassed = countPassed + 1
    If RevalidationCycle(140#, 100#, 20#) = 2 Then countPassed = countPassed + 1
    If RevalidationCycle(120#, 100#, 0#) = -1 Then countPassed = countPassed + 1
    If RevalidationCycle(90#, 100#, 20#) = -1 Then countPassed = countPassed + 1
    If CanAttemptRefresh(1#, -1#, 100#, 100#, False) Then countPassed = countPassed + 1
    If Not CanAttemptRefresh(1#, 1#, 100#, 100#, False) Then countPassed = countPassed + 1
    If CanAttemptRefresh(2#, 1#, 100#, 100#, False) Then countPassed = countPassed + 1
    If Not CanAttemptRefresh(2#, 1#, 100#, 100#, True) Then countPassed = countPassed + 1
    If Not CanAttemptRefresh(-1#, -1#, 100#, 100#, False) Then countPassed = countPassed + 1
    testKeyIds = Cfg_ActivationKeyIds()
    testPublicKeys = Cfg_ActivationPublicKeysBase64()
    If PublicKeyForKid(CStr(testKeyIds(0))) = CStr(testPublicKeys(0)) Then countPassed = countPassed + 1
    If Len(PublicKeyForKid("unknown-mock-key")) = 0 Then countPassed = countPassed + 1
    If Len(FindPublicKey(vbNullString, Array("one", ""), Array(testPublicKeys(0), ""))) = 0 Then countPassed = countPassed + 1
    If KeyPairsValid(Array("one", ""), Array(testPublicKeys(0), "")) Then countPassed = countPassed + 1
    If Not KeyPairsValid(Array("one", "one"), Array(testPublicKeys(0), testPublicKeys(0))) Then countPassed = countPassed + 1
    If Not KeyPairsValid(Array("one", "two"), Array(testPublicKeys(0), "")) Then countPassed = countPassed + 1
    If Not KeyPairsValid(Array("one", ""), Array("YQ==", "")) Then countPassed = countPassed + 1
    If TokenKeyIdsMatch("one", "one") Then countPassed = countPassed + 1
    If Not TokenKeyIdsMatch("one", "two") Then countPassed = countPassed + 1
    If DpapiRoundTripTest() Then countPassed = countPassed + 1
    If DpapiLegacyRejectTest() Then countPassed = countPassed + 1
    If CacheEntropyLabel() = "DigitalLicenseVbaV1|cache-v2|" & Cfg_ProductKey() Then countPassed = countPassed + 1
    If SystemDriveName("C:") = "C:" Then countPassed = countPassed + 1
    If SystemDriveName("C:\") = "C:" Then countPassed = countPassed + 1
    If SystemDriveName("D:") = "D:" Then countPassed = countPassed + 1
    If ArrayItemCount(Array()) = 0 Then countPassed = countPassed + 1
    If countPassed = 64 Then
        MsgBox "64 phep thu gia lap dat. Khong goi API va khong cap quyen.", vbInformation
    Else
        MsgBox "Phep thu gia lap that bai: " & CStr(countPassed) & "/64.", vbExclamation
    End If
End Sub

Private Function DpapiRoundTripTest() As Boolean
    Dim plain() As Byte
    Dim recovered() As Byte
    Dim cipher As String
    On Error GoTo Failed
    plain = Utf8Bytes("dpapi-round-trip-only")
    cipher = DpapiProtectHex(plain)
    If Len(cipher) = 0 Then Exit Function
    recovered = DpapiUnprotectHex(cipher)
    DpapiRoundTripTest = (Utf8Text(recovered) = "dpapi-round-trip-only")
Failed:
End Function

Private Function DpapiLegacyRejectTest() As Boolean
    Dim plain() As Byte
    Dim legacyCipher() As Byte
    Dim recovered() As Byte
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    On Error GoTo Cleanup
    plain = Utf8Bytes("legacy-cache-test-only")
    inputBlob.cbData = ByteCount(plain)
    inputBlob.pbData = VarPtr(plain(0))
    If CryptProtectData(inputBlob, 0, 0, 0, 0, CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup
    If outputBlob.cbData <= 0 Then GoTo Cleanup
    ReDim legacyCipher(0 To outputBlob.cbData - 1)
    CopyMemory VarPtr(legacyCipher(0)), outputBlob.pbData, outputBlob.cbData
    recovered = DpapiUnprotectHex(BytesToHex(legacyCipher))
    DpapiLegacyRejectTest = (ByteCount(recovered) = 0)
Cleanup:
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
End Function

Private Function Ed25519KnownAnswerTest() As Boolean
    ' RFC 8032 section 7.1 test vector 1 (empty message): the genuine
    ' signature must verify and a one-bit change must not.
    Dim emptyMessage() As Byte
    Dim signature() As Byte
    Dim publicKey() As Byte
    On Error GoTo Failed
    publicKey = HexToBytes("d75a980182b10ab7d54bfed3c964073a0ee172f3daa62325af021a68f707511a")
    signature = HexToBytes("e5564300c360ac729086e2cc806e828a84877f1eb8e5d974d873e06522490155" & _
                           "5fb8821590a33bacc61e39701cf9b46bd25bf5f0595bbe24655141438e7a100b")
    If Not VerifyEd25519Detached(emptyMessage, signature, publicKey) Then Exit Function
    signature(0) = signature(0) Xor 1
    If VerifyEd25519Detached(emptyMessage, signature, publicKey) Then Exit Function
    Ed25519KnownAnswerTest = True
    Exit Function
Failed:
    Ed25519KnownAnswerTest = False
End Function
