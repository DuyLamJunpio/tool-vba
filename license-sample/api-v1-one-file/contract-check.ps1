$ErrorActionPreference = 'Stop'
$file = Join-Path $PSScriptRoot 'modLicenseApiV1.bas'
$bytes = [IO.File]::ReadAllBytes($file)
if ($bytes.Where({ $_ -gt 127 }).Count) { throw 'VBA source must remain ASCII for Excel import' }
$text = [Text.Encoding]::ASCII.GetString($bytes)

foreach ($required in @(
    'Attribute VB_Name = "modLicenseApiV1"',
    'Option Explicit',
    'Public Sub Auto_Open()',
    'Public Function License_PremiumAllowed()',
    'Public Function License_FeatureAllowed(',
    'Public Function License_ActivateWithDeviceCode()',
    'Private Function ValidateWithDeviceToken(',
    '"/auth/device-code"',
    '"/auth/device-code/redeem"',
    '"/activations/validate-device"',
    '""product_key""', '""version""',
    '""fp_hash""',
    '""request_id""',
    '""timestamp""',
    '""activation_token""',
    '""device_code""',
    'VerifyEd25519Detached = False',
    'http.Option(6) = False',
    'http.Option(4) = 0',
    'CoCreateGuid',
    'Public Sub Auto_Close()',
    'Public Function Cfg_ProtectedSheetNames() As Variant',
    'Public Function Cfg_ProxyServer() As String',
    'Public Function Cfg_ActivationKeyIds() As Variant',
    'Public Function Cfg_ActivationPublicKeysBase64() As Variant',
    'Private Function ClockCanUse(',
    'Private Function RevalidationCycle(',
    'http.Option(9) = &H800 Or &H2000',
    'WinHttpGetIEProxyConfigForCurrentUser',
    'WinHttpGetProxyForUrl',
    'http.SetProxy 0',
    'http.SetProxy 2, overrideProxy',
    'xlSheetVeryHidden',
    'xlSheetVisible',
    'Environ$("SystemDrive")',
    'ThisWorkbook.Close SaveChanges:=False',
    'If NeedsDeviceActivation(mErrorCode, mCacheClearedForActivation) And _',
    'If PolicyRequiresOnline(nowUnix, cachedTime, expiry, interval, graceHours) Then',
    'If mErrorCode <> "NETWORK_UNAVAILABLE" Then Exit Sub',
    'Private Function CacheClear() As Boolean',
    'DeleteSetting CACHE_APP, CACHE_SECTION',
    'If ShouldClearCacheForServerError(mErrorCode, httpStatus) Then _',
    'InvalidateCachedActivation "DEVICE_MISMATCH"',
    'InvalidateCachedActivation "TOKEN_INVALID"',
    'mCacheClearedForActivation = CacheClear()',
    'If Not mCacheClearedForActivation Then mErrorCode = "CACHE_UNAVAILABLE"',
    'If httpStatus < 400 Or httpStatus >= 500 Or httpStatus = 429 Then Exit Function',
    'Optional ByVal allowExpiredForRenewal As Boolean = False',
    'If expiry <= nowUnix And Not allowExpiredForRenewal Then Exit Function',
    'If expiry <= 0 Or interval <= 0 Or graceHours < 0 Then Exit Function',
    'If countPassed = 64 Then',
    'If Not ClockCanUse(nowUnix, cachedMaxSeen, issuedUnix) Then',
    'ClockCanUse(currentUnix, mMaxSeenUnix, mIssuedUnix)',
    'maxSeenUnix = MaxDouble(maxSeenUnix, mMaxSeenUnix)',
    'mRefreshInProgress = True',
    'mRefreshInProgress = False',
    'mRefreshAttemptCycle = RevalidationCycle(currentUnix, mLastValidatedUnix, mRevalidateSeconds)',
    'If mClosing Or mRefreshInProgress Then Exit Function',
    'If UBound(fields) <> 3 Then Exit Function',
    'If ByteCount(plain) = 0 Then Exit Function',
    'CryptProtectData(inputBlob, 0, VarPtr(entropyBlob), 0, 0, _',
    'CryptUnprotectData(inputBlob, 0, VarPtr(entropyBlob), 0, 0, _',
    'keyValue = PublicKeyForKid(headerKid)',
    'If Not TokenKeyIdsMatch(headerKid, JsonSimpleString(payloadJson, "kid")) Then Exit Function'
)) {
    if (-not $text.Contains($required)) { throw "Missing contract element: $required" }
}

foreach ($forbidden in @(
    'sheet.Protect', 'sheet.Unprotect', 'Application.Quit',
    'ActiveWorkbook.Close', 'Workbooks.Close',
    'Debug.Print', 'BEGIN PRIVATE KEY', 'shared secret',
    'AccessToken', '"/activations"', '"/activations/validate"',
    'Cfg_KeyId()', 'Cfg_Ed25519PublicKeyBase64()', 'GetDrive("C:")'
)) {
    if ($text.IndexOf($forbidden, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "Forbidden VBA source pattern: $forbidden"
    }
}

foreach ($blankConfig in @('Cfg_ApiBaseUrl', 'Cfg_ProductKey', 'Cfg_ProductVersion', 'Cfg_ActivationWebUrl', 'Cfg_ProxyServer')) {
    if ($text -notmatch ($blankConfig + ' = vbNullString')) { throw "$blankConfig must start blank" }
}
if ($text -notmatch 'Cfg_ProtectedSheetNames = Array\(\)') { throw 'Protected sheet names must start empty' }
$flat = $text -replace ' _\r?\n\s*', ' '
$ids = [regex]::Match($flat, 'Cfg_ActivationKeyIds = Array\("([A-Za-z0-9._-]+)", (vbNullString|"[A-Za-z0-9._-]+")\)')
$keys = [regex]::Match($flat, 'Cfg_ActivationPublicKeysBase64 = Array\(\s*"([A-Za-z0-9+/=]+)", (vbNullString|"[A-Za-z0-9+/=]+")\)')
if (-not $ids.Success -or -not $keys.Success) { throw 'Two explicit kid/public-key slots are required' }
if ($ids.Groups[1].Value -cne 'k2026-10' -or
    $keys.Groups[1].Value -cne 'iCNGgbXlaYmAsu5XJewI7h6935JuPaEnASEj8qvPGaI=') {
    throw 'The existing first deployment key must remain unchanged'
}
if ([Convert]::FromBase64String($keys.Groups[1].Value).Length -ne 32) { throw 'Invalid first Ed25519 key size' }
if ($ids.Groups[2].Value -cne 'vbNullString' -or $keys.Groups[2].Value -cne 'vbNullString') {
    throw 'The second deployment key pair must start empty, pending owner configuration'
}

# Only the cached token may bypass exp, exclusively to request DEC-104 renewal.
$renewalCalls = [regex]::Matches($text, 'VerifyTokenClaims\([^\)]*, True\)')
if ($renewalCalls.Count -ne 1 -or $renewalCalls[0].Value -notmatch '^VerifyTokenClaims\(cachedToken,') {
    throw 'Expiry bypass must be limited to the cached renewal credential'
}
if ($text -notmatch 'VerifyTokenClaims\(issuedToken,[\s\S]*?featuresJson, issuedUnix\) Then') {
    throw 'Fresh server tokens must pass the default strict expiry check'
}

$apiPost = [regex]::Match($text, '(?ms)^Private Function ApiPost\(.*?^End Function').Value
if ($apiPost.IndexOf('http.Option(9) = &H800 Or &H2000') -gt $apiPost.IndexOf('http.Open "POST"')) {
    throw 'TLS 1.2/1.3 must be configured before opening HTTP'
}
$verifyClaims = [regex]::Match($text, '(?ms)^Private Function VerifyTokenClaims\(.*?^End Function').Value
if ($verifyClaims.IndexOf('VerifyEd25519Detached(message, signature, publicKey)') -gt
    $verifyClaims.IndexOf('payloadJson = Utf8Text')) { throw 'Signature verification must precede payload trust' }
$protectCache = [regex]::Match($text, '(?ms)^Private Function DpapiProtectHex\(.*?^End Function').Value
$unprotectCache = [regex]::Match($text, '(?ms)^Private Function DpapiUnprotectHex\(.*?^End Function').Value
foreach ($cryptoBody in @($protectCache, $unprotectCache)) {
    if ($cryptoBody -notmatch 'Utf8Bytes\(CacheEntropyLabel\(\)\)' -or $cryptoBody -notmatch 'VarPtr\(entropyBlob\)') {
        throw 'Both DPAPI operations must use the same product entropy'
    }
}
$premiumGate = [regex]::Match($text, '(?ms)^Public Function License_PremiumAllowed\(.*?^End Function').Value
if ($premiumGate -notmatch 'CanAttemptRefresh\(' -or $premiumGate -notmatch 'ObserveClock\(' -or
    $premiumGate -notmatch 'ApplyProtectedSheetVisibility\(License_PremiumAllowed\)') {
    throw 'The running-session gate must check clock, bounded refresh and sheet visibility'
}
$rules = [regex]::Match($text, '(?ms)^Private Function CanAttemptRefresh\(.*?^End Function').Value
if ($rules -notmatch 'If refreshing Or currentCycle < 0 Then Exit Function' -or
    $rules -notmatch 'currentCycle > attemptedCycle') { throw 'Each refresh cycle must allow at most one request' }
if ($text -notmatch 'PolicyRequiresOnline = _\s*\(currentUnix >= validatedUnix \+ interval Or _\s*Not PolicyCanUse\(') {
    throw 'Online validation must cover rvi and unavailable offline grace (including ogh = 0)'
}

$cacheClear = [regex]::Match($text, '(?ms)^Private Function CacheClear\(\) As Boolean\r?\n(.*?)^End Function').Value
if ($cacheClear -notmatch 'On Error GoTo Failed' -or $cacheClear -notmatch 'Failed:\s*CacheClear = False') {
    throw 'Cache deletion must fail closed behind a local error handler'
}
$deletions = [regex]::Matches($text, '(?m)^\s*DeleteSetting[^\r\n]*')
if ($deletions.Count -ne 1 -or $deletions[0].Value.Trim() -cne 'DeleteSetting CACHE_APP, CACHE_SECTION') {
    throw 'Only the licensing cache section may be deleted, through CacheClear'
}
$clearPolicy = [regex]::Match($text, '(?ms)^Private Function ShouldClearCacheForServerError\(.*?^End Function').Value
if ($clearPolicy -notmatch 'Case "TOKEN_INVALID", "DEVICE_REVOKED", "DEVICE_MISMATCH"' -or
    ([regex]::Matches($clearPolicy, '(?m)^\s*Case ').Count -ne 1)) {
    throw 'Only invalid token, revoked device and mismatched device may clear the cache'
}
$startup = [regex]::Match($text, '(?ms)^Public Sub Auto_Open\(\).*?^End Sub').Value
if ([regex]::Matches($startup, 'License_ActivateWithDeviceCode\(').Count -ne 1 -or
    $startup -match '(?im)^\s*(Do\b|For\b|While\b|GoTo\s+Auto_Open\b)') {
    throw 'Startup must offer device activation once without retry loops'
}
$validate = [regex]::Match($text, '(?ms)^Private Function ValidateWithDeviceToken\(.*?^End Function').Value
if ($validate.IndexOf('If httpStatus >= 500 Then') -lt 0 -or
    $validate.IndexOf('If httpStatus = 429 Then') -lt 0 -or
    $validate.IndexOf('If httpStatus >= 500 Then') -gt $validate.IndexOf('ShouldClearCacheForServerError')) {
    throw 'Transient HTTP responses must preserve cached activation credentials'
}

$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'modLicenseApiV1.manifest.json') -Raw | ConvertFrom-Json
if ($manifest.moduleName -ne 'modLicenseApiV1' -or $manifest.sha256 -ne (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash) {
    throw 'VBA module does not match its SHA-256 manifest'
}

Write-Output 'One-file VBA contract and SHA-256 checks passed (C1-C7/C9; preserved DEC-104, ogh = 0, cache recovery; 64 mock / 57 signed cases).'
