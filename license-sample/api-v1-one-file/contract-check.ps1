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
    'If cachedDenied Or PolicyRequiresOnline(nowUnix, cachedTime, expiry, interval, graceHours) Then',
    'If mErrorCode <> "NETWORK_UNAVAILABLE" Then Exit Sub',
    'Private Const CACHE_DENIED_MARK As String = "denied"',
    'Private Const ISSUED_AT_TOLERANCE_SECONDS As Double = 300#',
    'ClockCanUse = (currentUnix > 0 And currentUnix >= maxSeenUnix And _',
    'currentUnix + ISSUED_AT_TOLERANCE_SECONDS >= issuedUnix)',
    'Private Function ShouldRememberServerDenial(',
    'Private Function IsServerRefusal(',
    'Private Function ApiErrorCode(',
    'Private Sub RememberServerDenial()',
    'Private Function SaveClockWatermark() As Boolean',
    'Private Sub AdoptSharedCacheState(',
    'If Not IsServerRefusal(apiCode, httpStatus) Then Exit Function',
    'If ShouldRememberServerDenial(mErrorCode, httpStatus) Then RememberServerDenial',
    'Private Function CacheClear() As Boolean',
    'DeleteSetting CACHE_APP, CacheSection()',
    'CacheSection = CACHE_SECTION & "_" & Cfg_ProductKey()',
    '""device_name""',
    'Private Function SafeDeviceName(',
    'SafeDeviceName(Environ$("COMPUTERNAME"))',
    'Private Function CopyTextToClipboard(',
    'If CopyTextToClipboard(userCode) Then',
    'vbYesNo + vbInformation, "Kich hoat giay phep"',
    'If ShouldClearCacheForServerError(mErrorCode, httpStatus) Then _',
    'InvalidateCachedActivation "DEVICE_MISMATCH"',
    'InvalidateCachedActivation "TOKEN_INVALID"',
    'mCacheClearedForActivation = CacheClear()',
    'If Not mCacheClearedForActivation Then mErrorCode = "CACHE_UNAVAILABLE"',
    'If httpStatus < 400 Or httpStatus >= 500 Or httpStatus = 429 Then Exit Function',
    'Optional ByVal allowExpiredForRenewal As Boolean = False',
    'If expiry <= nowUnix And Not allowExpiredForRenewal Then Exit Function',
    'If expiry <= 0 Or interval <= 0 Or graceHours < 0 Then Exit Function',
    'If countPassed = 71 Then',
    'If Not ClockCanUse(nowUnix, cachedMaxSeen, issuedUnix) Then',
    'ClockCanUse(currentUnix, mMaxSeenUnix, mIssuedUnix)',
    'maxSeenUnix = MaxDouble(maxSeenUnix, mMaxSeenUnix)',
    'mRefreshInProgress = True',
    'mRefreshInProgress = False',
    'mRefreshAttemptCycle = RevalidationCycle(currentUnix, mLastValidatedUnix, mRevalidateSeconds)',
    'If mClosing Or mRefreshInProgress Then Exit Function',
    'If UBound(fields) <> 3 And UBound(fields) <> 4 Then Exit Function',
    'If fields(4) <> CACHE_DENIED_MARK Then Exit Function',
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
    'Cfg_KeyId()', 'Cfg_Ed25519PublicKeyBase64()', 'GetDrive("C:")',
    'vbYesNoCancel', 'CACHE_APP, CACHE_SECTION',
    'Application.OnTime'
)) {
    if ($text.IndexOf($forbidden, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "Forbidden VBA source pattern: $forbidden"
    }
}

# Per-product values stay blank in the source; tool-vba fills them when packaging a product.
foreach ($blankConfig in @('Cfg_ProductKey', 'Cfg_ProductVersion', 'Cfg_ProxyServer')) {
    if ($text -notmatch ($blankConfig + ' = vbNullString')) { throw "$blankConfig must start blank" }
}
# Global production values (owner, 08/11/2026).
foreach ($globalConfig in @(
    'Cfg_ApiBaseUrl = "https://shop.sunprime.vn/api/v1"',
    'Cfg_ActivationWebUrl = "https://shop.sunprime.vn/account/activate"'
)) {
    if (-not $text.Contains($globalConfig)) { throw "Missing production configuration: $globalConfig" }
}
if ($text -notmatch 'Cfg_ProtectedSheetNames = Array\(\)') { throw 'Protected sheet names must start empty' }
$flat = $text -replace ' _\r?\n\s*', ' '
$ids = [regex]::Match($flat, 'Cfg_ActivationKeyIds = Array\("([A-Za-z0-9._-]+)", (vbNullString|"[A-Za-z0-9._-]+")\)')
$keys = [regex]::Match($flat, 'Cfg_ActivationPublicKeysBase64 = Array\(\s*"([A-Za-z0-9+/=]+)", (vbNullString|"[A-Za-z0-9+/=]+")\)')
if (-not $ids.Success -or -not $keys.Success) { throw 'Two explicit kid/public-key slots are required' }
# Production activation key of shop.sunprime.vn (owner, 08/11/2026); replaced the 29/10 k2026-10 test key.
if ($ids.Groups[1].Value -cne 'k-d7b49d36a6111497' -or
    $keys.Groups[1].Value -cne 'BdUaYauU+Te7xcYAb7ShcTlNUrQ47uRAsJqKboQm63s=') {
    throw 'The first key slot must hold the production activation key'
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
if ($deletions.Count -ne 1 -or $deletions[0].Value.Trim() -cne 'DeleteSetting CACHE_APP, CacheSection()') {
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

# DEC-107 addendum (11/11/2026). Change 2: 429, 5xx and replies without a known API refusal are a lost connection.
if ($validate.IndexOf('mErrorCode = "NETWORK_UNAVAILABLE"', $validate.IndexOf('If httpStatus < 200 Or httpStatus >= 300 Then')) -lt 0 -or
    $validate.IndexOf('If httpStatus >= 500 Then Exit Function') -lt 0 -or
    $validate.IndexOf('If httpStatus = 429 Then Exit Function') -lt 0 -or
    $validate.IndexOf('apiCode = ApiErrorCode(responseText)') -lt 0 -or
    $validate.IndexOf('If Not IsServerRefusal(apiCode, httpStatus) Then Exit Function') -gt $validate.IndexOf('ShouldClearCacheForServerError') -or
    $validate -match 'SERVER_REJECTED|RATE_LIMITED|SERVER_RESPONSE_INVALID') {
    throw 'validate-device: only a known API refusal may deny; 429, 5xx and other replies count as no network'
}
$apiErrorCode = [regex]::Match($text, '(?ms)^Private Function ApiErrorCode\(.*?^End Function').Value
if (-not $apiErrorCode.Contains('re.Pattern = "^\s*\{[\s\S]*\}\s*$"') -or
    $apiErrorCode.IndexOf('If Not re.Test(responseText) Then Exit Function') -gt $apiErrorCode.IndexOf('JsonSimpleString(responseText, "code")')) {
    throw 'An API error code may only be read from a JSON object body'
}
$refusal = [regex]::Match($text, '(?ms)^Private Function IsServerRefusal\(.*?^End Function').Value
if (-not $refusal.Contains('If httpStatus < 400 Or httpStatus >= 500 Or httpStatus = 429 Then Exit Function') -or
    -not $refusal.Contains('ShouldClearCacheForServerError(status, httpStatus)') -or
    -not $refusal.Contains('ShouldRememberServerDenial(status, httpStatus)') -or
    -not $refusal.Contains('status = "REQUEST_TIMESTAMP_INVALID"')) {
    throw 'A server refusal is a 4xx (not 429) reply with one of the known API codes'
}
# Change 1: the refusal is stored beside the token and switches offline time off.
$rememberPolicy = [regex]::Match($text, '(?ms)^Private Function ShouldRememberServerDenial\(.*?^End Function').Value
if (-not $rememberPolicy.Contains('If httpStatus < 400 Or httpStatus >= 500 Or httpStatus = 429 Then Exit Function') -or
    $rememberPolicy -notmatch 'Case "LICENSE_REVOKED", "LICENSE_SUSPENDED", "LICENSE_EXPIRED", _\s*"DEVICE_LIMIT_EXCEEDED", "LICENSE_NOT_FOUND"' -or
    ([regex]::Matches($rememberPolicy, '(?m)^\s*Case ').Count -ne 1)) {
    throw 'Exactly the five entitlement refusals are remembered in the cache'
}
$remember = [regex]::Match($text, '(?ms)^Private Sub RememberServerDenial\(\).*?^End Sub').Value
if ($remember -notmatch 'CacheSave token, fingerprint, validatedUnix, maxSeenUnix, True') {
    throw 'The refusal mark must be written beside the cached token'
}
$onOpen = [regex]::Match($text, '(?ms)^Public Sub License_OnOpen\(\).*?^End Sub').Value
if ($onOpen -notmatch 'If cachedDenied Then\s+mHardDenied = True\s+mErrorCode = "REVALIDATION_REQUIRED"\s+Exit Sub') {
    throw 'A stored refusal must refuse offline time on open'
}
$acceptToken = [regex]::Match($text, '(?ms)^Private Function AcceptServerToken\(.*?^End Function').Value
if ($acceptToken -notmatch 'CacheSave\(issuedToken, fingerprint, currentUnix, mMaxSeenUnix, False\)') {
    throw 'Only a token accepted from the server may store the cache without the refusal mark'
}
$unmarkedSaves = [regex]::Matches($text, 'CacheSave[ (][^\r\n]*, False\)?')
if ($unmarkedSaves.Count -ne 1) { throw 'Exactly one place may store the cache without the refusal mark' }
# Change 4: the signed iat may lead this PC's clock by the tolerance, and never raises the stored watermark.
foreach ($iatFold in @('MaxDouble(currentUnix, issuedUnix)', 'MaxDouble(currentUnix, mIssuedUnix)', 'MaxDouble(maxSeenUnix, issuedUnix)')) {
    if ($text.Contains($iatFold)) { throw "The token's iat must not be folded into the clock watermark: $iatFold" }
}
if ([regex]::Matches($text, 'ISSUED_AT_TOLERANCE_SECONDS').Count -ne 2) {
    throw 'The iat tolerance may only be used in ClockCanUse'
}
# Change 3: copies share one cache entry; a copy never overwrites newer shared state with its own.
$observe = [regex]::Match($text, '(?ms)^Private Function ObserveClock\(.*?^End Function').Value
if ($observe -match 'CacheSave' -or $observe -notmatch 'If Not SaveClockWatermark\(\) Then') {
    throw 'The clock watermark must be merged into the shared cache, not written from memory'
}
$watermark = [regex]::Match($text, '(?ms)^Private Function SaveClockWatermark\(\).*?^End Function').Value
if ($watermark -notmatch 'CacheLoad\(token, fingerprint, validatedUnix, maxSeenUnix, denied\)' -or
    $watermark -notmatch 'CacheSave\(token, fingerprint, validatedUnix, _\s*MaxDouble\(maxSeenUnix, mMaxSeenUnix\), denied\)') {
    throw 'A watermark save must keep the token, validation time and refusal mark already stored'
}
if ($watermark -notmatch 'observedUnix = mMaxSeenUnix' -or $watermark -notmatch 'mMaxSeenUnix = observedUnix' -or
    $watermark -notmatch 'StrComp\(fingerprint, mFingerprint, vbBinaryCompare\) <> 0') {
    throw 'A watermark save must not raise this copy''s own clock floor or rewrite an entry of another fingerprint'
}
if ($premiumGate.IndexOf('AdoptSharedCacheState currentUnix') -lt 0 -or
    $premiumGate.IndexOf('AdoptSharedCacheState currentUnix') -gt $premiumGate.IndexOf('AttemptRevalidation(')) {
    throw 'A due copy must read the shared cache before it sends a request'
}
$adopt = [regex]::Match($text, '(?ms)^Private Sub AdoptSharedCacheState\(.*?^End Sub').Value
if ($adopt -notmatch 'VerifyTokenClaims\(token, currentUnix,' -or $adopt -notmatch 'ClockCanUse\(currentUnix, maxSeenUnix, issuedUnix\)' -or
    $adopt.IndexOf('VerifyTokenClaims(') -gt $adopt.IndexOf('StageVerifiedCredential')) {
    throw 'A token taken from the shared cache must be verified before it is used'
}
if ($adopt -notmatch 'If currentUnix = mSharedCacheReadUnix Then Exit Sub' -or
    $adopt -notmatch 'StrComp\(fingerprint, mFingerprint, vbBinaryCompare\) <> 0' -or
    $adopt -match 'VerifyTokenClaims\([^\)]*, True\)') {
    throw 'The shared cache is read at most once per second, for this fingerprint, without the expiry bypass'
}

$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'modLicenseApiV1.manifest.json') -Raw | ConvertFrom-Json
if ($manifest.moduleName -ne 'modLicenseApiV1' -or $manifest.sha256 -ne (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash) {
    throw 'VBA module does not match its SHA-256 manifest'
}

Write-Output 'One-file VBA contract and SHA-256 checks passed (C1-C7/C9; preserved DEC-104, ogh = 0, cache recovery; 11/11/2026 addendum: remembered refusal, transient replies, shared cache).'
