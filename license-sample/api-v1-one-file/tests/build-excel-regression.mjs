// Creates a test-only module. Never import its generated .bas into a customer file.
// Uses the API signer with ephemeral test keys, never deployment secrets or .env.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { generateKeyPairSync, sign, verify } from 'node:crypto';
import assert from 'node:assert/strict';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const apiRoot = process.argv[2];
assert(apiRoot, 'Usage: node tests/build-excel-regression.mjs <API repository>');
const { ActivationTokenService } = await import(pathToFileURL(resolve(apiRoot,
  'apps/api/src/modules/licensing/token.service.ts')).href);
let source = readFileSync(resolve(root, 'modLicenseApiV1.bas'), 'ascii');
const kid = /Cfg_ActivationKeyIds = Array\("([^"]+)"/.exec(source)[1];
assert.equal(kid, 'k-d7b49d36a6111497', 'The first key slot must hold the production kid (owner, 08/11/2026)');
const primary = generateKeyPairSync('ed25519');
const secondary = generateKeyPairSync('ed25519');
const primaryPublic = primary.publicKey.export({ format: 'der', type: 'spki' }).subarray(-32).toString('base64');
const secondaryPublic = secondary.publicKey.export({ format: 'der', type: 'spki' }).subarray(-32).toString('base64');
const secondaryKid = 'rotation-regression-key';
const signer = new ActivationTokenService({
  privateKeyPem: primary.privateKey.export({ format: 'pem', type: 'pkcs8' }),
  publicKeyPem: primary.publicKey.export({ format: 'pem', type: 'spki' }),
  kid,
});
const now = Math.floor(Date.now() / 1000);
const realDateNow = Date.now;
Date.now = () => now * 1000;
const base = {
  lic: 'vba-regression-license-not-in-db', dev: 'vba-regression-device-not-in-db',
  jti: '11111111-1111-4111-8111-111111111111', prd: 'vba_regression', ver: '1.2.3',
  feat: ['regression_feature'], exp: now + 86400, rvi: 3600, ogh: 24,
};
const tokens = {
  valid: signer.signClaims(base),
  expired: signer.signClaims({ ...base, exp: now - 1 }),
  zero: signer.signClaims({ ...base, ogh: 0 }),
  wrongProduct: signer.signClaims({ ...base, prd: 'another_product' }),
  wrongVersion: signer.signClaims({ ...base, ver: '9.9.9' }),
  shortLived: signer.signClaims({ ...base, exp: now + 30 }),
};
const retiredSigner = new ActivationTokenService({
  privateKeyPem: primary.privateKey.export({ format: 'pem', type: 'pkcs8' }),
  publicKeyPem: primary.publicKey.export({ format: 'pem', type: 'spki' }),
  kid: 'retired-regression-key',
});
tokens.wrongKid = retiredSigner.signClaims(base);
const rotationSigner = new ActivationTokenService({
  privateKeyPem: secondary.privateKey.export({ format: 'pem', type: 'pkcs8' }),
  publicKeyPem: secondary.publicKey.export({ format: 'pem', type: 'spki' }),
  kid: secondaryKid,
});
tokens.secondary = rotationSigner.signClaims(base);
// DEC-107 addendum (11/11/2026), change 4: the server clock may lead this PC by up to 300 s.
Date.now = () => (now + 60) * 1000;
tokens.aheadIssued = signer.signClaims(base);
Date.now = () => (now + 301) * 1000;
tokens.futureIssued = signer.signClaims(base);
// Replies to a validation sent at now + 3600: issued one second, and 301 seconds, ahead of this PC.
Date.now = () => (now + 3601) * 1000;
tokens.refreshedAhead = signer.signClaims(base);
Date.now = () => (now + 3901) * 1000;
tokens.refreshedFuture = signer.signClaims(base);
// What another open copy receives from a successful validation 3590 s later (same jti, newer iat).
Date.now = () => (now + 3590) * 1000;
tokens.refreshed = signer.signClaims(base);
Date.now = realDateNow;
function resignFixture(token, headerChanges, payloadChanges) {
  const segments = token.split('.');
  const header = { ...JSON.parse(Buffer.from(segments[0], 'base64url')), ...headerChanges };
  const payload = { ...JSON.parse(Buffer.from(segments[1], 'base64url')), ...payloadChanges };
  const message = `${Buffer.from(JSON.stringify(header)).toString('base64url')}.${Buffer.from(JSON.stringify(payload)).toString('base64url')}`;
  return `${message}.${sign(null, Buffer.from(message), primary.privateKey).toString('base64url')}`;
}
tokens.mismatchedKid = resignFixture(tokens.valid, {}, { kid: secondaryKid });
tokens.wrongKeyForKid = resignFixture(tokens.valid, { kid: secondaryKid }, { kid: secondaryKid });
assert(verify(null, Buffer.from(tokens.valid.split('.').slice(0, 2).join('.')), primary.publicKey,
  Buffer.from(tokens.valid.split('.')[2], 'base64url')));
// Mutate a signature byte while preserving a syntactically valid JWT.
const parts = tokens.valid.split('.');
const sig = Buffer.from(parts[2], 'base64url');
sig[0] ^= 1;
tokens.tampered = `${parts[0]}.${parts[1]}.${sig.toString('base64url')}`;

function replaceProcedure(name, kind, body) {
  const re = new RegExp(`^(?:Public|Private) ${kind} ${name}\\b[\\s\\S]*?^End ${kind}`, 'm');
  assert(re.test(source), `Missing procedure ${name}`);
  source = source.replace(re, () => body.trim());
}
const replaceFunction = (name, body) => replaceProcedure(name, 'Function', body);
source = source.replace('Attribute VB_Name = "modLicenseApiV1"', 'Attribute VB_Name = "modLicenseRegression"');
source = source.replace(/^Public /gm, 'Private ');
// The production URLs are replaced so the harness can never reach shop.sunprime.vn.
function replaceOnce(from, to) {
  assert(source.includes(from), `Missing source text: ${from}`);
  source = source.replace(from, to);
}
replaceOnce('Cfg_ApiBaseUrl = "https://shop.sunprime.vn/api/v1"', 'Cfg_ApiBaseUrl = "https://regression.invalid/api/v1"');
source = source.replace('Cfg_ProductKey = vbNullString', 'Cfg_ProductKey = "vba_regression"');
source = source.replace('Cfg_ProductVersion = vbNullString', 'Cfg_ProductVersion = "1.2.3"');
replaceOnce('Cfg_ActivationWebUrl = "https://shop.sunprime.vn/account/activate"', 'Cfg_ActivationWebUrl = "https://regression.invalid/activate"');
assert(!source.includes('shop.sunprime.vn'), 'The harness must not keep a production URL');
assert(source.includes('ThisWorkbook.Close SaveChanges:=False'));
source = source.replaceAll('ThisWorkbook.Close SaveChanges:=False', 'tCloses = tCloses + 1');
const globals = `
Private tCached As String
Private tCachedTime As Double
Private tCachedFp As String
Private tResponse As String
Private tCode As String
Private tNetwork As Boolean
Private tRequests As Long
Private tSaves As Long
Private tClears As Long
Private tClearSucceeds As Boolean
Private tHttpStatus As Long
Private tActivationAttempts As Long
Private tGuidance As Long
Private tCloses As Long
Private tNow As Double
Private tMaxSeen As Double
Private tProductKey As String
Private tKeyMode As Long
Private tSheetsVisible As Boolean
Private tSheetConfigValid As Boolean
Private tCiphertext As String
Private tUseNativeCacheLoad As Boolean
Private tDenied As Boolean
Private tRawBody As String
Private tLoads As Long
Private tExpectedToken As String
`;
// Module-level variables must precede every procedure, so insert them before the first one.
const firstProcedure = /^Private (?:Function|Sub) /m.exec(source);
assert(firstProcedure, 'No procedure found');
source = source.slice(0, firstProcedure.index) + globals + '\n' + source.slice(firstProcedure.index);
replaceFunction('UtcUnixNow', `Private Function UtcUnixNow() As Double
    UtcUnixNow = tNow
End Function`);
replaceFunction('Cfg_ProductKey', `Private Function Cfg_ProductKey() As String
    Cfg_ProductKey = tProductKey
End Function`);
replaceFunction('Cfg_ActivationKeyIds', `Private Function Cfg_ActivationKeyIds() As Variant
    Select Case tKeyMode
        Case 1
            Cfg_ActivationKeyIds = Array("${kid}", "")
        Case 3
            Cfg_ActivationKeyIds = Array("${kid}", "${kid}")
        Case Else
            Cfg_ActivationKeyIds = Array("${kid}", "${secondaryKid}")
    End Select
End Function`);
replaceFunction('Cfg_ActivationPublicKeysBase64', `Private Function Cfg_ActivationPublicKeysBase64() As Variant
    Select Case tKeyMode
        Case 1, 2
            Cfg_ActivationPublicKeysBase64 = Array("${primaryPublic}", "")
        Case 4
            Cfg_ActivationPublicKeysBase64 = Array("${primaryPublic}", "YQ==")
        Case Else
            Cfg_ActivationPublicKeysBase64 = Array("${primaryPublic}", "${secondaryPublic}")
    End Select
End Function`);
replaceFunction('ApplyProtectedSheetVisibility', `Private Function ApplyProtectedSheetVisibility(ByVal allowed As Boolean) As Boolean
    tSheetsVisible = (allowed And tSheetConfigValid)
    ApplyProtectedSheetVisibility = tSheetConfigValid
End Function`);
replaceFunction('DeviceFingerprintHash', `Private Function DeviceFingerprintHash() As String
    DeviceFingerprintHash = String$(64, "a")
End Function`);
const nativeLoad = /^Private Function CacheLoad\([\s\S]*?^End Function/m.exec(source)[0]
  .replaceAll('CacheLoad', 'NativeCacheLoadTest')
  .replace('GetSetting(CACHE_APP, CacheSection(), "Ciphertext", vbNullString)', 'tCiphertext');
const nativeSave = /^Private Function CacheSave\([\s\S]*?^End Function/m.exec(source)[0]
  .replaceAll('CacheSave', 'NativeCacheSaveTest')
  .replace('SaveSetting CACHE_APP, CacheSection(), "Ciphertext", cipher', 'tCiphertext = cipher');
replaceFunction('CacheLoad', `Private Function CacheLoad( _
    ByRef token As String, ByRef fingerprint As String, ByRef validatedUnix As Double, _
    ByRef maxSeenUnix As Double, ByRef denied As Boolean) As Boolean
    tLoads = tLoads + 1
    If tUseNativeCacheLoad Then
        CacheLoad = NativeCacheLoadTest(token, fingerprint, validatedUnix, maxSeenUnix, denied)
        Exit Function
    End If
    token = tCached
    fingerprint = tCachedFp
    validatedUnix = tCachedTime
    maxSeenUnix = tMaxSeen
    denied = tDenied
    CacheLoad = (Len(tCached) > 0)
End Function`);
replaceFunction('CacheSave', `Private Function CacheSave( _
    ByVal token As String, ByVal fingerprint As String, ByVal validatedUnix As Double, _
    ByVal maxSeenUnix As Double, ByVal denied As Boolean) As Boolean
    tSaves = tSaves + 1
    If (token <> tResponse And token <> tCached) Or fingerprint <> String$(64, "a") Then Err.Raise 513
    ' Same effect on the module's clock state as the real CacheSave.
    maxSeenUnix = MaxDouble(maxSeenUnix, mMaxSeenUnix)
    If maxSeenUnix < tMaxSeen Then Err.Raise 513, , "time watermark decreased"
    If validatedUnix < tCachedTime Then Err.Raise 513, , "validation time decreased"
    mMaxSeenUnix = maxSeenUnix
    mSavedMaxSeenUnix = maxSeenUnix
    tCached = token
    tCachedFp = fingerprint
    tCachedTime = validatedUnix
    tMaxSeen = maxSeenUnix
    tDenied = denied
    CacheSave = True
End Function`);
replaceFunction('CacheClear', `Private Function CacheClear() As Boolean
    tClears = tClears + 1
    If Not tClearSucceeds Then Exit Function
    tCached = vbNullString
    tCachedFp = vbNullString
    tCachedTime = 0
    tMaxSeen = 0
    tDenied = False
    CacheClear = True
End Function`);
replaceFunction('License_ActivateWithDeviceCode', `Private Function License_ActivateWithDeviceCode() As Boolean
    tActivationAttempts = tActivationAttempts + 1
    ResetLicenseState
    mChecked = True
    mErrorCode = "ACTIVATION_CANCELLED"
End Function`);
replaceProcedure('License_ShowGuidance', 'Sub', `Private Sub License_ShowGuidance()
    tGuidance = tGuidance + 1
End Sub`);
replaceFunction('ApiPost', `Private Function ApiPost( _
    ByVal path As String, ByVal body As String, ByVal activationToken As String, _
    ByRef statusCode As Long, ByRef responseText As String) As Boolean
    Dim expectedToken As String
    tRequests = tRequests + 1
    ' The request must carry the shared cache's token, or this copy's own token when a case says so.
    expectedToken = tCached
    If Len(tExpectedToken) > 0 Then expectedToken = tExpectedToken
    If path <> "/activations/validate-device" Then Err.Raise 513
    If activationToken <> expectedToken Then Err.Raise 513
    If JsonSimpleString(body, "activation_token") <> expectedToken Then Err.Raise 513
    If Not tNetwork Then Exit Function
    ApiPost = True
    If Len(tRawBody) > 0 Then
        ' A reply that does not come from the API, such as a proxy block page.
        statusCode = tHttpStatus
        responseText = tRawBody
    ElseIf Len(tCode) > 0 Then
        statusCode = tHttpStatus
        responseText = "{""success"":false,""statusCode"":" & CStr(tHttpStatus) & _
                       ",""code"":""" & tCode & """,""message"":""regression""}"
    Else
        statusCode = 200
        responseText = "{""success"":true,""data"":{""token"":""" & tResponse & """}}"
    End If
End Function`);

const fixtureFunctions = Object.entries(tokens).map(([name, value]) =>
  `Private Function Fixture_${name}() As String\n    Fixture_${name} = "${value}"\nEnd Function`).join('\n\n');
const cases = [
  // label, cache, age, network, code, reply, allowed, status, HTTP calls, saves,
  // clear attempts, offer activation, HTTP status (default 403), cached fp, clear succeeds,
  // refusal mark stored beside the token (default false).
  ['cached within rvi', 'valid', 10, true, '', 'valid', true, 'ACTIVE', 0, 0, 0, false],
  ['expired renewed online', 'expired', 10, true, '', 'valid', true, 'ACTIVE', 1, 1, 0, false],
  ['expired offline preserves cache', 'expired', 10, false, '', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false],
  ['zero grace online', 'zero', 10, true, '', 'zero', true, 'ACTIVE', 1, 1, 0, false],
  ['zero grace offline preserves cache', 'zero', 0, false, '', 'zero', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false],
  ['device revoked clears cache', 'valid', 3600, true, 'DEVICE_REVOKED', 'valid', false, 'DEVICE_REVOKED', 1, 0, 1, true],
  ['license revoked blocks grace and is remembered', 'valid', 3600, true, 'LICENSE_REVOKED', 'valid', false, 'LICENSE_REVOKED', 1, 1, 0, false, 403, 'a', true, true],
  ['server invalid token clears cache', 'valid', 3600, true, 'TOKEN_INVALID', 'valid', false, 'TOKEN_INVALID', 1, 0, 1, true, 401],
  ['fresh response expired preserves credential', 'expired', 10, true, '', 'expired', false, 'TOKEN_INVALID', 1, 0, 0, false],
  ['tampered cache clears locally', 'tampered', 10, true, '', 'valid', false, 'TOKEN_INVALID', 0, 0, 1, true],
  ['wrong product clears locally', 'wrongProduct', 10, true, '', 'valid', false, 'TOKEN_INVALID', 0, 0, 1, true],
  ['wrong version clears locally', 'wrongVersion', 10, true, '', 'valid', false, 'TOKEN_INVALID', 0, 0, 1, true],
  ['rvi boundary online', 'valid', 3600, true, '', 'valid', true, 'ACTIVE', 1, 1, 0, false],
  ['network failure within grace', 'valid', 3600, false, '', 'valid', true, 'ACTIVE', 1, 0, 0, false],
  ['network failure past grace', 'valid', 86401, false, '', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false],
  ['tampered response preserves credential', 'expired', 10, true, '', 'tampered', false, 'TOKEN_INVALID', 1, 0, 0, false],
  ['changed hardware clears locally', 'valid', 10, true, '', 'valid', false, 'DEVICE_MISMATCH', 0, 0, 1, true, 403, 'b'],
  ['server device mismatch clears cache', 'valid', 3600, true, 'DEVICE_MISMATCH', 'valid', false, 'DEVICE_MISMATCH', 1, 0, 1, true],
  // DEC-107 addendum (11/11/2026): 429, 5xx and replies without a known API refusal count as no network.
  ['rate limited within grace', 'valid', 3600, true, 'RATE_LIMITED', 'valid', true, 'ACTIVE', 1, 0, 0, false, 429],
  ['rate limited past grace', 'valid', 86401, true, 'RATE_LIMITED', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false, 429],
  ['rate limited with zero grace stays closed', 'zero', 10, true, 'RATE_LIMITED', 'zero', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false, 429],
  ['rate limited with expired token stays closed', 'expired', 10, true, 'RATE_LIMITED', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false, 429],
  ['5xx within grace', 'valid', 3600, true, 'INTERNAL_ERROR', 'valid', true, 'ACTIVE', 1, 0, 0, false, 500],
  ['5xx past grace', 'valid', 86401, true, 'INTERNAL_ERROR', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false, 500],
  ['5xx invalid code preserves cache', 'valid', 3600, true, 'TOKEN_INVALID', 'valid', true, 'ACTIVE', 1, 0, 0, false, 503],
  ['5xx refusal code is not a refusal', 'valid', 3600, true, 'LICENSE_REVOKED', 'valid', true, 'ACTIVE', 1, 0, 0, false, 502],
  ['429 invalid code preserves cache', 'valid', 3600, true, 'TOKEN_INVALID', 'valid', true, 'ACTIVE', 1, 0, 0, false, 429],
  ['429 refusal code is not a refusal', 'valid', 3600, true, 'LICENSE_REVOKED', 'valid', true, 'ACTIVE', 1, 0, 0, false, 429],
  ['unknown api code within grace', 'valid', 3600, true, 'SOMETHING_NEW', 'valid', true, 'ACTIVE', 1, 0, 0, false, 400],
  ['unknown api code past grace', 'valid', 86401, true, 'SOMETHING_NEW', 'valid', false, 'REVALIDATION_REQUIRED', 1, 0, 0, false, 400],
  ['clock skew preserves cache', 'valid', 3600, true, 'REQUEST_TIMESTAMP_INVALID', 'valid', false, 'REQUEST_TIMESTAMP_INVALID', 1, 0, 0, false, 400],
  ['expired license preserves cache and is remembered', 'valid', 3600, true, 'LICENSE_EXPIRED', 'valid', false, 'LICENSE_EXPIRED', 1, 1, 0, false, 403, 'a', true, true],
  ['suspended license preserves cache and is remembered', 'valid', 3600, true, 'LICENSE_SUSPENDED', 'valid', false, 'LICENSE_SUSPENDED', 1, 1, 0, false, 403, 'a', true, true],
  ['device limit preserves cache and is remembered', 'valid', 3600, true, 'DEVICE_LIMIT_EXCEEDED', 'valid', false, 'DEVICE_LIMIT_EXCEEDED', 1, 1, 0, false, 403, 'a', true, true],
  ['missing license preserves cache and is remembered', 'valid', 3600, true, 'LICENSE_NOT_FOUND', 'valid', false, 'LICENSE_NOT_FOUND', 1, 1, 0, false, 404, 'a', true, true],
  ['server clear failure stays closed', 'valid', 3600, true, 'DEVICE_REVOKED', 'valid', false, 'CACHE_UNAVAILABLE', 1, 0, 1, false, 403, 'a', false],
  ['wrong kid clears locally', 'wrongKid', 10, true, '', 'valid', false, 'TOKEN_INVALID', 0, 0, 1, true],
  ['local clear failure stays closed', 'tampered', 10, true, '', 'valid', false, 'CACHE_UNAVAILABLE', 0, 0, 1, false, 403, 'a', false],
];
const vbBool = v => v ? 'True' : 'False';
const calls = cases.map(([label, cache, age, network, code, reply, allowed, status, requests, saves,
  clears, activate, http = 403, fp = 'a', clearSucceeds = true, denied = false]) =>
  `    RunCase "${label}", Fixture_${cache}(), ${age}#, ${vbBool(network)}, "${code}", _\n` +
  `            Fixture_${reply}(), ${vbBool(allowed)}, "${status}", ${requests}, ${saves}, _\n` +
  `            ${clears}, ${vbBool(activate)}, ${http}, "${fp}", ${vbBool(clearSucceeds)}, ${vbBool(denied)}\n` +
  '    passed = passed + 1').join('\n');
const hardeningTests = readFileSync(resolve(root, 'tests/signed-hardening-cases.bas'), 'ascii');
const extraTests = [...hardeningTests.matchAll(/^Private Sub (Case_[A-Za-z0-9_]+)\(\)/gm)].map(match => match[1]);
assert(extraTests.length > 0, 'Hardening cases must be included');
const total = cases.length + 5 + extraTests.length;
const harness = `
Private Sub PrepareCase(ByVal cached As String, ByVal age As Double, _
                        ByVal code As String, ByVal fp As String)
    tNow = ${now}#
    tMaxSeen = tNow
    tProductKey = "vba_regression"
    tKeyMode = 0
    tSheetConfigValid = True
    tSheetsVisible = False
    tUseNativeCacheLoad = False
    tCached = cached
    tCachedTime = UtcUnixNow() - age
    tCachedFp = String$(64, fp)
    tNetwork = True
    tCode = code
    tResponse = Fixture_valid()
    tHttpStatus = 403
    tClearSucceeds = True
    tRequests = 0
    tSaves = 0
    tClears = 0
    tActivationAttempts = 0
    tGuidance = 0
    tCloses = 0
    tDenied = False
    tRawBody = vbNullString
    tLoads = 0
    tExpectedToken = vbNullString
End Sub

Private Sub RunCase(ByVal label As String, ByVal cached As String, ByVal age As Double, _
                    ByVal network As Boolean, ByVal code As String, ByVal reply As String, _
                    ByVal allowed As Boolean, ByVal status As String, _
                    ByVal requests As Long, ByVal saves As Long, ByVal clears As Long, _
                    ByVal activate As Boolean, ByVal httpStatus As Long, _
                    ByVal fp As String, ByVal clearSucceeds As Boolean, _
                    ByVal denied As Boolean)
    Dim validated As Double
    PrepareCase cached, age, code, fp
    validated = tCachedTime
    tNetwork = network
    tResponse = reply
    tHttpStatus = httpStatus
    tClearSucceeds = clearSucceeds
    License_OnOpen
    If License_PremiumAllowed() <> allowed Then Err.Raise 513, , label & ": allowed"
    If License_StatusCode() <> status Then Err.Raise 513, , label & ": status"
    If tRequests <> requests Then Err.Raise 513, , label & ": requests"
    If tSaves <> saves Then Err.Raise 513, , label & ": cache save"
    If tClears <> clears Then Err.Raise 513, , label & ": cache clear"
    If NeedsDeviceActivation(mErrorCode, mCacheClearedForActivation) <> activate Then _
        Err.Raise 513, , label & ": activation prompt"
    If tDenied <> denied Then Err.Raise 513, , label & ": refusal mark"
    If clears = 1 And clearSucceeds Then
        If Len(tCached) <> 0 Then Err.Raise 513, , label & ": stale cache"
    ElseIf saves = 0 Or denied Then
        If tCached <> cached Or tCachedFp <> String$(64, fp) Or tCachedTime <> validated Then _
            Err.Raise 513, , label & ": cache not preserved"
    End If
    If allowed And (tGuidance <> 0 Or tCloses <> 0) Then _
        Err.Raise 513, , label & ": no message while access is allowed"
End Sub

Private Sub RunGateCase(ByVal label As String, ByVal code As String, _
                        ByVal fp As String, ByVal requests As Long)
    PrepareCase Fixture_valid(), 3600#, code, fp
    Auto_Open
    If tRequests <> requests Or tClears <> 1 Or tSaves <> 0 Then _
        Err.Raise 513, , label & ": gate cache/request count"
    If tActivationAttempts <> 1 Or tGuidance <> 1 Or tCloses <> 1 Then _
        Err.Raise 513, , label & ": gate must prompt once and close on cancel"
    If Len(tCached) <> 0 Or License_PremiumAllowed() Then _
        Err.Raise 513, , label & ": gate must stay closed"
    Auto_Open
    If tRequests <> requests Or tClears <> 1 Or tActivationAttempts <> 2 Or _
       tGuidance <> 2 Or tCloses <> 2 Then Err.Raise 513, , label & ": next open still stuck"
End Sub

Private Sub RunResumeCase(ByVal code As String)
    PrepareCase Fixture_valid(), 3600#, code, "a"
    License_OnOpen
    If License_PremiumAllowed() Or tClears <> 0 Or tCached <> Fixture_valid() Or _
       License_StatusCode() <> code Or Not tDenied Then Err.Raise 513, , code & ": initial denial"
    tCode = vbNullString
    License_OnOpen
    If Not License_PremiumAllowed() Or tRequests <> 2 Or tClears <> 0 Or _
       tSaves <> 2 Or tDenied Then Err.Raise 513, , code & ": resume after renewal/unlock"
End Sub

Public Sub License_RunSignedRegressionTests()
    Dim passed As Long
    On Error GoTo Failed
${calls}
    RunGateCase "revoked same-open recovery", "DEVICE_REVOKED", "a", 1
    passed = passed + 1
    RunGateCase "hardware same-open recovery", vbNullString, "b", 0
    passed = passed + 1
    RunGateCase "server invalid same-open recovery", "TOKEN_INVALID", "a", 1
    passed = passed + 1
    RunResumeCase "LICENSE_EXPIRED"
    passed = passed + 1
    RunResumeCase "LICENSE_SUSPENDED"
    passed = passed + 1
${extraTests.map(name => `    ${name}\n    passed = passed + 1`).join('\n')}
    MsgBox CStr(passed) & "/${total} API-signed regression cases passed. No live HTTP or real license.", vbInformation
    Exit Sub
Failed:
    MsgBox "Regression failed after " & CStr(passed) & " cases: " & Err.Description, vbExclamation
End Sub
`;
const output = resolve(root, 'TestResults', 'modLicenseRegression.bas');
mkdirSync(dirname(output), { recursive: true });
const generated = source + '\n' + fixtureFunctions + '\n' + harness + '\n' + nativeLoad + '\n' + nativeSave + '\n' + hardeningTests;
assert(!generated.includes('SaveSetting ') && !generated.includes('GetSetting(') && !generated.includes('DeleteSetting '),
  'The generated harness must never access a real Registry cache');
assert(!generated.includes('BEGIN PRIVATE KEY'), 'The harness contains public test keys only');
writeFileSync(output, generated.replace(/\r?\n/g, '\r\n'), 'ascii');
console.log('Generated test-only Excel harness: ' + total + ' cases; API signer, ephemeral keys, synthetic claims, mocked HTTP/cache/UI. VBA execution still required.');
