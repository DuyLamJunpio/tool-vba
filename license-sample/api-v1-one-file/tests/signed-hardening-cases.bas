' Included by build-excel-regression.mjs; do not import this fragment alone.
Private Sub AssertCase(ByVal satisfied As Boolean, ByVal label As String)
    If Not satisfied Then Err.Raise 513, , label
End Sub

Private Sub Case_SecondaryKeyAccepted()
    PrepareCase Fixture_secondary(), 10#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 0 And tClears = 0, "secondary key accepted"
End Sub

Private Sub Case_SecondaryKeyDisabled()
    PrepareCase Fixture_secondary(), 10#, vbNullString, "a"
    tKeyMode = 1
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 1 And tRequests = 0, "disabled secondary refused"
    AssertCase License_StatusCode() = "TOKEN_INVALID", "disabled secondary status"
End Sub

Private Sub Case_HeaderPayloadKidMismatch()
    PrepareCase Fixture_mismatchedKid(), 10#, vbNullString, "a"
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 1 And tRequests = 0, "signed kid mismatch refused"
End Sub

Private Sub Case_WrongKeyForKnownKid()
    PrepareCase Fixture_wrongKeyForKid(), 10#, vbNullString, "a"
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 1 And tRequests = 0, "no fallback to another kid key"
End Sub

Private Sub Case_DuplicateKidConfiguration()
    PrepareCase Fixture_valid(), 10#, vbNullString, "a"
    tKeyMode = 3
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 0 And tRequests = 0, "duplicate kid config closed"
    AssertCase License_StatusCode() = "CONFIG_REQUIRED", "duplicate kid config status"
End Sub

Private Sub Case_PartialSecondKeyConfiguration()
    PrepareCase Fixture_valid(), 10#, vbNullString, "a"
    tKeyMode = 2
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 0, "partial second pair closed"
    AssertCase License_StatusCode() = "CONFIG_REQUIRED", "partial second pair status"
End Sub

Private Sub Case_MalformedSecondKeyConfiguration()
    PrepareCase Fixture_valid(), 10#, vbNullString, "a"
    tKeyMode = 4
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 0, "malformed second public key closed"
    AssertCase License_StatusCode() = "CONFIG_REQUIRED", "malformed second public key status"
End Sub

Private Sub Case_ClockBehindCachedMaximum()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    tMaxSeen = tNow + 1#
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 0 And tRequests = 0, "cached clock rollback closed"
    AssertCase License_StatusCode() = "CLOCK_ROLLBACK" And tSaves = 0, "rollback preserves cache"
End Sub

Private Sub Case_IssuedAtIsClockFloor()
    PrepareCase Fixture_futureIssued(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tClears = 0 And tRequests = 0, "iat floor closed"
    AssertCase License_StatusCode() = "CLOCK_ROLLBACK", "iat is not an invalid signature"
End Sub

Private Sub Case_ClockAtFloorAllowed()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tClears = 0 And tRequests = 0, "clock floor equality allowed"
End Sub

Private Sub Case_ClockRollbackDuringSession()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed(), "session starts allowed"
    tNow = tNow - 1#
    AssertCase Not License_PremiumAllowed(), "session clock rollback closed"
    AssertCase License_StatusCode() = "CLOCK_ROLLBACK" And Not tSheetsVisible, "rollback hides sheets"
    AssertCase tClears = 0 And tRequests = 0, "rollback keeps credential without HTTP"
    AssertCase tGuidance = 1 And tCloses = 1, "rollback shows guidance then closes once"
End Sub

Private Sub Case_ClockMaximumOnlyIncreases()
    Dim observed As Double
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 10#
    AssertCase License_PremiumAllowed(), "forward time allowed"
    observed = tMaxSeen
    tNow = tNow - 1#
    AssertCase Not License_PremiumAllowed() And tMaxSeen = observed, "watermark never decreases"
    AssertCase tSaves = 1 And tClears = 0, "only forward watermark persisted"
End Sub

Private Sub Case_ClockSaveKeepsValidationDeadline()
    Dim validated As Double
    PrepareCase Fixture_valid(), 100#, vbNullString, "a"
    validated = tCachedTime
    License_OnOpen
    tNow = tNow + 20#
    AssertCase License_PremiumAllowed(), "watermark observation allowed"
    AssertCase tCachedTime = validated And tSaves = 1 And tRequests = 0, "watermark cannot postpone rvi"
End Sub

Private Sub Case_SessionRefreshSuccess()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed(), "due session refresh succeeds"
    AssertCase tRequests = 1 And tCachedTime = tNow, "refresh updates validation anchor"
    AssertCase License_PremiumAllowed() And tRequests = 1, "successful refresh does not loop"
End Sub

Private Sub Case_SessionNetworkOnlyOncePerCycle()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed(), "network failure within grace"
    AssertCase License_PremiumAllowed() And tRequests = 1, "same cycle retries blocked"
    AssertCase tClears = 0, "network cache preserved"
End Sub

Private Sub Case_SessionRetriesNextCycle()
    Dim validated As Double
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    validated = tCachedTime
    License_OnOpen
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed() And tRequests = 1, "first failed cycle attempted"
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed() And tRequests = 2, "next cycle gets one attempt"
    AssertCase License_PremiumAllowed() And tRequests = 2, "next cycle also cannot loop"
    AssertCase tCachedTime = validated, "offline grace keeps last successful anchor"
End Sub

Private Sub Case_SessionReentryGuard()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    mRefreshInProgress = True
    AssertCase Not License_PremiumAllowed() And tRequests = 0, "reentry cannot send HTTP"
    mRefreshInProgress = False
    AssertCase License_PremiumAllowed() And tRequests = 1, "refresh works after guard release"
End Sub

Private Sub Case_SessionExpiredTokenRenews()
    PrepareCase Fixture_shortLived(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 31#
    AssertCase License_PremiumAllowed() And tRequests = 1, "exp in session renews via device token"
    AssertCase tCached = Fixture_valid() And tActivationAttempts = 0, "exp alone needs no device code"
End Sub

Private Sub Case_SessionZeroGraceNetworkFailure()
    PrepareCase Fixture_zero(), 0#, vbNullString, "a"
    tResponse = Fixture_zero()
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1, "zero grace open still validates"
    tNow = tNow + 3600#
    tNetwork = False
    AssertCase Not License_PremiumAllowed() And tRequests = 2, "zero grace session failure closed"
    AssertCase Not License_PremiumAllowed() And tRequests = 2, "zero grace does not loop"
    AssertCase tClears = 0 And Not tSheetsVisible, "zero grace failure keeps cache and hides sheets"
    AssertCase tGuidance = 1 And tCloses = 1, "zero grace session closes only once"
End Sub

Private Sub Case_SessionHardRejectionCannotUseGrace()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tCode = "LICENSE_REVOKED"
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "hard rejection closed"
    tNow = tNow + 3600#
    tCode = vbNullString
    tNetwork = False
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "closed session cannot retry or use grace"
    AssertCase tClears = 0 And Not tSheetsVisible, "hard refusal preserves previous cache rule"
    AssertCase tGuidance = 1 And tCloses = 1, "hard refusal closes once after guidance"
End Sub

Private Sub Case_DpapiNativeCacheRoundTrip()
    Dim token As String
    Dim fingerprint As String
    Dim validated As Double
    Dim greatest As Double
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    ResetLicenseState
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow - 100#, tNow), "native cache encrypt"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest), "native cache decrypt"
    AssertCase token = Fixture_valid() And greatest = tNow And validated = tNow - 100#, "native four-field cache"
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), validated, tNow - 50#), "native lower input accepted safely"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest) And greatest = tNow, "native watermark monotonic"
End Sub

Private Function LegacyTestCipher(ByVal value As String) As String
    Dim plain() As Byte
    Dim cipher() As Byte
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    On Error GoTo Cleanup
    plain = Utf8Bytes(value)
    inputBlob.cbData = ByteCount(plain)
    inputBlob.pbData = VarPtr(plain(0))
    If CryptProtectData(inputBlob, 0, 0, 0, 0, CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup
    ReDim cipher(0 To outputBlob.cbData - 1)
    CopyMemory VarPtr(cipher(0)), outputBlob.pbData, outputBlob.cbData
    LegacyTestCipher = BytesToHex(cipher)
Cleanup:
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
End Function

Private Sub Case_DpapiLegacyCacheIgnored()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    tCiphertext = LegacyTestCipher(Fixture_valid() & vbLf & String$(64, "a") & vbLf & Format$(tNow, "0"))
    AssertCase Len(tCiphertext) > 0, "legacy test encryption available"
    tUseNativeCacheLoad = True
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And License_StatusCode() = "ACTIVATION_REQUIRED", "legacy cache silently ignored"
    AssertCase tClears = 0 And tRequests = 0, "legacy cache is not a server rejection"
End Sub

Private Sub Case_DpapiOtherProductRejected()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    ResetLicenseState
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow, tNow), "product A encryption"
    tProductKey = "different_product"
    tUseNativeCacheLoad = True
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And License_StatusCode() = "ACTIVATION_REQUIRED", "other product cannot decrypt cache"
    AssertCase tClears = 0, "entropy failure is not a definitive refusal"
End Sub

Private Sub Case_SheetsHideOnDenial()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase tSheetsVisible, "verified open reveals premium sheets"
    tNow = tNow + 3600#
    tCode = "LICENSE_REVOKED"
    AssertCase Not License_PremiumAllowed() And Not tSheetsVisible, "refusal hides premium sheets"
End Sub

Private Sub Case_SheetsHideOnClose()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase tSheetsVisible, "allowed sheets visible before close"
    AssertCase License_OnClose() And Not tSheetsVisible, "close hides premium sheets"
    AssertCase Not License_PremiumAllowed() And tRequests = 0, "close cannot refresh or reveal sheets"
End Sub

Private Sub Case_SheetConfigurationFailsClosed()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    tSheetConfigValid = False
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And Not tSheetsVisible, "bad sheet config closed"
    AssertCase License_StatusCode() = "SHEET_CONFIG_INVALID" And tRequests = 0 And tClears = 0, "sheet error keeps cache"
End Sub
