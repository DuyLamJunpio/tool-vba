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

' Change 4 (11/11/2026): the signed iat may lead this PC's clock by up to 300 s, as the server allows.
Private Sub Case_IssuedAtAheadWithinToleranceAllowed()
    PrepareCase Fixture_aheadIssued(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tClears = 0 And tRequests = 0, "iat 60 s ahead of this clock is accepted"
    AssertCase tMaxSeen = tNow And tGuidance = 0 And tCloses = 0, "iat does not move the stored watermark past this clock"
End Sub

Private Sub Case_ServerClockSlightlyAheadValidates()
    ' Real case of 11/11/2026: this PC was 0.8 s behind the server and every new token was refused.
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tResponse = Fixture_refreshedAhead()
    AssertCase License_PremiumAllowed() And tRequests = 1, "token issued one second ahead is accepted"
    AssertCase tCached = Fixture_refreshedAhead() And tCachedTime = tNow, "that token and this clock's time are stored"
    AssertCase tMaxSeen = tNow, "watermark stays at this clock, not at the server's iat"
    AssertCase License_PremiumAllowed() And tRequests = 1 And tGuidance = 0 And tCloses = 0, "next check in the same second stays allowed"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1, "reopening in the same second stays allowed"
    tNow = tNow - 1#
    AssertCase Not License_PremiumAllowed() And License_StatusCode() = "CLOCK_ROLLBACK", "a clock below the stored watermark is still refused"
End Sub

Private Sub Case_ServerClockFarAheadRefused()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tResponse = Fixture_refreshedFuture()
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "token issued 301 s ahead is refused"
    AssertCase License_StatusCode() = "CLOCK_ROLLBACK", "refusal names the clock"
    AssertCase tCached = Fixture_valid() And tClears = 0 And Not tDenied, "refused token is not stored and the cache is kept"
    AssertCase tGuidance = 1 And tCloses = 1, "guidance then one close"
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

' DEC-107 addendum (11/11/2026), change 1: a server refusal is stored beside the token.
Private Sub AssertRefusalRemembered(ByVal code As String)
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed(), code & ": session starts allowed"
    tNow = tNow + 3600#
    tCode = code
    AssertCase Not License_PremiumAllowed() And tRequests = 1, code & ": refusal ends the session"
    AssertCase License_StatusCode() = code, code & ": refusal status"
    AssertCase tGuidance = 1 And tCloses = 1, code & ": guidance then one close"
    AssertCase tDenied And tClears = 0 And tCached = Fixture_valid(), code & ": refusal stored beside the kept token"
    ' The customer switches the network off and opens the file again, still inside ogh.
    tCode = vbNullString
    tNetwork = False
    tNow = tNow + 60#
    Auto_Open
    AssertCase Not License_PremiumAllowed() And tRequests = 2, code & ": offline reopen refused after one online attempt"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", code & ": offline reopen asks for a connection"
    AssertCase tGuidance = 2 And tCloses = 2 And tActivationAttempts = 0, code & ": offline reopen closes without a new code"
    AssertCase tDenied And tClears = 0 And tCached = Fixture_valid(), code & ": mark and token survive the offline reopen"
End Sub

Private Sub Case_RevokedRememberedOffline()
    AssertRefusalRemembered "LICENSE_REVOKED"
End Sub

Private Sub Case_SuspendedRememberedOffline()
    AssertRefusalRemembered "LICENSE_SUSPENDED"
End Sub

Private Sub Case_ExpiredRememberedOffline()
    AssertRefusalRemembered "LICENSE_EXPIRED"
End Sub

Private Sub Case_DeviceLimitRememberedOffline()
    AssertRefusalRemembered "DEVICE_LIMIT_EXCEEDED"
End Sub

Private Sub Case_LicenseNotFoundRememberedOffline()
    AssertRefusalRemembered "LICENSE_NOT_FOUND"
End Sub

Private Sub Case_RefusalMarkBlocksGraceInsideRvi()
    ' Without the mark this open is allowed from the cache with no request at all.
    PrepareCase Fixture_valid(), 10#, vbNullString, "a"
    tDenied = True
    tNetwork = False
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "mark forces an online attempt inside rvi"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", "mark with no network asks for a connection"
    AssertCase tDenied And tClears = 0 And tSaves = 0 And tCached = Fixture_valid(), "mark with no network keeps the cache"
End Sub

Private Sub AssertMarkedTransientStaysClosed(ByVal label As String, ByVal httpStatus As Long, _
                                             ByVal code As String, ByVal rawBody As String)
    PrepareCase Fixture_valid(), 10#, code, "a"
    tDenied = True
    tHttpStatus = httpStatus
    tRawBody = rawBody
    Auto_Open
    AssertCase Not License_PremiumAllowed() And tRequests = 1, label & ": marked cache stays closed"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", label & ": asks for a connection"
    AssertCase tGuidance = 1 And tCloses = 1 And tActivationAttempts = 0, label & ": guidance then close, no new code"
    AssertCase tDenied And tClears = 0 And tCached = Fixture_valid(), label & ": mark and token kept"
End Sub

Private Sub Case_RefusalMarkWithRateLimitStaysClosed()
    AssertMarkedTransientStaysClosed "429", 429, "RATE_LIMITED", vbNullString
End Sub

Private Sub Case_RefusalMarkWithServerFaultStaysClosed()
    AssertMarkedTransientStaysClosed "503", 503, "INTERNAL_ERROR", vbNullString
End Sub

Private Sub Case_RefusalMarkWithProxyPageStaysClosed()
    AssertMarkedTransientStaysClosed "proxy page", 403, vbNullString, "<html><body>Access denied</body></html>"
End Sub

Private Sub Case_RefusalMarkClearedBySuccess()
    PrepareCase Fixture_valid(), 10#, vbNullString, "a"
    tDenied = True
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1, "marked cache validates online on open"
    AssertCase Not tDenied And tSaves = 1 And tCachedTime = tNow, "successful validation removes the mark"
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed() And tRequests = 2, "offline time works again in the session"
    tNow = tNow + 60#
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 3 And Not tDenied, "offline time works again on the next open"
    AssertCase tGuidance = 0 And tCloses = 0, "no message after the mark is gone"
End Sub

Private Sub Case_RefusalMarkThenDeviceRevoked()
    PrepareCase Fixture_valid(), 10#, "DEVICE_REVOKED", "a"
    tDenied = True
    Auto_Open
    AssertCase tRequests = 1 And tClears = 1 And tSaves = 0, "marked cache is cleared on DEVICE_REVOKED"
    AssertCase Len(tCached) = 0 And Not tDenied, "cleared cache has no token and no mark"
    AssertCase tActivationAttempts = 1 And tGuidance = 1 And tCloses = 1, "one new code offer, then close on cancel"
End Sub

' Change 2: 429, 5xx and replies that are not an API refusal are handled like a lost connection.
Private Sub Case_SessionTransientRepliesUseGrace()
    Dim validated As Double
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    validated = tCachedTime
    License_OnOpen
    tNow = tNow + 3600#
    tCode = "RATE_LIMITED"
    tHttpStatus = 429
    AssertCase License_PremiumAllowed() And tRequests = 1, "429 in session uses offline time"
    AssertCase License_PremiumAllowed() And tRequests = 1, "429 is not retried in the same cycle"
    tNow = tNow + 3600#
    tCode = "INTERNAL_ERROR"
    tHttpStatus = 503
    AssertCase License_PremiumAllowed() And tRequests = 2, "5xx in session uses offline time"
    tNow = tNow + 3600#
    tCode = vbNullString
    tHttpStatus = 403
    tRawBody = "<html><body>Blocked ""code"":""DEVICE_REVOKED""</body></html>"
    AssertCase License_PremiumAllowed() And tRequests = 3, "proxy page in session uses offline time"
    AssertCase tCachedTime = validated And tCached = Fixture_valid(), "transient replies keep token and validation time"
    AssertCase tClears = 0 And Not tDenied And tGuidance = 0 And tCloses = 0, "transient replies are silent"
    tNow = tNow + 3600#
    tRawBody = vbNullString
    AssertCase License_PremiumAllowed() And tRequests = 4 And tCachedTime = tNow, "next cycle validates normally"
End Sub

Private Sub AssertSessionTransientPastGrace(ByVal label As String, ByVal httpStatus As Long, ByVal code As String)
    ' Validated 1000 s before the open, so ogh (24 h) runs out while exp is still 1000 s away.
    PrepareCase Fixture_valid(), 1000#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed(), label & ": session starts allowed"
    tNow = tNow + 85399#
    tCode = code
    tHttpStatus = httpStatus
    AssertCase License_PremiumAllowed() And tRequests = 1, label & ": last second of offline time"
    tNow = tNow + 2#
    AssertCase Not License_PremiumAllowed(), label & ": past offline time closes"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", label & ": asks for a connection"
    AssertCase tGuidance = 1 And tCloses = 1 And tClears = 0 And Not tDenied, label & ": guidance, close, cache kept"
    AssertCase tCached = Fixture_valid(), label & ": token kept"
End Sub

Private Sub Case_SessionRateLimitPastGraceCloses()
    AssertSessionTransientPastGrace "429", 429, "RATE_LIMITED"
End Sub

Private Sub Case_SessionServerFaultPastGraceCloses()
    AssertSessionTransientPastGrace "5xx", 500, "INTERNAL_ERROR"
End Sub

Private Sub Case_ProxyHtmlPageIsNetworkLoss()
    Dim validated As Double
    PrepareCase Fixture_valid(), 3600#, vbNullString, "a"
    validated = tCachedTime
    tHttpStatus = 403
    tRawBody = "<html><body>Blocked by policy. ""code"":""DEVICE_REVOKED""</body></html>"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1, "proxy block page uses offline time"
    AssertCase tClears = 0 And tSaves = 0 And Not tDenied, "proxy block page never clears or marks the cache"
    AssertCase tCached = Fixture_valid() And tCachedTime = validated, "proxy block page keeps token and validation time"
    AssertCase tGuidance = 0 And tCloses = 0, "proxy block page is silent inside ogh"
    PrepareCase Fixture_valid(), 86401#, vbNullString, "a"
    tHttpStatus = 403
    tRawBody = "<html><body>Blocked by policy</body></html>"
    Auto_Open
    AssertCase Not License_PremiumAllowed() And License_StatusCode() = "REVALIDATION_REQUIRED", "proxy page past ogh asks for a connection"
    AssertCase tClears = 0 And tActivationAttempts = 0 And tGuidance = 1 And tCloses = 1, "proxy page past ogh closes with the cache kept"
End Sub

Private Sub Case_CaptivePortalOkPageIsNetworkLoss()
    PrepareCase Fixture_valid(), 3600#, vbNullString, "a"
    tHttpStatus = 200
    tRawBody = "<html><body>Sign in to this network</body></html>"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1, "200 page without a token uses offline time"
    AssertCase tClears = 0 And tSaves = 0 And Not tDenied And tCached = Fixture_valid(), "200 page keeps the cache"
End Sub

Private Sub Case_RedirectIsNetworkLoss()
    PrepareCase Fixture_valid(), 3600#, vbNullString, "a"
    tHttpStatus = 302
    tRawBody = "Found"
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 1 And tClears = 0, "blocked redirect uses offline time"
End Sub

' Change 3: copies of one product on one computer share the validation stored in the cache.
Private Sub Case_CopyUsesValidationOfOtherCopy()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    ' Another copy validated ten seconds ago and stored its new token and time.
    tCached = Fixture_refreshed()
    tCachedTime = tNow - 10#
    tMaxSeen = tNow - 10#
    AssertCase License_PremiumAllowed() And tRequests = 0, "due copy reuses the other copy's validation"
    AssertCase tCached = Fixture_refreshed() And tCachedTime = tNow - 10#, "shared cache keeps the newer token and time"
    AssertCase tSheetsVisible And tGuidance = 0 And tCloses = 0, "reuse is transparent"
    tNow = tNow + 3000#
    AssertCase License_PremiumAllowed() And tRequests = 0, "no request for the rest of the shared cycle"
    tNow = tNow + 600#
    ' The mocked ApiPost refuses any token other than the one in the shared cache.
    AssertCase License_PremiumAllowed() And tRequests = 1, "next cycle sends one request with the shared token"
    AssertCase tCached = Fixture_valid() And tCachedTime = tNow, "that request stores the new validation"
End Sub

Private Sub Case_CopyValidatesWhenSharedCacheIsAlsoDue()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 7200#
    ' The other copy validated more than one rvi ago and was closed since.
    tCached = Fixture_refreshed()
    tCachedTime = tNow - 3610#
    tMaxSeen = tNow - 3610#
    AssertCase License_PremiumAllowed() And tRequests = 1, "stale shared validation still needs one request"
    AssertCase tCachedTime = tNow And tCached = Fixture_valid(), "request used the shared token and stored the result"
End Sub

Private Sub Case_StaleCopyKeepsNewerSharedCache()
    Dim sharedTime As Double
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3595#
    tCached = Fixture_refreshed()
    tCachedTime = tNow - 5#
    tMaxSeen = tNow - 5#
    sharedTime = tCachedTime
    ' Not due yet for this copy: it only stores its clock watermark.
    AssertCase License_PremiumAllowed() And tRequests = 0 And tSaves = 1, "copy inside its own rvi stores the watermark"
    AssertCase tCached = Fixture_refreshed() And tCachedTime = sharedTime, "watermark save keeps the newer token and time"
    AssertCase tMaxSeen = tNow, "watermark still advances"
    tDenied = True
    tNow = tNow + 1#
    AssertCase License_PremiumAllowed() And tSaves = 2, "second watermark save"
    AssertCase tDenied And tCached = Fixture_refreshed() And tCachedTime = sharedTime, "watermark save keeps a refusal mark"
End Sub

Private Sub Case_WatermarkSaveSkipsClearedCache()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    ' Another copy cleared the cache (DEVICE_REVOKED) while this one is open.
    tCached = vbNullString
    tCachedTime = 0
    tNow = tNow + 5#
    AssertCase License_PremiumAllowed() And tSaves = 0, "cleared cache is not written back"
    AssertCase Len(tCached) = 0 And tClears = 0, "cleared cache stays cleared"
End Sub

Private Sub Case_CopyRefusedAfterOtherCopyDenied()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    AssertCase License_PremiumAllowed(), "copy starts allowed"
    ' Another copy was refused by the server and stored the mark; then the network goes away.
    tDenied = True
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "other copy's refusal switches offline time off"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", "refused copy asks for a connection"
    AssertCase tGuidance = 1 And tCloses = 1 And Not tSheetsVisible, "refused copy shows guidance and closes"
    AssertCase tDenied And tClears = 0 And tCached = Fixture_valid(), "refused copy keeps mark and token"
End Sub

Private Sub Case_CopyOnGraceRefusedAfterOtherCopyDenied()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed() And tRequests = 1, "copy runs on offline time"
    tDenied = True
    tNow = tNow + 5#
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "mark ends offline time without a second request"
    AssertCase License_StatusCode() = "REVALIDATION_REQUIRED", "copy on offline time asks for a connection"
    AssertCase tGuidance = 1 And tCloses = 1 And tDenied, "copy on offline time closes once"
End Sub

Private Sub Case_CopyAllowedAgainAfterOtherCopyRecovered()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    tDenied = True
    tNetwork = False
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And tRequests = 1, "marked open with no network refused"
    ' Another copy validated successfully: the mark is gone and the validation is newer.
    tNow = tNow + 3590#
    tCached = Fixture_refreshed()
    tCachedTime = tNow
    tMaxSeen = tNow
    tDenied = False
    tNow = tNow + 20#
    AssertCase License_PremiumAllowed() And tRequests = 1, "copy takes over the other copy's successful validation"
End Sub

' A token found in the shared cache is used only after the same checks as on open.
Private Sub AssertUnusableSharedTokenIgnored(ByVal label As String, ByVal sharedToken As String)
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tCached = sharedToken
    tCachedTime = tNow - 10#
    tMaxSeen = tNow - 10#
    ' The mocked ApiPost now accepts only this copy's own verified token.
    tExpectedToken = Fixture_valid()
    AssertCase License_PremiumAllowed() And tRequests = 1, label & ": shared token not trusted, own token validated"
    AssertCase tCached = Fixture_valid() And tCachedTime = tNow And tClears = 0, label & ": server result stored"
End Sub

Private Sub Case_CopyIgnoresUnverifiableSharedToken()
    AssertUnusableSharedTokenIgnored "tampered", Fixture_tampered()
    AssertUnusableSharedTokenIgnored "other product", Fixture_wrongProduct()
    AssertUnusableSharedTokenIgnored "expired", Fixture_expired()
    AssertUnusableSharedTokenIgnored "unknown key", Fixture_wrongKid()
End Sub

Private Sub Case_CopyIgnoresSharedCacheOfOtherFingerprint()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tCached = Fixture_refreshed()
    tCachedTime = tNow - 10#
    tMaxSeen = tNow - 10#
    tCachedFp = String$(64, "b")
    tExpectedToken = Fixture_valid()
    AssertCase License_PremiumAllowed() And tRequests = 1, "entry of another fingerprint is neither adopted nor rewritten"
    AssertCase tCachedFp = String$(64, "a") And tCached = Fixture_valid(), "own validation replaces that entry"
End Sub

Private Sub Case_CopyIgnoresSharedCacheAheadOfClock()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 3600#
    tCached = Fixture_refreshed()
    tCachedTime = tNow - 10#
    tMaxSeen = tNow + 1#
    tNetwork = False
    tExpectedToken = Fixture_valid()
    AssertCase License_PremiumAllowed() And tRequests = 1, "entry stored ahead of this clock sample is not adopted"
    AssertCase tCached = Fixture_refreshed() And tMaxSeen = tNow + 1#, "that entry is left as it is"
    AssertCase tGuidance = 0 And tCloses = 0, "no clock message for a one-second lead"
End Sub

Private Sub Case_SharedWatermarkAheadIsNotARollback()
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNow = tNow + 5#
    ' Another Excel process stored a watermark one second ahead of the sample this copy holds.
    tMaxSeen = tNow + 1#
    AssertCase License_PremiumAllowed() And tRequests = 0, "watermark one second ahead does not close this copy"
    AssertCase License_StatusCode() = "ACTIVE" And tMaxSeen = tNow + 1#, "stored watermark kept, status active"
    AssertCase License_PremiumAllowed() And tRequests = 0 And tGuidance = 0 And tCloses = 0, "still allowed on the next call"
End Sub

Private Sub Case_SharedCacheReadOncePerSecond()
    Dim loadsAfterFirstCall As Long
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    License_OnOpen
    tNetwork = False
    tNow = tNow + 3600#
    AssertCase License_PremiumAllowed() And tRequests = 1, "copy runs on offline time"
    loadsAfterFirstCall = tLoads
    AssertCase License_PremiumAllowed() And License_PremiumAllowed() And License_PremiumAllowed(), "repeated checks stay allowed"
    AssertCase tLoads = loadsAfterFirstCall And tRequests = 1, "no further cache read or request in the same second"
    tNow = tNow + 1#
    AssertCase License_PremiumAllowed() And tLoads > loadsAfterFirstCall, "shared cache is read again in the next second"
End Sub

Private Sub Case_DpapiNativeCacheRoundTrip()
    Dim token As String
    Dim fingerprint As String
    Dim validated As Double
    Dim greatest As Double
    Dim denied As Boolean
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    ResetLicenseState
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow - 100#, tNow, False), "native cache encrypt"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest, denied), "native cache decrypt"
    AssertCase token = Fixture_valid() And greatest = tNow And validated = tNow - 100#, "native four-field cache"
    AssertCase Not denied, "plain cache has no refusal mark"
    AssertCase UBound(Split(Utf8Text(DpapiUnprotectHex(tCiphertext)), vbLf)) = 3, "unmarked cache keeps the four-field layout"
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), validated, tNow - 50#, False), "native lower input accepted safely"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest, denied) And greatest = tNow, "native watermark monotonic"
End Sub

Private Sub Case_DpapiRefusalMarkRoundTrip()
    Dim token As String
    Dim fingerprint As String
    Dim validated As Double
    Dim greatest As Double
    Dim denied As Boolean
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    ResetLicenseState
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow - 100#, tNow, True), "marked cache encrypt"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest, denied), "marked cache decrypt"
    AssertCase denied And token = Fixture_valid() And validated = tNow - 100# And greatest = tNow, "mark stored beside the four fields"
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow, tNow, False), "successful validation rewrites the cache"
    AssertCase NativeCacheLoadTest(token, fingerprint, validated, greatest, denied) And Not denied, "rewritten cache has no mark"
End Sub

Private Sub Case_DpapiFourFieldCacheStillLoads()
    ' Layout written by the module before 11/11/2026: customers activated with it must not activate again.
    Dim plain() As Byte
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    plain = Utf8Bytes(Fixture_valid() & vbLf & String$(64, "a") & vbLf & Format$(tNow - 10#, "0") & vbLf & Format$(tNow, "0"))
    tCiphertext = DpapiProtectHex(plain)
    AssertCase Len(tCiphertext) > 0, "four-field test encryption available"
    tUseNativeCacheLoad = True
    License_OnOpen
    AssertCase License_PremiumAllowed() And tRequests = 0 And tClears = 0, "earlier four-field cache is accepted unmarked"
End Sub

Private Sub Case_DpapiUnknownFifthFieldRejected()
    Dim plain() As Byte
    PrepareCase Fixture_valid(), 0#, vbNullString, "a"
    plain = Utf8Bytes(Fixture_valid() & vbLf & String$(64, "a") & vbLf & Format$(tNow - 10#, "0") & vbLf & Format$(tNow, "0") & vbLf & "allowed")
    tCiphertext = DpapiProtectHex(plain)
    tUseNativeCacheLoad = True
    License_OnOpen
    AssertCase Not License_PremiumAllowed() And License_StatusCode() = "ACTIVATION_REQUIRED", "unknown fifth field is not a usable cache"
    AssertCase tClears = 0 And tRequests = 0, "unknown fifth field is not a server refusal"
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
    AssertCase NativeCacheSaveTest(Fixture_valid(), String$(64, "a"), tNow, tNow, False), "product A encryption"
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
