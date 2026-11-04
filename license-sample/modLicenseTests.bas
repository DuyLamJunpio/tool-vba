Attribute VB_Name = "modLicenseTests"
Option Explicit

Private Sub Mock_Expect( _
    ByVal condition As Boolean, _
    ByVal caseName As String)

    If Not condition Then
        Err.Raise vbObjectError + 710, _
                  "License_RunMockTests", caseName
    End If
End Sub

Public Sub License_RunMockTests()
    Dim currentUtc As Date
    Dim verifiedUtc As Date
    Dim validFrom As Date
    Dim expiresUtc As Date
    Dim source() As Byte
    Dim digest() As Byte

    currentUtc = DateSerial(2026, 10, 25) + TimeSerial(12, 0, 0)
    verifiedUtc = DateAdd("h", -25, currentUtc)
    validFrom = DateAdd("d", -2, currentUtc)
    expiresUtc = DateAdd("d", 2, currentUtc)

    Mock_Expect Not License_CacheUsable( _
        False, validFrom, expiresUtc, verifiedUtc, _
        currentUtc, 72), "Invalid signature or claims"

    Mock_Expect License_CacheUsable( _
        True, validFrom, expiresUtc, verifiedUtc, _
        currentUtc, 72), "Offline within grace period"

    Mock_Expect Not License_CacheUsable( _
        True, validFrom, expiresUtc, verifiedUtc, _
        currentUtc, 24), "Verification interval exceeded"

    Mock_Expect Not License_CacheUsable( _
        True, validFrom, DateAdd("h", -1, currentUtc), _
        verifiedUtc, currentUtc, 72), "Expired token"

    Mock_Expect Not License_CacheUsable( _
        True, DateAdd("h", 1, currentUtc), expiresUtc, _
        verifiedUtc, currentUtc, 72), "Not-before claim"

    Mock_Expect Not License_CacheUsable( _
        True, validFrom, expiresUtc, DateAdd("h", 1, currentUtc), _
        currentUtc, 72), "Clock rollback"

    Mock_Expect Not License_CacheUsable( _
        True, validFrom, expiresUtc, DateAdd("h", -73, currentUtc), _
        currentUtc, 72), "Offline grace exceeded"

    source = Crypto_Utf8Bytes("abc")
    digest = Crypto_Sha256(source)
    Mock_Expect Crypto_BytesToHex(digest) = _
        "BA7816BF8F01CFEA414140DE5DAE2223" & _
        "B00361A396177A9CB410FF61F20015AD", _
        "SHA-256 known vector"

    MsgBox "8 mock checks passed.", vbInformation, "License tests"
End Sub
