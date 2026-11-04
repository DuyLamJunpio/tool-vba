Attribute VB_Name = "modLicenseCore"
Option Explicit

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

Private Declare PtrSafe Sub GetSystemTime _
    Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)

Private mChecked As Boolean
Private mAllowed As Boolean

Private Function Clock_UtcNow() As Date
    Dim value As SYSTEMTIME
    GetSystemTime value

    Clock_UtcNow = DateSerial( _
        value.wYear, value.wMonth, value.wDay) + _
        TimeSerial(value.wHour, value.wMinute, value.wSecond)
End Function

Private Function Token_Authentic( _
    ByVal rawToken As String, _
    ByVal expectedDeviceHash As String, _
    ByRef token As TLicenseToken) As Boolean

    Dim signedBytes() As Byte
    Dim signature() As Byte

    On Error GoTo Failed

    If Not Contract_DecodeToken(rawToken, token) Then Exit Function
    If token.Audience <> Cfg_ApplicationId() Then Exit Function
    If token.DeviceHash <> expectedDeviceHash Then Exit Function
    If token.NotBeforeUtc = 0 Or token.ExpiresUtc = 0 Then Exit Function
    If token.ExpiresUtc <= token.NotBeforeUtc Then Exit Function

    signedBytes = Crypto_Base64ToBytes(token.SignedPayloadBase64)
    signature = Crypto_Base64ToBytes(token.SignatureBase64)

    Token_Authentic = Crypto_VerifyRsaSha256( _
        signedBytes, signature, _
        Cfg_PublicModulusHex(), Cfg_PublicExponentHex())
    Exit Function

Failed:
    Token_Authentic = False
End Function

Public Function License_CacheUsable( _
    ByVal proofValid As Boolean, _
    ByVal notBeforeUtc As Date, _
    ByVal expiresUtc As Date, _
    ByVal verifiedUtc As Date, _
    ByVal currentUtc As Date, _
    ByVal maxHours As Long) As Boolean

    If Not proofValid Then Exit Function
    If maxHours <= 0 Then Exit Function
    If verifiedUtc = 0 Then Exit Function
    If currentUtc < verifiedUtc Then Exit Function
    If currentUtc < notBeforeUtc Then Exit Function
    If currentUtc >= expiresUtc Then Exit Function
    If currentUtc > DateAdd("h", maxHours, verifiedUtc) Then Exit Function

    License_CacheUsable = True
End Function

Public Sub License_Initialize()
    Dim rawToken As String
    Dim deviceHash As String
    Dim verifiedUtc As Date
    Dim currentUtc As Date
    Dim oldToken As TLicenseToken
    Dim newToken As TLicenseToken
    Dim httpStatus As Long
    Dim responseBody As String
    Dim decision As Long
    Dim replacementToken As String

    On Error GoTo Failed

    mChecked = True
    mAllowed = False

    If Len(Cfg_ApplicationId()) = 0 Then Exit Sub
    If Len(Cfg_PublicModulusHex()) = 0 Then Exit Sub
    If Cfg_VerifyEveryHours() <= 0 Then Exit Sub
    If Cfg_OfflineGraceHours() < Cfg_VerifyEveryHours() Then Exit Sub

    deviceHash = Device_FingerprintHash()
    If Len(deviceHash) = 0 Then Exit Sub

    If Not Store_Load(rawToken, verifiedUtc) Then Exit Sub
    If Not Token_Authentic(rawToken, deviceHash, oldToken) Then Exit Sub

    currentUtc = Clock_UtcNow()
    If currentUtc < verifiedUtc Then Exit Sub

    If License_CacheUsable( _
        True, oldToken.NotBeforeUtc, oldToken.ExpiresUtc, _
        verifiedUtc, currentUtc, Cfg_VerifyEveryHours()) Then
        mAllowed = True
        Exit Sub
    End If

    If Api_PostVerify( _
        rawToken, deviceHash, httpStatus, responseBody) Then

        ' A received HTTP response must be understood by the contract.
        If Not Contract_ReadVerifyResponse( _
            httpStatus, responseBody, decision, replacementToken) Then
            Exit Sub
        End If

        If decision <> CONTRACT_ACCEPTED Then Exit Sub
        If Not Token_Authentic( _
            replacementToken, deviceHash, newToken) Then Exit Sub
        If currentUtc < newToken.NotBeforeUtc Then Exit Sub
        If currentUtc >= newToken.ExpiresUtc Then Exit Sub

        If Store_Save(replacementToken, currentUtc) Then
            mAllowed = True
        End If
        Exit Sub
    End If

    ' Only transport failure or a missing request adapter uses grace time.
    mAllowed = License_CacheUsable( _
        True, oldToken.NotBeforeUtc, oldToken.ExpiresUtc, _
        verifiedUtc, currentUtc, Cfg_OfflineGraceHours())
    Exit Sub

Failed:
    mAllowed = False
End Sub

Public Function License_PremiumAllowed() As Boolean
    If Not mChecked Then License_Initialize
    License_PremiumAllowed = mAllowed
End Function

Public Sub License_ShowGuidance()
    MsgBox _
        "Workbook da mo. Tinh nang cao cap chua duoc phep su dung." & _
        vbCrLf & vbCrLf & Cfg_ActivationHelp(), _
        vbInformation, "Giay phep"
End Sub

Public Sub Sample_PremiumAction()
    If Not License_PremiumAllowed() Then
        License_ShowGuidance
        Exit Sub
    End If

    ' Replace this line with the premium operation after integration.
    MsgBox "Tinh nang cao cap duoc phep chay.", vbInformation
End Sub
