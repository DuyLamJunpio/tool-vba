Attribute VB_Name = "modLicenseContract"
Option Explicit

Public Type TLicenseToken
    SignedPayloadBase64 As String
    SignatureBase64 As String
    Audience As String
    DeviceHash As String
    NotBeforeUtc As Date
    ExpiresUtc As Date
End Type

Public Const CONTRACT_ACCEPTED As Long = 1
Public Const CONTRACT_REJECTED As Long = 2

Public Function Contract_DecodeToken( _
    ByVal rawToken As String, _
    ByRef token As TLicenseToken) As Boolean

    ' Contract boundary: decode the exact signed bytes and the signature.
    ' Audience, device hash and times must come from the signed payload.
    ' Fail closed until the token format is defined and implemented.
    Contract_DecodeToken = False
End Function

Public Function Contract_BuildVerifyJson( _
    ByVal rawToken As String, _
    ByVal deviceHash As String) As String

    ' Contract boundary: define request JSON only after API agreement.
    Contract_BuildVerifyJson = vbNullString
End Function

Public Function Contract_ReadVerifyResponse( _
    ByVal httpStatus As Long, _
    ByVal responseBody As String, _
    ByRef decision As Long, _
    ByRef replacementToken As String) As Boolean

    ' Return True only for a fully recognized server response.
    ' CONTRACT_ACCEPTED requires a signed replacement token.
    decision = 0
    replacementToken = vbNullString
    Contract_ReadVerifyResponse = False
End Function
