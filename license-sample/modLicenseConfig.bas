Attribute VB_Name = "modLicenseConfig"
Option Explicit

Public Function Cfg_ApplicationId() As String
    ' Set a stable, unique product/workbook identifier after agreement.
    Cfg_ApplicationId = vbNullString
End Function

Public Function Cfg_VerifyUrl() As String
    ' Set only an HTTPS endpoint defined by the API contract.
    Cfg_VerifyUrl = vbNullString
End Function

Public Function Cfg_PublicModulusHex() As String
    ' Provisional RSA sample only; design document calls for Ed25519.
    Cfg_PublicModulusHex = vbNullString
End Function

Public Function Cfg_PublicExponentHex() As String
    ' Keep this only if the agreed RSA public key uses exponent 65537.
    Cfg_PublicExponentHex = "010001"
End Function

Public Function Cfg_VerifyEveryHours() As Long
    Cfg_VerifyEveryHours = 24
End Function

Public Function Cfg_OfflineGraceHours() As Long
    Cfg_OfflineGraceHours = 72
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

Public Function Cfg_ActivationHelp() As String
    Cfg_ActivationHelp = "Vui long lien he don vi cap phep de kich hoat hoac xac minh giay phep."
End Function
