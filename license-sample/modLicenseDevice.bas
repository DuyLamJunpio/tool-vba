Attribute VB_Name = "modLicenseDevice"
Option Explicit

Public Function Device_FingerprintHash() As String
    Dim shell As Object
    Dim machineGuid As String
    Dim inputBytes() As Byte
    Dim digest() As Byte
    Dim appId As String

    On Error GoTo Failed

    appId = Cfg_ApplicationId()
    If Len(appId) = 0 Then Exit Function

    Set shell = CreateObject("WScript.Shell")
    machineGuid = CStr(shell.RegRead( _
        "HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid"))
    If Len(machineGuid) = 0 Then Exit Function

    inputBytes = Crypto_Utf8Bytes( _
        "excel-license-device-v1|" & appId & "|" & LCase$(machineGuid))
    digest = Crypto_Sha256(inputBytes)
    Device_FingerprintHash = Crypto_BytesToHex(digest)
    Exit Function

Failed:
    Device_FingerprintHash = vbNullString
End Function
