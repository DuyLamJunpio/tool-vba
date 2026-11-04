Attribute VB_Name = "modLicenseStore"
Option Explicit

Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type

Private Declare PtrSafe Function CryptProtectData _
    Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, _
    ByVal szDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, _
    ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB) As Long

Private Declare PtrSafe Function CryptUnprotectData _
    Lib "crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, _
    ByVal ppszDataDescr As LongPtr, _
    ByVal pOptionalEntropy As LongPtr, _
    ByVal pvReserved As LongPtr, _
    ByVal pPromptStruct As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB) As Long

Private Declare PtrSafe Function LocalFree _
    Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal length As LongPtr)

Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = 1
Private Const STORE_APP As String = "VbaLicenseSample"
Private Const STORE_SECTION As String = "Cache"

Private Function Store_Key() As String
    Dim source() As Byte
    Dim digest() As Byte

    source = Crypto_Utf8Bytes(Cfg_ApplicationId())
    digest = Crypto_Sha256(source)
    Store_Key = Crypto_BytesToHex(digest)
End Function

Private Function Utc_Text(ByVal value As Date) As String
    Utc_Text = Format$(Year(value), "0000") & "-" & _
               Format$(Month(value), "00") & "-" & _
               Format$(Day(value), "00") & "T" & _
               Format$(Hour(value), "00") & ":" & _
               Format$(Minute(value), "00") & ":" & _
               Format$(Second(value), "00") & "Z"
End Function

Private Function Utc_Parse(ByVal value As String) As Date
    Dim parsed As Date

    If Len(value) <> 20 Then Err.Raise 5, , "Invalid UTC time"

    parsed = DateSerial( _
        CLng(Mid$(value, 1, 4)), _
        CLng(Mid$(value, 6, 2)), _
        CLng(Mid$(value, 9, 2))) + _
        TimeSerial( _
        CLng(Mid$(value, 12, 2)), _
        CLng(Mid$(value, 15, 2)), _
        CLng(Mid$(value, 18, 2)))

    If Utc_Text(parsed) <> value Then Err.Raise 5, , "Invalid UTC time"
    Utc_Parse = parsed
End Function

Public Function Store_Save( _
    ByVal rawToken As String, _
    ByVal verifiedUtc As Date) As Boolean

    Dim plain As String
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    Dim cipher() As Byte

    On Error GoTo Cleanup

    If Len(rawToken) = 0 Or Len(rawToken) > 32768 Then Exit Function
    If Len(Cfg_ApplicationId()) = 0 Then Exit Function

    plain = CStr(Len(rawToken)) & ":" & rawToken & Utc_Text(verifiedUtc)
    inputBlob.cbData = LenB(plain)
    inputBlob.pbData = StrPtr(plain)

    If CryptProtectData( _
        inputBlob, 0, 0, 0, 0, _
        CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup

    If outputBlob.cbData <= 0 Then GoTo Cleanup

    ReDim cipher(0 To outputBlob.cbData - 1)
    CopyMemory VarPtr(cipher(0)), outputBlob.pbData, outputBlob.cbData

    SaveSetting STORE_APP, STORE_SECTION, _
                Store_Key(), Crypto_BytesToHex(cipher)
    Store_Save = True

Cleanup:
    On Error Resume Next
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
End Function

Public Function Store_Load( _
    ByRef rawToken As String, _
    ByRef verifiedUtc As Date) As Boolean

    Dim storedHex As String
    Dim cipher() As Byte
    Dim inputBlob As DATA_BLOB
    Dim outputBlob As DATA_BLOB
    Dim plain As String
    Dim separator As Long
    Dim tokenLength As Long

    On Error GoTo Cleanup

    rawToken = vbNullString
    verifiedUtc = 0

    If Len(Cfg_ApplicationId()) = 0 Then Exit Function

    storedHex = GetSetting( _
        STORE_APP, STORE_SECTION, Store_Key(), vbNullString)
    If Len(storedHex) = 0 Then Exit Function

    cipher = Crypto_HexToBytes(storedHex)
    inputBlob.cbData = UBound(cipher) + 1
    inputBlob.pbData = VarPtr(cipher(0))

    If CryptUnprotectData( _
        inputBlob, 0, 0, 0, 0, _
        CRYPTPROTECT_UI_FORBIDDEN, outputBlob) = 0 Then GoTo Cleanup

    If outputBlob.cbData <= 0 Or _
       (outputBlob.cbData Mod 2) <> 0 Then GoTo Cleanup

    plain = String$(outputBlob.cbData \ 2, vbNullChar)
    CopyMemory StrPtr(plain), outputBlob.pbData, outputBlob.cbData

    separator = InStr(1, plain, ":", vbBinaryCompare)
    If separator < 2 Then GoTo Cleanup

    tokenLength = CLng(Left$(plain, separator - 1))
    If tokenLength <= 0 Or tokenLength > 32768 Then GoTo Cleanup
    If Len(plain) <> separator + tokenLength + 20 Then GoTo Cleanup

    rawToken = Mid$(plain, separator + 1, tokenLength)
    verifiedUtc = Utc_Parse( _
        Mid$(plain, separator + tokenLength + 1, 20))
    Store_Load = True

Cleanup:
    On Error Resume Next
    If outputBlob.pbData <> 0 Then LocalFree outputBlob.pbData
    If Not Store_Load Then
        rawToken = vbNullString
        verifiedUtc = 0
    End If
End Function
