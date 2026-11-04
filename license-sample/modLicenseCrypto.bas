Attribute VB_Name = "modLicenseCrypto"
Option Explicit

Private Type BCRYPT_PKCS1_PADDING_INFO
    pszAlgId As LongPtr
End Type

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider _
    Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider _
    Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptHash _
    Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal pbInput As LongPtr, _
    ByVal cbInput As Long, _
    ByVal pbOutput As LongPtr, _
    ByVal cbOutput As Long) As Long

Private Declare PtrSafe Function BCryptImportKeyPair _
    Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal hImportKey As LongPtr, _
    ByVal pszBlobType As LongPtr, _
    ByRef phKey As LongPtr, _
    ByVal pbInput As LongPtr, _
    ByVal cbInput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptVerifySignature _
    Lib "bcrypt.dll" ( _
    ByVal hKey As LongPtr, _
    ByVal pPaddingInfo As LongPtr, _
    ByVal pbHash As LongPtr, _
    ByVal cbHash As Long, _
    ByVal pbSignature As LongPtr, _
    ByVal cbSignature As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDestroyKey _
    Lib "bcrypt.dll" (ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Function WideCharToMultiByte _
    Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function CryptStringToBinaryW _
    Lib "crypt32.dll" ( _
    ByVal pszString As LongPtr, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, _
    ByRef pcbBinary As Long, _
    ByVal pdwSkip As LongPtr, _
    ByVal pdwFlags As LongPtr) As Long

Private Const CP_UTF8 As Long = 65001
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const BCRYPT_PAD_PKCS1 As Long = 2
Private Const BCRYPT_RSAPUBLIC_MAGIC As Long = &H31415352

Public Function Crypto_Utf8Bytes(ByVal value As String) As Byte()
    Dim size As Long
    Dim result() As Byte

    If Len(value) = 0 Then Err.Raise 5, , "Empty input"

    size = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(value), Len(value), 0, 0, 0, 0)
    If size <= 0 Then Err.Raise 5, , "UTF-8 conversion failed"

    ReDim result(0 To size - 1)
    If WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(value), Len(value), _
        VarPtr(result(0)), size, 0, 0) <> size Then
        Err.Raise 5, , "UTF-8 conversion failed"
    End If

    Crypto_Utf8Bytes = result
End Function

Public Function Crypto_Sha256(ByRef inputBytes() As Byte) As Byte()
    Dim algorithm As LongPtr
    Dim algorithmName As String
    Dim result() As Byte
    Dim inputLength As Long

    On Error GoTo Failed
    inputLength = UBound(inputBytes) - LBound(inputBytes) + 1
    If inputLength <= 0 Then GoTo Failed

    algorithmName = "SHA256"
    If BCryptOpenAlgorithmProvider( _
        algorithm, StrPtr(algorithmName), 0, 0) <> 0 Then GoTo Failed

    ReDim result(0 To 31)
    If BCryptHash( _
        algorithm, 0, 0, VarPtr(inputBytes(LBound(inputBytes))), _
        inputLength, VarPtr(result(0)), 32) <> 0 Then GoTo Failed

    Crypto_Sha256 = result
    BCryptCloseAlgorithmProvider algorithm, 0
    Exit Function

Failed:
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    Err.Raise 5, , "SHA-256 failed"
End Function

Public Function Crypto_BytesToHex(ByRef value() As Byte) As String
    Dim i As Long
    Dim result As String

    For i = LBound(value) To UBound(value)
        result = result & Right$("0" & Hex$(value(i)), 2)
    Next i
    Crypto_BytesToHex = result
End Function

Public Function Crypto_HexToBytes(ByVal value As String) As Byte()
    Dim result() As Byte
    Dim i As Long
    Dim pair As String
    Dim j As Long

    value = UCase$(value)
    If Len(value) = 0 Or (Len(value) Mod 2) <> 0 Then
        Err.Raise 5, , "Invalid hex"
    End If

    For i = 1 To Len(value)
        If InStr(1, "0123456789ABCDEF", Mid$(value, i, 1), _
                 vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Invalid hex"
        End If
    Next i

    ReDim result(0 To Len(value) \ 2 - 1)
    For j = 0 To UBound(result)
        pair = Mid$(value, j * 2 + 1, 2)
        result(j) = CByte(CLng("&H" & pair))
    Next j
    Crypto_HexToBytes = result
End Function

Public Function Crypto_Base64ToBytes(ByVal value As String) As Byte()
    Dim size As Long
    Dim result() As Byte

    If Len(value) = 0 Then Err.Raise 5, , "Empty Base64"

    If CryptStringToBinaryW( _
        StrPtr(value), Len(value), CRYPT_STRING_BASE64, _
        0, size, 0, 0) = 0 Then Err.Raise 5, , "Invalid Base64"

    If size <= 0 Then Err.Raise 5, , "Invalid Base64"
    ReDim result(0 To size - 1)

    If CryptStringToBinaryW( _
        StrPtr(value), Len(value), CRYPT_STRING_BASE64, _
        VarPtr(result(0)), size, 0, 0) = 0 Then
        Err.Raise 5, , "Invalid Base64"
    End If

    Crypto_Base64ToBytes = result
End Function

Private Sub PutDwordLE( _
    ByRef target() As Byte, ByVal offset As Long, ByVal value As Long)

    target(offset) = CByte(value And &HFF&)
    target(offset + 1) = CByte((value \ &H100&) And &HFF&)
    target(offset + 2) = CByte((value \ &H10000) And &HFF&)
    target(offset + 3) = CByte((value \ &H1000000) And &HFF&)
End Sub

Public Function Crypto_VerifyRsaSha256( _
    ByRef signedBytes() As Byte, _
    ByRef signature() As Byte, _
    ByVal modulusHex As String, _
    ByVal exponentHex As String) As Boolean

    Dim rsaAlgorithm As LongPtr
    Dim keyHandle As LongPtr
    Dim algorithmName As String
    Dim blobType As String
    Dim hashName As String
    Dim padding As BCRYPT_PKCS1_PADDING_INFO
    Dim modulus() As Byte
    Dim exponent() As Byte
    Dim digest() As Byte
    Dim blob() As Byte
    Dim i As Long
    Dim modulusLength As Long
    Dim exponentLength As Long
    Dim signatureLength As Long

    On Error GoTo Cleanup

    modulus = Crypto_HexToBytes(modulusHex)
    exponent = Crypto_HexToBytes(exponentHex)
    digest = Crypto_Sha256(signedBytes)

    modulusLength = UBound(modulus) + 1
    exponentLength = UBound(exponent) + 1
    signatureLength = UBound(signature) - LBound(signature) + 1

    If modulusLength <> 256 And modulusLength <> 384 Then GoTo Cleanup
    If exponentLength < 1 Or exponentLength > 8 Then GoTo Cleanup
    If modulus(0) = 0 Then GoTo Cleanup
    If signatureLength <> modulusLength Then GoTo Cleanup

    ReDim blob(0 To 24 + exponentLength + modulusLength - 1)
    PutDwordLE blob, 0, BCRYPT_RSAPUBLIC_MAGIC
    PutDwordLE blob, 4, modulusLength * 8
    PutDwordLE blob, 8, exponentLength
    PutDwordLE blob, 12, modulusLength
    PutDwordLE blob, 16, 0
    PutDwordLE blob, 20, 0

    For i = 0 To exponentLength - 1
        blob(24 + i) = exponent(i)
    Next i
    For i = 0 To modulusLength - 1
        blob(24 + exponentLength + i) = modulus(i)
    Next i

    algorithmName = "RSA"
    blobType = "RSAPUBLICBLOB"
    hashName = "SHA256"

    If BCryptOpenAlgorithmProvider( _
        rsaAlgorithm, StrPtr(algorithmName), 0, 0) <> 0 Then GoTo Cleanup

    If BCryptImportKeyPair( _
        rsaAlgorithm, 0, StrPtr(blobType), keyHandle, _
        VarPtr(blob(0)), UBound(blob) + 1, 0) <> 0 Then GoTo Cleanup

    padding.pszAlgId = StrPtr(hashName)

    Crypto_VerifyRsaSha256 = (BCryptVerifySignature( _
        keyHandle, VarPtr(padding), VarPtr(digest(0)), 32, _
        VarPtr(signature(LBound(signature))), signatureLength, _
        BCRYPT_PAD_PKCS1) = 0)

Cleanup:
    If keyHandle <> 0 Then BCryptDestroyKey keyHandle
    If rsaAlgorithm <> 0 Then BCryptCloseAlgorithmProvider rsaAlgorithm, 0
End Function
