Attribute VB_Name = "basCrypt"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' basCrypt                                                  '
'                                                           '
' (c) 2017-2025 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20250225                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines a library of cryptographic functions  '
' based on the Windows Secure Hash Algorithm (SHA) and Data '
' Encryption Standard (DES) algorithms.                     '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' System, mscorlib.dll, .NET 3.5                            '
' basLibVBA, basLibArray                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the module has only been tested with      '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Binary
Option Explicit

'''''''''''''''''''''''
' Types and Constants '
'''''''''''''''''''''''

Private Const kPadding = 2  ' PaddingMode.PKCS7
Private Const kMode = 1     ' CipherMode.CBC

''''''''''''''''''
' Hash Functions '
''''''''''''''''''

Public Function MD5( _
    ByVal aString As String, _
    Optional ByVal aBase64 As Boolean = False _
) As String
    '
    ' Returns the MD5 hash of aString as a hexadecimal string,
    ' or optionally as a base64 string if aBase64 is set.
    '
    ' For an empty input string, returns:
    '       Hex:    d41d8cd98f00b204e9800998ecf8427e
    '       Base64: 1B2M2Y8AsgTpgAmY7PhCfg==
    '
    Dim bytes() As Byte
    
    bytes = NewMd5Hash().ComputeHash_2(NewUtf8().GetBytes_4(aString))
    If aBase64 = True Then
       MD5 = BytesToBase64(bytes)
    Else
       MD5 = BytesToHex(bytes)
    End If
End Function

Public Function SHA1( _
    ByVal aString As String, _
    Optional aBase64 As Boolean = False _
) As String
    '
    ' Returns the SHA1 hash of aString as a hexadecimal string,
    ' or optionally as a base64 string if aBase64 is set.
    '
    ' For an empty input string, returns:
    '       Hex:    da39a3ee5e6b4b0d3255bfef95601890afd80709
    '       Base64: 2jmj7l5rSw0yVb/vlWAYkK/YBwk=
    '
    Dim bytes() As Byte
    
    bytes = NewSHA1Hash().ComputeHash_2(NewUtf8().GetBytes_4(aString))
    If aBase64 = True Then
       SHA1 = BytesToBase64(bytes)
    Else
       SHA1 = BytesToHex(bytes)
    End If
End Function

Public Function SHA256( _
    ByVal aString As String, _
    Optional aBase64 As Boolean = False _
) As String
    '
    ' Returns the SHA256 hash of aString as a hexadecimal string,
    ' or optionally as a base64 string if aBase64 is set.
    '
    ' For an empty input string, returns:
    '       Hex:    e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855
    '       Base64: 47DEQpj8HBSa+/TImW+5JCeuQeRkm5NMpJWZG3hSuFU=/vlWAYkK/YBwk=
    '
    Dim bytes() As Byte
    
    bytes = NewSHA256Hash().ComputeHash_2(NewUtf8().GetBytes_4(aString))
    If aBase64 = True Then
       SHA256 = BytesToBase64(bytes)
    Else
       SHA256 = BytesToHex(bytes)
    End If
End Function

Public Function SHA384( _
    ByVal aString As String, _
    Optional aBase64 As Boolean = False _
) As String
    '
    ' Returns the SHA384 hash of aString as a hexadecimal string,
    ' or optionally as a base64 string if aBase64 is set.
    '
    ' For an empty input string, returns:
    '       Hex:    38b060a751ac96384cd9327eb1b1e36a21fdb71114be07434c0cc7bf63f6e1da274edebfe76f65fbd51ad2f14898b95b
    '       Base64: OLBgp1GsljhM2TJ+sbHjaiH9txEUvgdDTAzHv2P24donTt6/529l+9Ua0vFImLlb
    '
    Dim bytes() As Byte
    
    bytes = NewSHA384Hash().ComputeHash_2(NewUtf8().GetBytes_4(aString))
    If aBase64 = True Then
       SHA384 = BytesToBase64(bytes)
    Else
       SHA384 = BytesToHex(bytes)
    End If
End Function

Public Function SHA512( _
    ByVal aString As String, _
    Optional aBase64 As Boolean = False _
) As String
    '
    ' Returns the SHA384 hash of aString as a hexadecimal string,
    ' or optionally as a base64 string if aBase64 is set.
    '
    ' For an empty input string, returns:
    '       Hex:    cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e
    '       Base64: z4PhNX7vuL3xVChQ1m2AB9Yg5AULVxXcg/SpIdNs6c5H0NE8XYXysP+DGNKHfuwvY7kxvUdBeoGlODJ6+SfaPg==
    '
    Dim bytes() As Byte
    
    bytes = NewSHA512Hash().ComputeHash_2(NewUtf8().GetBytes_4(aString))
    If aBase64 = True Then
       SHA512 = BytesToBase64(bytes)
    Else
       SHA512 = BytesToHex(bytes)
    End If
End Function

Function SHA512Salt( _
    ByVal aString As String, _
    ByVal aKey As String, _
    Optional ByVal aBase64 As Boolean = False _
) As String
    '
    ' Returns the SHA512 hash of aString, modified by the aKey argument,
    ' as a hexadecimal string or optionally as a base64 string if aBase64
    ' is set. Empty input string results depend on the key provided.
    '
    Dim utf8 As Object, crypto As Object, bytes() As Byte
    
    On Error GoTo Catch
    Set utf8 = NewUtf8()
    Set crypto = NewHMACSHA512()
    crypto.key = utf8.GetBytes_4(aKey)
    bytes = crypto.ComputeHash_2(utf8.GetBytes_4(aString))
    If aBase64 = True Then
       SHA512Salt = BytesToBase64(bytes)
    Else
       SHA512Salt = BytesToHex(bytes)
    End If
    
Catch:
    Set utf8 = Nothing
    Set crypto = Nothing
End Function

'''''''''''''''''''''''''''''''
' DES Cryptographic Functions '
'''''''''''''''''''''''''''''''

Public Function EncryptString( _
    ByVal aClearText As String, _
    ByVal aKey As String, _
    ByVal aInitVector As String _
) As Variant
    '
    ' Returns aClearText encrypted with the Windows
    ' Data Encryption Standard (DES) algorithm using
    ' aKey as the encryption key and aInitVector as
    ' the initialization vector.
    '
    Dim crypto As New TripleDESCryptoServiceProvider, utf8 As Object
    Dim clear_text() As Byte, cypher_data() As Byte, cypher_text As String

    On Error GoTo Catch
    Set utf8 = NewUtf8()
    clear_text = utf8.GetBytes_4(aClearText)
    crypto.Padding = kPadding
    crypto.Mode = kMode
    crypto.key = utf8.GetBytes_4(aKey)
    crypto.IV = utf8.GetBytes_4(aInitVector)
    cypher_data = crypto.CreateEncryptor().TransformFinalBlock(clear_text, 0, ArraySize(clear_text))
    cypher_text = BytesToBase64(cypher_data)
    EncryptString = cypher_text

Catch:
    Set utf8 = Nothing
    Set crypto = Nothing
End Function

Public Function DecryptString( _
    ByVal aCipherText As String, _
    ByVal aKey As String, _
    ByVal aInitVector As String _
) As Variant
    '
    ' Returns aCipherText decrypted with the Windows
    ' Data Encryption Standard (DES) algorithm using
    ' aKey as the encryption key and aInitVector as
    ' the initialization vector.
    '
    Dim crypto As New TripleDESCryptoServiceProvider, utf8 As Object
    Dim cipher_data() As Byte, clear_data() As Byte, clear_text As String

    On Error GoTo Catch
    Set utf8 = NewUtf8()
    cipher_data = Base64ToBytes(aCipherText)
    crypto.Padding = kPadding
    crypto.Mode = kMode
    crypto.key = utf8.GetBytes_4(aKey)
    crypto.IV = utf8.GetBytes_4(aInitVector)
    clear_data = crypto.CreateDecryptor().TransformFinalBlock(cipher_data, 0, ArraySize(cipher_data))
    clear_text = utf8.GetString(clear_data)
    DecryptString = clear_text

Catch:
    Set utf8 = Nothing
    Set crypto = Nothing
End Function

'''''''''''''''''''''''''''''''
' String Conversion Functions '
'''''''''''''''''''''''''''''''

Public Function Base64ToBytes(aBase64 As String) As Byte()
    '
    ' Returns a base64 string converted to an array of bytes.
    '
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = aBase64
         Base64ToBytes = .nodeTypedValue
    End With
End Function

Public Function BytesToBase64(aBytes() As Byte) As String
    '
    ' Returns an array of bytes converted to a base64 string.
    '
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = aBytes
        BytesToBase64 = Replace(.Text, vbLf, "")
    End With
End Function

Public Function BytesToHex(aBytes() As Byte) As String
    '
    ' Returns an array of bytes converted to a hexadecimal string.
    '
    With CreateObject("MSXML2.DomDocument").createElement("b64")
        .DataType = "bin.Hex"
        .nodeTypedValue = aBytes
        BytesToHex = Replace(.Text, vbLf, "")
    End With
End Function

Public Function HexToBytes(aHex As String) As Byte
    '
    ' Returns a hexadecimal string converted to an array of bytes.
    '
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.Hex"
         .Text = aHex
         HexToBytes = .nodeTypedValue
    End With
End Function

Public Function BytesToUnicode(aBytes() As Byte) As String
    '
    ' Returns an array of bytes converted to a Unicode string.
    '
    BytesToUnicode = StrConv(aBytes, vbUnicode)
End Function

Public Function UnicodeToBytes(ByVal aUnicode As String) As Byte()
    '
    ' Returns a Unicode string converted to an array of bytes.
    '
    Dim bytes() As Byte, i As Integer

    ReDim bytes(Len(aUnicode) - 1)
    For i = 0 To UBound(bytes)
        bytes(i) = asc(mid(aUnicode, i + 1, 1))
    Next i
    UnicodeToBytes = bytes
End Function

''''''''''''''''''''''''''''
' Class Factory Extensions '
''''''''''''''''''''''''''''

Public Function NewUtf8() As Object
    Dim obj As Object
    
    Set obj = CreateObject("System.Text.UTF8Encoding")
    Set NewUtf8 = obj
    Set obj = Nothing
End Function

Public Function NewMd5Hash() As MD5CryptoServiceProvider
    Dim obj As New MD5CryptoServiceProvider
    
    Set NewMd5Hash = obj
    Set obj = Nothing
End Function

Public Function NewSHA1Hash() As SHA1Managed
    Dim obj As New SHA1Managed
    
    Set NewSHA1Hash = obj
    Set obj = Nothing
End Function

Public Function NewSHA256Hash() As SHA256Managed
    Dim obj As New SHA256Managed
    
    Set NewSHA256Hash = obj
    Set obj = Nothing
End Function

Public Function NewSHA384Hash() As SHA384Managed
    Dim obj As New SHA384Managed
    
    Set NewSHA384Hash = obj
    Set obj = Nothing
End Function

Public Function NewSHA512Hash() As SHA512Managed
    Dim obj As New SHA512Managed
    
    Set NewSHA512Hash = obj
    Set obj = Nothing
End Function

Public Function NewHMACSHA512() As HMACSHA512
    Dim obj As New HMACSHA512
    
    Set NewHMACSHA512 = obj
    Set obj = Nothing
End Function

