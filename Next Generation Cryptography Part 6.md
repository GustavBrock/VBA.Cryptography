# HMAC hashing in VBA using the Microsoft NG Cryptography (CNG) API

![Title](images/EE%20Title%20HMAC.png)

This article series will show you how to utilise the *Next Generation Cryptography (CNG)* API from Microsoft for modern *hashing* and *encrypting/decrypting*, maintaining 100% compatibility with the implementation of CNG in *.Net* as well as *PowerShell* scripts.

---
## Introduction

### Purpose of these articles

The Microsoft CNG APIs constitute a collection of more than a dozen APIs that handle all the aspects and supporting functions to calculate hash values and perform encryption and decryption meeting modern high demands and standards.

This series details one method to implement these features in applications supported by VBA (Visual Basic for Applications), for example *Microsoft Access* and *Microsoft Excel*.

> The theory for and the clever mathematics behind hashing and encryption will not be touched; this is covered massively in books and articles published over a long period of time by dedicated experts. A search with Bing or Google will return a long list of resources to study.

### Sections

The series has been split in six parts. This allows you to skip parts you are either familiar with or wish to implement later if at all.

1. [Utilise Microsoft's Next Generation Cryptography (CNG) API in VBA](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%201.md)
2. [Hashing in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%202.md)
3. [Encryption in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%203.md)
4. [Using binary storage to serve the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%204.md)
5. [Storing passwords in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%205.md)
6. [HMAC hashing in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%206.md)

The three first and the last, as stated, with hashing and encrypting/decrypting based on the CNG API, while the forth explains how to combine these techniques with the little known feature of Access, *storing binary data*, as this in many cases will represent the optimal storage method for hashed or encrypted data.
The fifth demonstrates how to save and verify passwords totally safe using these tools.

## Part 6. HMAC

### Basic usage

HMAC (Hash-Based Message Authentication Code) is the process of creating a non-reversible fixed length code that identifies the input.

Note, that the code is not unique as, in theory, the same code could be generated from different inputs, though - due to the extreme count of combinations for a given string length - this is unlikely to happen in real life.

This, however, is neither the point. The point is to be able to create a code based on an input and a key, (later) to recreate that code using the same key, and then to check if the later code matches the first; if it does, the later input is assumed to match the original input.

The typical usage of HMAC values is for exchanging a hash value encrypted with the key which only the sender and the receiver have, thus impossible to recreate by a third party not having the key. Instead of exchanging the real value, the encrypted hash value is exchanged. To verify the value, its encrypted hash value is generated and compared with the received encrypted hash of the value. If these values match, the passed information is accepted.

### Hashing methods

The mathematical method used to hash a string is called an algorithm. These have evolved over time, and the oldest - also producing the shortest hash values - are now deprecated, though still widely used for less demanding purposes. For example, **SHA1**, and even the very old **MD5**, are often used. 

But for HMAC however, only the newer SHA algorithms are accepted by the CNG API.

The algorithms supported and their byte lengths are:

| Name  | Byte Length |
| ----- | ------ |
| SHA1  |     20 |
| SHA256|     32 |
| SHA384|     48 |
| SHA512|     64 |

While **SHA1** can be well suited for some purposes, the general consensus is, that **SHA256** is the minimum method to use for hashing passwords. For higher demands, use **SHA512** or **SHA384**, but be aware, that **SHA384** is not supported in all environments.

### Working with HMAC values

Note, that the lengths listed above for the hash algorithms are the lengths of the generated *byte arrays*. As these bytes can take many other values than those for normal ASCII characters - even control characters and other non-printable characters - the byte array must for many practical purposes be converted to something printable.

The simple method would be to convert to the hexadecimal representation of the bytes, for example four bytes (here character `<tab>` is the third):

    aX<tab>! -> 6158092D

but that consumes *twice as many characters*, thus a **SHA256** hash would have a length of 64 characters - not very efficient.

A better method is to use *Base64* encoding. It offers the same - a printable and portable string of ASCII characters - but consumes only 50% or less additional space then the byte count:

| Name  | Base64 Length |
| ----- | ------ |
| SHA1  |     28 |
| SHA256|     44 |
| SHA384|     64 |
| SHA512|     88 |

To conclude: For only a little overhead, the hash values can be treated as normal plain text.

### Creating a HMAC

A set of functions is used to create HMAC values:

~~~txt
Hmac
    HmacData
    TextUtf8Bytes
        CngHmac
    ByteBase64
~~~

The top function, `Hmac`, the one that creates a Base64 encoded HMAC value from a passed text and key, contains less than a handful of code lines:

~~~vb
' Return a Base64 encoded HMAC (Hash-Based Message Authentication Code) hash of a string
' using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' Example:
'   Text = "Several Species of Small Furry Animals Gathered Together in a Cave and Grooving With a Pict Lyrics."
'   Key = "Sysyphus"
'   Value = Hmac(Text, Key)
'   Value -> "KoSbRMsshQwgNKi9H2uG5NDyl+qdzNM6s2tG8AM9wk8="
'   Value = Hmac(Text, Key, bcSha512)
'   Value -> "+B2HzGp8bWfdE1WihvobFQEbAX91zdDA8fiGXFuAT4fjjBquEU55K5ooSOO3jJYN8NPLdySQf5DVR1y31rHkhQ=="
'
' Length of the generated Base64 encoded hash string:
'
'   Encoding    Length
'   SHA1        28
'   SHA256      44      ' Default.
'   SHA384      64
'   SHA512      88
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Hmac( _
    ByVal Text As String, _
    ByVal Key As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As String
    
    Dim TextData()          As Byte
    Dim KeyData()           As Byte
    
    Dim HmacBase64          As String
    
    If Text = "" Or Key = "" Then
        ' No data. Nothing to do.
    Else
        ' UTF-8 encode Text and Key.
        TextData = TextUtf8Bytes(Text)
        KeyData = TextUtf8Bytes(Key)
        HmacBase64 = ByteBase64(HmacData(TextData, KeyData, BcryptHashAlgorithmId))
    End If
    
    Hmac = HmacBase64
    
End Function
~~~

The only optional argument is which hash algorithm to use; by default, **SHA256** is used.

The usage is extremely simple:

~~~vb
Dim Text    As String
Dim Key     As String
Dim Value   As String

Text = "Get your filthy hands off my desert."
Key = "Have a cigar!"
Value = Hmac(Text, Key)

Debug.Print "Text input:", Text
Debug.Print "HMAC value:", Value

' Output:
' Text input:   Get your filthy hands off my desert.
' Hash value:   WGlSSyl6QZxhlURHyJvLhEMU0MdvaXstBBntHJPPHus=
~~~

The few lines of code in the function is possible, because it pulls data from the second level function, `HmacData`, also taking an optional argument for the hash algorithm, which defaults to **SHA256**:

~~~vb
' Create and return the HMAC (Hash-Based Message Authentication Code) hash of a byte array as
' another byte array using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' To be called from function Hmac.
'
' 2025-11-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HmacData( _
    ByRef TextData() As Byte, _
    ByRef KeyData() As Byte, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()
    
    Dim Data()          As Byte
    
    Dim TextDataLength  As Long
    Dim KeyDataLength   As Long
    
    If StrPtr(TextData) = 0 Then
        ' Array TextData is empty.
    ElseIf StrPtr(KeyData) = 0 Then
        ' Array KeyData is empty.
    Else
        TextDataLength = UBound(TextData) - LBound(TextData) + 1
        KeyDataLength = UBound(KeyData) - LBound(KeyData) + 1
        Data = CngHmac(VarPtr(TextData(LBound(TextData))), TextDataLength, VarPtr(KeyData(LBound(KeyData))), KeyDataLength, BcryptHashAlgorithmId)
    End If
    
    HmacData = Data
    
End Function
~~~

The important feature of the function is, that both input and output are *byte arrays*. However, you may have noticed, that function `Hmac` above doesn't pass a byte array, but plain text:

~~~vb
HmacData((Text), (Key), BcryptHashAlgorithmId)
~~~

That's because you can assign a text variable directly to a byte array; no conversion is needed.

The output is, in function `Hmac`, fed directly to function `ByteBase64`, which converts the byte array to Base64 encoded text. Though quite convoluted, that function is trivial and won't be discussed here.

Function `HmacData` also contains very few code lines. Essentially, it only wraps the function, that does "the real work", `CngHmac`, which makes only a single call to the CNG API to obtain the HMAC value.

The details of the call will not be discussed here. However, each step is commented in-line, and links to the documentation at Microsoft are included, should you wish to study this further.

The usage of this function is also very straight:

~~~vb
Dim Data()  As Byte
Dim Key()   As Byte
Dim Value() As Byte

Data = "Text to hash"
Key = "Secret"
Value = HmacData(Data, Key)

' Convert byte array to unicode.
Debug.Print StrConv(Value, vbUnicode)

' Output.
' ‘Ëå@Ä”1-†>§üZkpòrzÑ†AL$”æÛC073
~~~

Part 4 of this series deals with various methods for storing hash and HMAC values, where indeed these two functions can prove useful:

- `SaveBinaryField`
- `ReadBinaryField`

### Supplemental functions

As you may have noticed, an *enum* is used for specifying the hash algorithm:

~~~vb
' Allowed BCrypt hash algorithms.
Public Enum BcHashAlgorithm
    [_First] = 1
    bcMd2 = 1
    bcMd4 = 2
    bcMd5 = 3
    [_FirstHmac] = 4
    bcSha1 = 4
    bcSha256 = 5
    bcSha384 = 6
    bcSha512 = 7
    [_Last] = 7
    [_LastHmac] = 7
End Enum
~~~

This is to validate input and make it easy to always supply the rightly spelled and uppercased hash algorithm name to function `CngHash` and `CngHmac`.

Function `BcryptHashAlgorithm` serves this purpose:

~~~vb
' Return the literal hash algorithm name determined by
' the passed value of enum BcHashAlgorithm.
'
' To be called from functions CngHash and IsBcryptHashAlgorithm.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function BcryptHashAlgorithm( _
    ByVal BcryptHashAlgorithmId As BcHashAlgorithm) _
    As String

    Dim HashAlgorithmName       As String

    ' Note: HashAlgorithmName must be in UPPERCASE.
    Select Case BcryptHashAlgorithmId
        Case BcHashAlgorithm.bcMd2
            HashAlgorithmName = "MD2"
        Case BcHashAlgorithm.bcMd4
            HashAlgorithmName = "MD4"
        Case BcHashAlgorithm.bcMd5
            HashAlgorithmName = "MD5"
        Case BcHashAlgorithm.bcSha1
            HashAlgorithmName = "SHA1"
        Case BcHashAlgorithm.bcSha256
            HashAlgorithmName = "SHA256"
        Case BcHashAlgorithm.bcSha384
            HashAlgorithmName = "SHA384"
        Case BcHashAlgorithm.bcSha512
            HashAlgorithmName = "SHA512"
    End Select

    BcryptHashAlgorithm = HashAlgorithmName

End Function
~~~

Likewise, functions are included to verify/validate either a literal algorithm name or an enum value for these names:

~~~vb
' Return True if the passed text value represents a value of
' enum BcHashAlgorithm valid for HMAC.
' Note: To validate, HashAlgorithm must be in UPPERCASE.
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashHmacAlgorithm( _
    ByVal HashAlgorithm As String) _
    As Boolean
    
    Dim Index           As BcHashAlgorithm
    Dim Result          As Boolean
    
    For Index = BcHashAlgorithm.[_FirstHmac] To BcHashAlgorithm.[_LastHmac]
        If BcryptHashAlgorithm(Index) = HashAlgorithm Then
            Result = True
            Exit For
        End If
    Next
    
    IsBcryptHashHmacAlgorithm = Result
    
End Function
~~~

~~~vb
' Return True if the passed value of enum BcHashAlgorithm is valid for HMAC.
'
' To be called from function CngHmac.
'
' 2025-11-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashHmacAlgorithmId( _
    ByVal HashAlgorithmId As BcHashAlgorithm) _
    As Boolean
    
    Dim Result          As Boolean
    
    If BcHashAlgorithm.[_FirstHmac] <= HashAlgorithmId And HashAlgorithmId <= BcHashAlgorithm.[_LastHmac] Then
        Result = True
    End If
    
    IsBcryptHashHmacAlgorithmId = Result
    
End Function
~~~

Finally, to determine the length of a HMAC value - either as a byte array or a Base64 encoded string - without actually hashing something, two simple functions, `HashByteLength` and `HashTextLength`, will return those values. These functions were discussed in Part 1.

These functions can come in handy for code that creates or alters table design where fields for hash values will be included.

### Conclusion

A full set of functions for HMAC meeting modern standards has been presented. The functions cover all common needs for the use of HMAC in typical applications written in VBA.

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)