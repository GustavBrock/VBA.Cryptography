# Hashing in VBA using the Microsoft NG Cryptography (CNG) API

![Title](images/EE%20Title%20Hashing.png)

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

## Part 2. Hashing

### Basic usage

Hashing, in this context, is the process of creating a non-reversible fixed length code that identifies the input.

Note, that the code is not unique as, in theory, the same code could be generated from different inputs, though - due to the extreme count of combinations for a given string length - this is unlikely to happen in real life.

This, however, is neither the point. The point is to be able to create a code based on an input, (later) to recreate that code, and then to check if the later code matches the first; if it does, the later input is assumed to match the original input.

The typical usage of hash values is for "storing" users' passwords. Instead of storing the password, the hash value is stored. To verify the password, its hash value is generated and compared with the stored hash value. If these values match, the password is accepted.

### Hashing methods

The mathematical method used to hash a string is called an algorithm. These have evolved over time, and the oldest - also producing the shortest hash values - are now deprecated, though still widely used for less demanding purposes. For example, **SHA1**, and even the very old **MD5**, are often used when downloading files for verification of the integrity of the downloaded file. The expected hash value of the file is listed, and this must match the calculated hash value of the actual downloaded file.

The algorithms supported and their byte lengths are:

| Name  | Byte Length |
| ----- | ------ |
| MD2   |     16 |
| MD4   |     16 |
| MD5   |     16 |
| SHA1  |     20 |
| SHA256|     32 |
| SHA384|     48 |
| SHA512|     64 |

While **MD5** and **SHA1** can be well suited for some purposes, the general consensus is, that **SHA256** is the minimum method to use for hashing passwords. For higher demands, use **SHA512** or **SHA384**, but be aware, that **SHA384** is not supported in all environments.

### Working with hash values

Note, that the lengths listed above for the hash algorithms are the lengths of the generated *byte arrays*. As these bytes can take many other values than those for normal ASCII characters - even control characters and other non-printable characters - the byte array must for many practical purposes be converted to something printable.

The simple method would be to convert to the hexadecimal representation of the bytes, for example four bytes (here character `<tab>` is the third):

    aX<tab>! -> 6158092D

but that consumes *twice as many characters*, thus a **SHA256** hash would have a length of 64 characters - not very efficient.

A better method is to use *Base64* encoding. It offers the same - a printable and portable string of ASCII characters - but consumes only 50% or less additional space then the byte count:

| Name  | Base64 Length |
| ----- | ------ |
| MD2   |     24 |
| MD4   |     24 |
| MD5   |     24 |
| SHA1  |     28 |
| SHA256|     44 |
| SHA384|     64 |
| SHA512|     88 |

To conclude: For only a little overhead, the hash values can be treated as normal plain text.

### Creating a hash

A set of functions is used to create hash values:

~~~txt
Hash
    HashData
        CngHash
    ByteBase64
~~~

The top function, `Hash`, the one that creates a Base64 encoded hash value from a passed text, contains less than a handful of code lines:

~~~vb
' Return a Base64 encoded hash of a string using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' Example:
'   Text = "Get your filthy hands off my desert."
'   Value = Hash(Text)
'   Value -> "AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc="
'
' Length of the generated Base64 encoded hash string:
'
'   Encoding    Length
'   MD2         24
'   MD4         24
'   MD5         24
'   SHA1        28
'   SHA256      44      ' Default.
'   SHA384      64
'   SHA512      88
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Hash( _
    ByVal Text As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As String

    Dim HashBase64          As String

    If Text = "" Then
        ' No data. Nothing to do.
    Else
        HashBase64 = ByteBase64(HashData((Text), BcryptHashAlgorithmId))
    End If

    Hash = HashBase64

End Function
~~~

The only optional argument is which hash algorithm to use; by default, **SHA256** is used.

The usage is extremely simple:

~~~vb
Dim Text    As String
Dim Value   As String

Text = "Get your filthy hands off my desert."
Value = Hash(Text)

Debug.Print "Text input:", Text
Debug.Print "Hash value:", Value

' Output:
' Text input:   Get your filthy hands off my desert.
' Hash value:   AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc=
~~~

The few lines of code in the function is possible, because it pulls data from the second level function, `HashData`, also taking an optional argument for the hash algorithm, which defaults to **SHA256**:

~~~vb
' Create and return the hash of a byte array as another byte array
' using the specified hash algorithm.
' By default, hash algorithm SHA256 is used.
'
' To be called from functions EncryptData, DecryptData, and Hash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashData( _
    ByRef TextData() As Byte, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = BcHashAlgorithm.bcSha256) _
    As Byte()

    Dim Data()          As Byte
    Dim DataLength      As Long

    If StrPtr(TextData) = 0 Then
        ' Array TextData is empty.
    Else
        DataLength = UBound(TextData) - LBound(TextData) + 1
        Data = CngHash(VarPtr(TextData(LBound(TextData))), DataLength, BcryptHashAlgorithmId)
    End If

    HashData = Data

End Function
~~~

The important feature of the function is, that both input and output are *byte arrays*. However, you may have noticed, that function `Hash` above doesn't pass a byte array, but plain text:

~~~vb
HashData((Text), BcryptHashAlgorithmId)
~~~

That's because you can assign a text variable directly to a byte array; no conversion is needed.

The output is, in function `Hash`, fed directly to function `ByteBase64`, which converts the byte array to Base64 encoded text. Though quite convoluted, that function is trivial and won't be discussed here.

Function `HashData` also contains very few code lines. Essentially, it only wraps the function, that does "the real work", `CngHash`, which makes no less than eight calls to the CNG API to obtain the hash value.

The details of these calls will not be discussed here. However, each step is commented in-line, and links to the documentation at Microsoft are included, should you wish to study this further.

The usage of this function is also very straight:

~~~vb
Dim Data()  As Byte
Dim Value() As Byte

Data = "Text to hash"
Value = HashData(Data)

' Convert byte array to unicode.
Debug.Print StrConv(Value, vbUnicode)

' Output.
' ÌêÊž ”Þû§3k¥ábÌ–t1Îôx[Œ’ö‡_šä
~~~

Part 4 of this series deals with various methods for storing hash values, where indeed these two functions can prove useful:

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
    bcSha1 = 4
    bcSha256 = 5
    bcSha384 = 6
    bcSha512 = 7
    [_Last] = 7
End Enum
~~~

This is to validate input and make it easy to always supply the rightly spelled and uppercased hash algorithm name to function `CngHash`.

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
' enum BcHashAlgorithm.
' Note: To validate, HashAlgorithm must be in UPPERCASE.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashAlgorithm( _
    ByVal HashAlgorithm As String) _
    As Boolean

    Dim Index           As BcHashAlgorithm
    Dim Result          As Boolean

    For Index = BcHashAlgorithm.[_First] To BcHashAlgorithm.[_Last]
        If BcryptHashAlgorithm(Index) = HashAlgorithm Then
            Result = True
            Exit For
        End If
    Next

    IsBcryptHashAlgorithm = Result

End Function
~~~

~~~vb
' Return True if the passed value of enum BcHashAlgorithm is valid.
'
' To be called from function CngHash.
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsBcryptHashAlgorithmId( _
    ByVal HashAlgorithmId As BcHashAlgorithm) _
    As Boolean

    Dim Result          As Boolean

    If BcHashAlgorithm.[_First] <= HashAlgorithmId And HashAlgorithmId <= BcHashAlgorithm.[_Last] Then
        Result = True
    End If

    IsBcryptHashAlgorithmId = Result

End Function
~~~

Finally, to determine the length of a hash value - either as a byte array or a Base64 encoded string - without actually hashing something, two simple functions, `HashByteLength` and `HashTextLength`, will return those values. These functions were discussed in Part 1.

These functions can come in handy for code that creates or alters table design where fields for hash values will be included.

### Conclusion

A full set of functions for hashing meeting modern standards has been presented. The functions cover all common needs for hashing in typical applications written in VBA.

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)