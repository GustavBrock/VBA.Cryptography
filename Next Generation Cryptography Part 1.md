# Utilise Microsoft's Next Generation Cryptography (CNG) API in VBA

![Title](images/EE%20Title%20Cryptography.png)

This article series will show you how to utilise the *Next Generation Cryptography (CNG)* API from Microsoft for modern *hashing* and *encrypting/decrypting*, maintaining 100% compatibility with the implementation of CNG in *.Net* as well as *PowerShell* scripts.

---
## Introduction

### Purpose of these articles

The Microsoft CNG APIs constitute a collection of more than a dozen APIs that handle all the aspects and supporting functions to calculate hash values and perform encryption and decryption meeting modern high demands and standards.

This series details one method to implement these features in applications supported by VBA (Visual Basic for Applications), for example *Microsoft Access* and *Microsoft Excel*.

> The theory for and the clever mathematics behind hashing and encryption will not be touched; this is covered massively in books and articles published over a long period of time by dedicated experts. A search with Bing or Google will return a long list of resources to study.

### Sections

The series has been split in five parts. This allows you to skip parts you are either familiar with or wish to implement later if at all.

1. [Utilise Microsoft's Next Generation Cryptography (CNG) API in VBA](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%201.md)
2. [Hashing in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%202.md)
3. [Encryption in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%203.md)
4. [Using binary storage to serve the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%204.md)
5. [Storing passwords in VBA using the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/VBA.Cryptography/blob/main/Next%20Generation%20Cryptography%20Part%205.md)

The three first deal, as stated, with hashing and encrypting/decrypting based on the CNG API, while the forth explains how to combine these techniques with the little known feature of Access, *storing binary data*, as this in many cases will represent the optimal storage method for hashed or encrypted data.
The last demonstrates how to save and verify passwords totally safe using these tools.

## Part 1. Compatible with other environments

However, as comprehensive as the Microsoft CNG API is, as is the complexity. Thus, it takes a lot of work and good understanding of cryptography to implement this API in high-level functions that are easy to apply in projects and applications.

In the **.Net framework**, the complexity of the API has been carefully hidden to make it quite simple and fast to, say, calculate a hash value:

~~~cs
// Input.
string text = "Get your filthy hands off my desert.";
// Set hash algorithm.
var hashMethod = new System.Security.Cryptography.SHA256Managed();

// Convert the string to a byte array.
byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
// Compute the hash value.
byte[] hashBytes = hashMethod.ComputeHash(textBytes);
// Convert the hash value to a Base64 encoded string
string hash = Convert.ToBase64String(hashBytes);

// Output:
Console.WriteLine(hash);
// AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc=
~~~

It hardly can't be simpler, not even in **PowerShell**:

~~~ps1
# Text to hash.
$text="Get your filthy hands off my desert."
# Hash algorithm to use.
$algorithm="SHA256"

# Convert text to bytes.
$data=[System.Text.Encoding]::Unicode.GetBytes($text)
# Calculate hash of text.
$hash=[System.Security.Cryptography.HashAlgorithm]::Create($algorithm).ComputeHash($data)
# Convert hash to Base64.
$value=[Convert]::ToBase64String($hash)

# Display result.
"Text input: " + $text
"Algorithm : " + $algorithm
"Hash value: " + $value

# Output:
# Text input: Get your filthy hands off my desert.
# Algorithm : SHA256
# Hash value: AIPgWDlQLv7bvLdg7Oa78dyRbC0tStuEXJRk0MMehOc=
~~~

However, VBA knows nothing about the CNG API and probably never will, thus - to reach a similar level of simplicity - the included modules in this project will serve to help you to achieve this in VBA.

As the included top level hash function uses the SHA256 hashing method by default and returns a Base64 encoded result, the VBA code can be extremely tight:

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

As you can see, the resulting hash is identical for all three environments, thus allows for easy interchange of data between different applications and systems.

## The building blocks

The module holding all the functions needed to wrap the CNG API is `BCrypt`. As the list of functions is long, it may not be easy to navigate and determine which function(s) to use for which purpose.

A map will help to clarify the dependencies between the functions. Top level is at left:

~~~
Encrypt
    EncryptData
        HashData
            CngHash
        RandomData
            CngRandom
        CngEncrypt
    ByteBase64

Decrypt
    Base64Bytes
    DecryptData
        HashData
            CngHash
        CngDecrypt
        CngHash

Hash
    HashData
        CngHash
    ByteBase64

Random
    RandomData
        CngRandom
    ByteBase64
~~~

The four columns (or levels) serves a variety of implementations.

- the first column holds the main (top level) functions using text (plain or Base64 encoded) for input and output. These wrap the functions of the second column
- the second column holds the functions (named with the suffix "Data") using byte arrays for input and output. They wrap the functions that handles the CNG API. Also listed are the functions used to Base64 encode/decode plain text
- the third and fourth columns hold the functions that wrap the CNG API. These are prefixed "Cng"

To sum it up:

- for handling *text data*, the main functions (at left) are used
- for handling *binary data* (byte arrays), the `*Data` functions are used
- for *low-level operations* with the CNG API, the `Cng*` functions are used

### Supporting functions

Not listed above are a dozen of low-level supporting and supplemental functions.

Among these is a full set of functions for *Base64* encoding and decoding of text, which are neither specific nor mandatory for encrypting and hashing but, nevertheless, widely used for storing and transport of encrypted data.

Also, as encrypted text takes up more volume than plain text, are functions included to calculate either how large an encrypted text will be or, vice versa, how much plain text that - when encrypted - can be accommodated in a given space like a table field of limited size.

The first is `EncryptedTextLength`:

~~~vb
' Return the length of a string encrypted and Base64 encoded with function Encrypt.
'
' To be called from function FitTextField.
'
' Example:
'   Text = "Careful with that axe, Eugene!"
'   DecryptedTextLength = Len(Text)     ' = 30
'   Length = EncryptedTextLength(DecryptedTextLength)
'   Length -> 152
'
' Example data:
'
' Length plain  Length encrypted
'   0             0
'   1            64
'   4            88
'  12           108
'  20           128
'  28           152
'  36           172
'  44           192
'  52           216
'  60           236
'  67           236
' 67 characters is the largest string to encrypt, if the
' encrypted string must fit a Short Text field of Access.
'  68           256
'  76           280
'  84           300
'  92           320
' 100           344
' 108           364
' 116           384
' 124           408
' 132           428
' 140           448
' 148           472
' 156           492
' 164           512
' 172           536
' 180           556
' 188           576
' 196           600
' 204           620
' 212           640
' 220           664
' 228           684
' 236           704
' 244           728
' 252           748
' The maximum length of an Access Short Text field is 255 characters.
' 255           748
' 260           768
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function EncryptedTextLength( _
    ByVal DecryptedTextLength As Long) _
    As Long

    Dim Length      As Long

    If DecryptedTextLength > 0 Then
        Length = 64 - Int(-(DecryptedTextLength - 3) / 8) * 20 - Int(-(DecryptedTextLength - 3) / 24) * 4
    End If

    EncryptedTextLength = Length

End Function
~~~

The second is `DecryptedTextLength`:

~~~vb
' Return for an available length of an encrypted and Base64 encoded string
' the possible maximum length of a string to be encrypted.
'
' To be called from function FitTextField.
'
' Examples:
'   Available encrypted     Length plain maximum
'     0                       0
'    63                       0
'    64                       3
'    87                       3
'    88                      11
'   107                      11
'   108                      19
'   127                      19
'   128                      27
'   151                      27
'   152                      35
'   171                      35
'   172                      43
'   191                      43
'   192                      51
'   215                      51
'   216                      59
'   235                      59
' Largest option for a Short Text field of Access
' holding the decrypted string:
'   236                      67
'   255                      67
'   256                      75
'   747                     251
' Largest option for a Short Text field of Access
' holding the encrypted string:
'   748                     259
'
' 2021-10-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DecryptedTextLength( _
    ByVal EncryptedTextLength As Long) _
    As Long

    Dim Length      As Long
    Dim Length64    As Long
    Dim Delta20     As Long
    Dim Delta       As Long

    Select Case EncryptedTextLength
        Case Is < 64
            Length = 0
        Case Is < 64 + 24
            Length = 3
        Case Else
            Length64 = -Int(-EncryptedTextLength / 64)
            Delta20 = -Int(-((Length64 * 64 - EncryptedTextLength) / 20))
            Delta = (Delta20 \ 4) * 3 + Delta20 Mod 4
            Length = (Length64 - 1) * 24 - (Delta * 8) - 4 + 7
    End Select

    DecryptedTextLength = Length

End Function
~~~

For both functions, documentation is included in-line. If you wonder about the math used for the calculations, it is purely empiric, reduced to an acceptable minimum to keep it readable.

As you can see, storing encrypted text raises some challenges compared to storing plain text. These will be discussed in Part 4: [Using binary storage to serve the Microsoft NG Cryptography (CNG) API](https://github.com/GustavBrock/Veeam.Linux/blob/main/Linux%20Repository%203.md).

Similar functions for binary storage are also included. Further, and much simpler, is a function, `HashTextLength`, to return the sizes of hash values. These are fixed, determined by the hash algorithm:

~~~vb
' Return the fixed length of a Base64 encoded hash from function Hash
' using the specified hash algorithm.
'
' Default length is 44.
' Maximum length is 88.
' Minimum length is 24.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashTextLength( _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Integer

    Dim Length  As Integer

    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 24
        Case bcSha1
            Length = 28
        Case bcSha256
            Length = 44
        Case bcSha384
            Length = 64
        Case bcSha512
            Length = 88
    End Select

    HashTextLength = Length

End Function
~~~

A similar function for binary storage is also included, `HashByteLength`:

~~~vb
' Return the byte count of a hash from function HashData.
'
' Default count is 32.
' Maximum count is 64.
' Minimum count is 16.
'
' 2022-02-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HashByteLength( _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Integer

    Dim Length  As Integer
    
    Select Case BcryptHashAlgorithmId
        Case bcMd2, bcMd4, bcMd5
            Length = 16
        Case bcSha1
            Length = 20
        Case bcSha256
            Length = 32
        Case bcSha384
            Length = 48
        Case bcSha512
            Length = 64
    End Select
    
    HashByteLength = Length
    
End Function
~~~

As you can see, storing hash values binary will save some space. However, the values are small and text values can easily be held in a *Short Text* field. Thus, the difference in required storage volume should hardly be what determines what data type to use for storing hash values.

### Conclusion

A top-level walk-through of the organisation of the building blocks for encrypting and hashing using the *Next Generation Cryptography* as well as the most important supporting functions has been presented.

The implementation of these as well as storing encrypted or hashed values will be presented in the next parts of this series.

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)