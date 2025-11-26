# Storing passwords in VBA using the Microsoft NG Cryptography (CNG) API

![Title](images/EE%20Title%20Passwords.png)

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

## Part 5. Handling passwords

### General concern

This statement you have seen and will see over and over again. And it is true, as you will not want someone - not even with legitimate access to the database - to be able to read users' password by any mean:

> **Never, ever store passwords as plain text.**

Thus, you must take action to store passwords encrypted or, both simpler and better, to store only hash values of the passwords.

- The reason for "simpler" is, that no key is required, neither for storing or validating a password.
- The reason for "better" is, that hash values are unidirectional or non-reversible, meaning that it is not possible to reverse engineer a password from its hash value.

Thus, saving and comparing hash values of passwords is considered to be the safe method for "storing" and validating passwords. This means, that saving a password is done by saving its hash value, and checking a password is done by comparing the hash value of the entered password with the stored hash value; if the two values match, the entered password is correct.

The only concern is the algorithm to be used to create the hash values. This has been discussed in Part 2 of this series, so study that for the details. Here we will proceed using the *SHA256* algorithm, which everywhere in this series is used as the default algorithm.

### Requirements

This part of the article series makes use of the functions (building blocks) from the previous parts. Thus, these functions will no be listed or explained here. So, if you haven't read the previous sections, at least read Part 1, 2, and 4.

### Prepare for storing hash values of passwords

The direct and optimum method for storing the hash value of a password, is to save the value as a byte array in a binary field. 

For this, an optimised function has been created, that will append a suitable field to a table in a database. By default, the field will be named "Password" and it will be dimensioned to fit hash values created with the widely used *SHA256* algorithm:

~~~vb
' Append, to an existing DAO table, a binary field optimised for
' storing a BCrypt hash value using a hash algorithm as specified by
' the argument BcryptHashAlgorithmId.
' Returns True if the field exists or was created.
'
' By default, the size of the field will be set to match SHA256.
'
' 2022-02-14. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendPasswordField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    Optional ByVal FieldName As String = "Password", _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Boolean
    
    Dim FieldSize   As Integer
    Dim Success     As Boolean
    
    ' Find the required field size for this hash algorithm.
    FieldSize = HashByteLength(BcryptHashAlgorithmId)
    Success = AppendBinaryField(Database, TableName, FieldName, FieldSize)
    
    AppendPasswordField = Success

End Function
~~~

The default usage is very simple, though the table will typically be located in another database than the current, for example:

~~~vb
Public Function TestAppendPasswordField()

    Dim Database    As DAO.Database
    Dim Success     As Boolean
    
    Set Database = DBEngine(0).OpenDatabase("C:\Test\Backend.accdb")
    Success = AppendPasswordField(Database, "User")
    
    Database.Close
    Set Database = Nothing
    
    TestAppendPasswordField = Success

End Function
~~~

Having the field for the hash values, it's time to save and verify these. 

### Storing hash values of passwords

The main tools for this operation, are the functions `HashData` and `SaveBinaryField` for creating and saving the hash value of a password as a byte array. The only arguments needed are the ID of the user and the entered password. Optionally, the hash algorithm to use can be specified if different from the default *SHA256*.

> Two (or more) users might use the same password, leaving identical hash values. This means that, if someone gains access to the table, having one user's password could reveal the password of another user.
>
> To prevent this, the entered password is "salted" with the unique ID of the user.

Note, that only byte arrays are used to hold the data:

~~~vb
' Save, for an ID, the hash value of the password passed in a
' binary field of a DAO table.
' The hash value will be salted with the unique Id.
' If argument Password is empty, the hash value will be reset (no password).
' Returns True if success.
'
' By default, the hash algorithm SHA256 is applied.
'
' 2022-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SavePassword( _
    ByVal Id As Long, _
    Optional ByVal Password As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Boolean
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "User"
    Const DefaultFieldName  As String = "Password"
    
    Dim Data()      As Byte
    Dim TextData()  As Byte
    Dim Success     As Boolean
    
    If Password = "" Then
        ' Reset saved password.
    Else
        ' Salt the password with the unique Id.
        TextData = Password & CStr(Id)
    End If
    
    Data = HashData(TextData, BcryptHashAlgorithmId)
    Success = SaveBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id, Data)
    
    SavePassword = Success
    
End Function
~~~

An important detail is, that an empty password is not allowed. Doing so will reset the hash value by setting the field value to *Null*.

This calls for a simple top-level function to reset the password for a user:

~~~vb
' Reset, for an ID, the hash value of a password stored in a
' binary field of a DAO table.
' Returns True if success.
'
' 2022-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ResetPassword( _
    ByVal Id As Long) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = SavePassword(Id)
    
    ResetPassword = Success
    
End Function
~~~

If you view the table, you'll see a strange display of the hash values as Access tries to read them as unicode characters:

![User table with password field](images/Table%20User.png)

Here, user Olav has no password recorded.

### Read and verify passwords

To verify an entered password, first the stored hash value of this must be read. A function for this already exists, `xxx`, and it can be used as in this example:

~~~vb
' Read, for an ID, the stored hash value of the password.
' To be used to verify a password.
' Returns a byte array if a hash value is found.
' Returns an empty byte array, if the ID is not found, or the password is empty.
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadPassword( _
    ByVal Id As Long) _
    As Byte()
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "User"
    Const DefaultFieldName  As String = "Password"
    
    Dim Data()      As Byte
    
    Data = ReadBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id)
        
    ReadPassword = Data
    
End Function
~~~

Next, having the byte array with hash value, this must be compared with the hash value of the entered password.

The full process is held in the function `VerifyPassword`, that utilises the above function and *True* for a match:

~~~vb
' Verify, for an ID, a password (salted with the unique Id) by comparing the
' hash value, using the specified hash algorithm, with the stored hash value.
' Returns True for a match.
'
' By default, the hash algorithm SHA256 is applied.
'
' 2022-04-08. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VerifyPassword( _
    ByVal Id As Long, _
    Optional ByVal Password As String, _
    Optional ByVal BcryptHashAlgorithmId As BcHashAlgorithm = bcSha256) _
    As Boolean
    
    Dim Data()      As Byte
    Dim TextData()  As Byte
    Dim Success     As Boolean
    
    Data = ReadPassword(Id)
    If Password = "" Then
        ' The stored hash of an empty password must be empty.
        Success = Not CBool(StrPtr(Data))
    Else
        ' Compare the stored hash value of the password with the hash value of the
        ' passed password (salted with the unique Id) using the specified hash algorithm.
        TextData = Password & CStr(Id)
        Success = Not CBool(StrComp(Data, HashData(TextData, BcryptHashAlgorithmId), vbBinaryCompare))
    End If
    
    VerifyPassword = Success
    
End Function
~~~

Note, that if an empty password is passed, the stored value must be empty for a match. This can be used to create a function that serves the common task to check if a password exists for a user:

~~~vb
' Verify, for an ID, that a hash value for a password is saved.
' Returns True if the ID exists and has a hash value for a password stored.
' Returns False if the ID does not exist or it has no password.
'
' 2022-02-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function PasswordExists( _
    ByVal Id As Long) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = Not VerifyPassword(Id)
        
    PasswordExists = Success
    
End Function
~~~

### Conclusion

With the functions listed here and in the previous parts of this series, it has been shown how to handle passwords safely using the *Next Generation Cryptography*, thus fulfilling the highest standards and the demands for users' privacy.

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)