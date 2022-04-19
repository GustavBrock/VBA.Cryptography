Attribute VB_Name = "BPassword"
Option Compare Database
Option Explicit
'
' BPassword V1.1.3
' Handling and binary storing of hashed passwords using DAO and the BCrypt API.
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Cryptography
'
' Requires:
'   Module BCrypt
'

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

