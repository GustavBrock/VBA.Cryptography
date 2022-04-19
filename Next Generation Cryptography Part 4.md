# Storing encrypted data in Microsoft Access

![Title](images/EE%20Title%20Binary.png)

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

## Part 4. Storing hashed or encrypted data

### Alternative storage methods

Hash values and encrypted data are, by nature, not readable, which allows for using storing methods not normally used for plain text. This is because a hashed value or an encrypted text has no code page or character set - it is "raw" *binary* data.

Most database engines - and that includes the JET/ACE engine of Microsoft Access - feature options for storing such data, often to hold large chunks of picture data or complete files. These data types are typically labelled:

- BLOB (Binary Large OBjects), which can hold data chunks up to gigabyte sizes
- Binary, for data of small size

In Microsoft Access, the limits are:

- BLOB (OLE Object field): 1 GB
- Binary field: 510 bytes

Note, that the expected or experienced usage of the Binary field type is so rare, that is has been excluded from the official documentation for [Access specifications](https://support.microsoft.com/en-us/office/access-specifications-0cf3c66f-9cf2-4e32-9568-98c1025bb47c).

However, while these data types are optimal for storing hashed or encrypted data seen from the perspective of the database engine, other factors - like easy copy-paste and transport - may dictate to store these data as *Base64* encoded text in normal *Short Text* or *Long Text* fields. Thus, both approaches will be discussed here.

But, first, the field may need to be appended an existing table, and DAO is the perfect choice when using Microsoft Access.

### Append a storage field

To append a field, suitable for storing either hashed values or encrypted data to an existing table, four functions have been created, one for each data type:

- `AppendShortTextField`
- `AppendLongTextField`
- `AppendShortBinaryField`
- `AppendLongBinaryField`

Here for the two for the text data types:

~~~vb
' Append a short text field to an existing table.
' Generic function.
' Returns True if the short text field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendShortTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    Optional ByVal FieldSize As Long) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbText, FieldSize)
    
    AppendShortTextField = Success
    
End Function
~~~

~~~vb
' Append a long text (Memo) field to an existing table.
' Generic function.
' Returns True if the long text field exists or was created.
'
' 2022-04-01. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendLongTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String) _
    As Boolean
    
    Dim Success     As Boolean
    
    Success = AppendStorageField(Database, TableName, FieldName, dbMemo)
    
    AppendLongTextField = Success
    
End Function
~~~

The two functions for the binary fields are similar and won't be listed here.

Usage is straight, for example:

~~~vb
Dim Success As Boolean

Success = AppendShortTextField(CurrentDb, "YourTable", "YourField", 44)
~~~
Don't forget, that hash values, using the same hashing algorithm, always have the same length which will fit a Short Text field, thus it may be preferable to specify the size of the field - as shown above.

The four functions are this simple, because they all call a common function, `AppendStorageField`, that validates the arguments and does the real work:

~~~vb
' Append to an existing table a field suitable for storing encrypted data.
' Generic function.
' Returns True if the storage field exists as specified or was created.
'
' The field type can be any of these:
'
'   dbBinary        Binary
'   dbLongBinary    Binary (BLOB - Binary Large OBjects)
'   dbText          Short Text
'   dbMemo          Long Text (Memo)
'
' The field size can be specified for Binary and Short Text only:
'
'   Maximum value Binary:   510 (Default)
'   Minimum value Binary:     1
'   Maximum value Binary:   255 (Default)
'   Minimum value text:       1
'
' 2022-03-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function AppendStorageField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal FieldType As DAO.DataTypeEnum, _
    Optional ByVal FieldSize As Long) _
    As Boolean
    
    ' Field size limits.
    Const MaximumSizeBinary As Long = 510
    Const MaximumSizeText   As Long = 255
    Const MinimumSize       As Long = 1
    Const DefaultSize       As Long = 0
    
    Dim Table       As DAO.TableDef
    Dim Field       As DAO.Field
    Dim Success     As Boolean
    
    ' Validate field type.
    Select Case FieldType
        Case dbBinary, dbLongBinary, dbText, dbMemo
            ' OK.
        Case Else
            ' Ignore other field types.
            Exit Function
    End Select
            
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Validate table name.
    For Each Table In Database.TableDefs
        If Table.Name = TableName Then
            Exit For
        End If
    Next
    If Table Is Nothing Then
        ' Table name not found in this database.
        Exit Function
    End If
    
    ' Validate field name.
    For Each Field In Table.Fields
        If Field.Name = FieldName Then
            Exit For
        End If
    Next
    
    If Field Is Nothing Then
        ' Create the field.
        Set Field = Table.CreateField(FieldName, FieldType)
        If FieldSize = DefaultSize Then
            Success = True
        ElseIf FieldSize >= MinimumSize Then
            ' Set the size of the field for relevant field types.
            Select Case FieldType
                Case dbBinary
                    If FieldSize <= MaximumSizeBinary Then
                        Field.Size = FieldSize
                        Success = True
                    End If
                Case dbText
                    If FieldSize <= MaximumSizeText Then
                        Field.Size = FieldSize
                        Success = True
                    End If
            End Select
        End If
        If Success = True Then
            Success = False
            Table.Fields.Append Field
            Success = True
        End If
    ElseIf Field.Type = FieldType Then
        ' The field exists.
        If Field.Size = FieldSize Then
            Success = True
        ElseIf FieldSize = DefaultSize Then
            Select Case Field.Type
                Case dbBinary
                    Success = (Field.Size = MaximumSizeBinary)
                Case dbText
                    Success = (Field.Size = MaximumSizeText)
            End Select
        End If
    End If
    
    Set Field = Nothing
    Set Table = Nothing
    
    AppendStorageField = Success
    
End Function
~~~

> For other database engines, like *SQL Server* and *MySQL*, you will have to use the SQL `ALTER TABLE ...` syntax in a *pass-through* query. That will, however, not be discussed here.

Now, having a text or binary field created, it's time to read and write the data. 

The scenarios may, of course, be many, but here we will imagine, that we have a table with records having unique IDs, and - having an ID - we wish to write/read data to/from this specific record.

First, handling of text data will be described. Next, storing binary data will be described.

### Storing text data

Two functions are at hand for this operation: `SaveTextField` and `ReadTextField`.

The first will write to either a *Short Text* or a *Long Text* field and has the special feature, that it will only return success, if the passed text to be saved actually has been saved, including saving a *Null* for an empty string or, alternatively, an empty string if the field allows either of these options. If the field size is too small to accommodate the text, or the ID was not found, no success will be returned.

Apart from this, it performs a standard field update using DAO:

~~~vb
' Save, for an ID, a text value to a text field of a table in the specified database.
' If argument Text is an empty string, Null will be saved if possible.
' Returns True if success, False if not.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long, _
    ByVal Text As String) _
    As Boolean
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Success             As Boolean
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbText, dbMemo
                If Nz(Records(FieldName).Value) = Text Then
                    ' Nothing to update.
                ElseIf Records(FieldName).Type = dbText And Len(Text) > Records(FieldName).Size Then
                    ' Text is too long to save in this field.
                Else
                    Records.Edit
                    If Text = "" Then
                        If Not Records(FieldName).Required Then
                            Records(FieldName).Value = Null
                        ElseIf Records(FieldName).AllowZeroLength Then
                            Records(FieldName).Value = ""
                        End If
                    Else
                        Records(FieldName).Value = Text
                    End If
                    Records.Update
                End If
                Success = (Nz(Records(FieldName).Value) = Text)
        End Select
    End If
    Records.Close
    
    SaveTextField = Success
    
End Function
~~~

A typical example shows how simple the function will be to implement:

~~~vb
' Save, for an ID, the AES encrypted value of the text passed to a text field.
' A key must be passed.
' If argument Text is empty, Null will be stored.
' Returns True if success.
' Returns False if no key is passed.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveEncryptedBase64( _
    ByVal Id As Long, _
    ByVal Key As String, _
    Optional ByVal Text As String) _
    As Boolean
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBase64"
    
    Dim EncryptedText       As String
    Dim Success             As Boolean
    
    EncryptedText = Encrypt(Text, Key)
    Success = SaveTextField(CurrentDb, DefaultTableName, DefaultFieldName, Id, EncryptedText)
    
    SaveEncryptedBase64 = Success
    
End Function
~~~

The second will read from either a *Short Text* or a *Long Text* field and has the special feature, that it will always return a string, though this will be empty if *Null* is saved or the ID was not found.

Apart from this, it performs a standard field reading using DAO:

~~~vb
' Read, for an ID, the text from a field of a table in the specified database.
' Returns the value as text if the ID is found.
' Returns an empty string, if the ID is not found.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadTextField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long) _
    As String
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Text                As String
    
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset, dbReadOnly)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbMemo, dbText
                ' Read data as is.
                If Not IsNull(Records(FieldName).Value) Then
                    Text = Records(FieldName).Value
                End If
        End Select
    Else
        ' Return an empty string.
    End If
    Records.Close
    
    ReadTextField = Text
    
End Function
~~~

Again, a typical example shows how simple the function will be to implement:

~~~vb
' Read, for an ID, the encrypted text from a text field.
' Returns the decrypted text if the key is right.
' Returns an empty string, if the ID is not found, or the key is wrong.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadEncryptedBase64( _
    ByVal Id As Long, _
    ByVal Key As String) _
    As String
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBase64"
    
    Dim EncryptedText       As String
    Dim Text                As String
    
    EncryptedText = ReadTextField(CurrentDb, DefaultTableName, DefaultFieldName, Id)
    Text = Decrypt(EncryptedText, (Key))
    
    ReadEncryptedBase64 = Text
    
End Function
~~~

The two example functions are intended to serve as skeletons for top-level functions reading and writing *Base64* encoded encrypted text (as described in Part 3 of this series) and should be easy to modify for any practical implementation.

### Storing binary data

Two functions are at hand for this operation: `SaveBinaryField` and `ReadBinaryField`.

The first will write to either a *Binary* or a *Long Binary* (OLE Object) field and has the special feature, that it will only return success, if the passed data to be saved actually has been saved, including saving a *Null* for an empty byte array or, alternatively, an empty array if the field allows either of these options. If the field size is too small to accommodate the data, or the ID was not found, no success will be returned.

Apart from this, it performs a standard field update using DAO:

~~~vb
' Save, for an ID, a byte array to a binary field of a table in the specified database.
' If argument Data() is an empty array, Null will be saved if possible.
' Returns True if success, False if not.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long, _
    ByRef Data() As Byte) _
    As Boolean
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Skip                As Boolean
    Dim Success             As Boolean
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbBinary, dbLongBinary
                If IsNull(Records(FieldName).Value) And StrPtr(Data) = 0 Then
                    ' Nothing to update.
                    Skip = True
                ElseIf StrPtr(Data) = 0 Then
                    ' Set field value to Null.
                ElseIf Records(FieldName).Type = dbBinary And (UBound(Data) - LBound(Data) + 1 > Records(FieldName).Size) Then
                    ' Array is too large to save in this field.
                    Skip = True
                End If
                
                If Not Skip Then
                    Records.Edit
                    If StrPtr(Data) = 0 Then
                        If Not Records(FieldName).Required Then
                            Records(FieldName).Value = Null
                        End If
                    Else
                        Records(FieldName).Value = Data
                    End If
                    Records.Update
                End If
                
                If IsNull(Records(FieldName).Value) And StrPtr(Data) = 0 Then
                    Success = True
                ElseIf UBound(Records(FieldName).Value) - LBound(Records(FieldName).Value) = UBound(Data) - LBound(Data) Then
                    Success = True
                End If
        End Select
    End If
    Records.Close
    
    SaveBinaryField = Success
    
End Function
~~~

A typical example shows how simple the function using the function `EncryptData` will be to implement:

~~~vb
' Save, for an ID, the AES encrypted value of the text passed to a binary field.
' A key must be passed.
' If argument Text is empty, Null will be stored.
' Returns True if success.
' Returns False if no key is passed.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SaveEncryptedBinary( _
    ByVal Id As Long, _
    ByVal Key As String, _
    Optional ByVal Text As String) _
    As Boolean
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBinary"
    
    Dim EncryptedData()     As Byte
    
    Dim Success             As Boolean
    
    If EncryptData((Text), (Key), EncryptedData) = True Then
        Success = SaveBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id, EncryptedData)
    End If
    
    SaveEncryptedBinary = Success
    
End Function
~~~

The second will read from either a *Binary* or a *Long Binary* (OLE Object) field and has the special feature, that it will always return an array, though this will be empty if *Null* is saved or the ID was not found.

> The function is also capable of reading data from a text field as these data will be casted to a byte array. However, for the current topic, encryption of data, this option doesn't offer any advantages and will not be discussed further.

Apart from this, it performs a standard field reading using DAO:

~~~vb
' Read, for an ID, the (casted) byte array from a field of a table in the specified database.
' Returns the value as a byte array if the ID is found.
' Returns an empty byte array, if the ID is not found.
'
' 2022-04-02. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadBinaryField( _
    ByVal Database As DAO.Database, _
    ByVal TableName As String, _
    ByVal FieldName As String, _
    ByVal Id As Long) _
    As Byte()
    
    Dim Records             As DAO.Recordset
    
    Dim Sql                 As String
    Dim Data()              As Byte
    
    ' Validate database and object names.
    If Database Is Nothing Or TableName = "" Or FieldName = "" Then
        ' Nothing to do. Exit.
        Exit Function
    End If
    
    ' Wrap table and field names in brackets.
    TableName = Replace("[{0}]", "{0}", TableName)
    FieldName = Replace("[{0}]", "{0}", FieldName)
    
    Sql = "Select " & FieldName & " From " & TableName & " Where Id = " & Id & ""
    Set Records = CurrentDb.OpenRecordset(Sql, dbOpenDynaset, dbReadOnly)
    
    If Records.RecordCount = 1 Then
        Select Case Records(FieldName).Type
            Case dbBinary, dbLongBinary, dbMemo, dbText
                ' Read data as is.
                If Not IsNull(Records(FieldName).Value) Then
                    Data = Records(FieldName).Value
                End If
        End Select
    Else
        ' Return an empty array.
    End If
    Records.Close
    
    ReadBinaryField = Data
    
End Function
~~~

Again, a typical example shows how simple the function will be to implement, here using the function `DecryptData`:

~~~vb
' Read, for an ID, the encrypted text from a  stored hash value of the password.
' Returns a byte array if a hash value is found.
' Returns an empty byte array, if the ID is not found, or the password is empty.
'
' 2022-04-05. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ReadEncryptedBinary( _
    ByVal Id As Long, _
    ByVal Key As String) _
    As String
    
    ' Table and field names. Modify as needed.
    Const DefaultTableName  As String = "Library"
    Const DefaultFieldName  As String = "ContentBinary"
    
    Dim EncryptedData()     As Byte
    Dim DecryptedData()     As Byte
    Dim Text                As String
    
    EncryptedData = ReadBinaryField(CurrentDb, DefaultTableName, DefaultFieldName, Id)
    
    If DecryptData(EncryptedData, (Key), DecryptedData) Then
        ' Return the decrypted text.
        Text = DecryptedData
    Else
        ' Return an empty string.
    End If
    
    ReadEncryptedBinary = Text
    
End Function
~~~

The two example functions are intended to serve as skeletons for top-level functions reading and writing binary encrypted data (as described in Part 3 of this series) and should be easy to modify for any practical implementation.

### Queries and encrypted data

While the above functions are well suited for single-record reading and writing using VBA, they are less useful in queries where you also - when reading data - must prepare for *Null* values.

To meet this demand, four functions are provided for usage in queries:

- `VEncryptBase64`
- `VEncryptBinary`
- `VDecryptBase64`
- `VDecryptBinary`

They are all listed here in full without further comments, as the in-line comments for each function fully describe the validating of the arguments and the typical usage in a query.

~~~vb
' Encrypt and Base64 encode a text directly to a Short Text or Long Text table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the encrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage (update query):
'
'   Parameters
'       TableNameId Long,
'       [Text] LongText,
'       [Key] LongText;
'   Update
'       TableName
'   Set
'       TableName.[Date] = Date(),
'       TableName.Content = [Text],
'       TableName.ContentBase64 = VEncryptBase64([Text],[Key])
'       Where TableNameId = [Id];
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VEncryptBase64( _
    ByVal Text As Variant, _
    Key As Variant) _
    As Variant

    Dim EncryptedText   As Variant
    
    If Nz(Key, "") = "" Or Nz(Text, "") = "" Then
        ' Either no key or no data.
        EncryptedText = Null
    Else
        ' Return encrypted data as Base64 encoded text.
        EncryptedText = Encrypt(Text, Key)
        If EncryptedText = "" Then
            ' Encryption failed.
            EncryptedText = Null
        End If
    End If
    
    VEncryptBase64 = EncryptedText
    
End Function
~~~
~~~vb
' Encrypt a text directly to a Binary or Long Binary (BLOB) table field
' using a passed key.
' Accepts Null values for both text and key.
'
' If success, the encrypted data is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
' Typical usage (update query):
'
'   Parameters
'       TableNameId Long,
'       [Text] LongText,
'       [Key] LongText;
'   Update
'       TableName
'   Set
'       TableName.[Date] = Date(),
'       TableName.Content = [Text],
'       TableName.ContentBinary = VEncryptBinary([Text],[Key])
'       Where TableNameId = [Id];
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VEncryptBinary( _
    ByVal Text As Variant, _
    Key As Variant) _
    As Variant

    Dim EncryptedData() As Byte
    Dim Data            As Variant
    
    If Nz(Key, "") = "" Or Nz(Text) = "" Then
        ' Either no key or no data.
        Data = Null
    ElseIf VarType(Key) <> vbString Or VarType(Text) <> vbString Then
        ' Key or Text is not text.
        Data = Null
    Else
        ' Arguments are valid. Attempt to encrypt the text.
        If EncryptData((Text), (Key), EncryptedData) = True Then
            ' Success.
            Data = EncryptedData
        Else
            ' Decryption failed.
            Data = Null
        End If
    End If
    
    VEncryptBinary = Data
    
End Function
~~~
~~~vb
' Decrypt a Base64 encoded encrypted text directly from a Short Text or Long Text table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the decrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
'   Parameters [Key] Text ( 255 );
'   Select *, VDecryptBase64([TextFieldName], [Key]) As Content
'   From TableName
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecryptBase64( _
    ByVal EncryptedText As Variant, _
    Key As Variant) _
    As Variant

    Dim Text        As Variant
    
    If Nz(Key, "") = "" Or Nz(EncryptedText, "") = "" Then
        ' Either no key or no data.
        Text = Null
    Else
        ' Return decrypted data as text.
        Text = Decrypt(EncryptedText, Key)
        If Text = "" Then
            ' Decryption failed.
            Text = Null
        End If
    End If
    
    VDecryptBase64 = Text
    
End Function
~~~
~~~vb
' Decrypt an encrypted text directly from a Binary or Long Binary (BLOB) table field
' using a passed key.
' Accepts Null values for both data and key.
'
' If success, the decrypted text is returned.
' If no data, no key, or wrong key, Null is returned.
'
' Typical usage:
'
'   Parameters [Key] Text ( 255 );
'   Select *, VDecryptBinary([BinaryFieldName], [Key]) As Content
'   From TableName
'
' 2022-04-07. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function VDecryptBinary( _
    ByVal EncryptedData As Variant, _
    Key As Variant) _
    As Variant

    Dim Data()      As Byte
    Dim Text        As Variant
    
    If Nz(Key, "") = "" Or IsNull(EncryptedData) Then
        ' Either no key or no data.
        Text = Null
    ElseIf VarType(Key) <> vbString Or VarType(EncryptedData) <> vbString Then
        ' Key is not text, or - as a binary field will be read as text - ,
        ' EncryptedData holds invalid data.
        Text = Null
    Else
        ' Arguments are valid. Attempt to decrypt the data.
        DecryptData (EncryptedData), (Key), Data
        If StrPtr(Data) = 0 Then
            ' Decryption failed.
            Text = Null
        Else
            ' Return decrypted data as text.
            Text = CStr(Data)
        End If
    End If
    
    VDecryptBinary = Text
    
End Function
~~~

Three queries are included in the demo application:

- `UpdateBase64`
- `UpdateBinary`
- `LibraryContent`

The two first are queries where you need to pass data as parameters:

~~~sql
PARAMETERS 
    LibraryId Long, 
    [Text] LongText, 
    [Key] LongText;
UPDATE 
    Library 
SET 
    Library.[Date] = Date(), 
    Library.ContentBase64 = VEncryptBase64([Text],[Key]), 
    Library.Content = [Text]
WHERE 
    LibraryId=[Id];
~~~
~~~sql
PARAMETERS 
    LibraryId Long, 
    [Text] LongText, 
    [Key] LongText;
UPDATE 
    Library 
SET 
    Library.[Date] = Date(), 
    Library.ContentBinary = VEncryptBinary([Text],[Key]), 
    Library.Content = [Text]
WHERE 
    LibraryId=[Id];
~~~

The third lists the sample entries:

~~~sql
PARAMETERS 
    [Key] LongText;
SELECT 
    Library.Id, 
    Library.Date, 
    VDecryptBase64([ContentBase64],[Key]) AS Base64Content, 
    VDecryptBinary([ContentBinary],[Key]) AS BinaryContent
FROM 
    Library;
~~~

Note that, when you run this query, the content fields are blank if you enter the wrong or no key.

> For all queries, while working with the demo, use a **single space** as key.

### Conclusion

A full set of functions for reading and writing encrypted data as Base64 encoded text or as binary data has been presented. The functions cover all common needs for storing hash values or encrypted data in typical applications written in VBA.

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)