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

### Credit

Based on code at [Stack Overflow](https://stackoverflow.com/questions/67294035/basic-encrypting-of-a-text-file/67294779#comment122708972_67294779) published 2021-04-28 (with later edits) by Stack Overflow user [Erik A](https://stackoverflow.com/users/7296893/erik-a).

### Documentation and demo

Official CNG API documentation:

[Cryptography API: Next Generation](https://docs.microsoft.com/en-us/windows/win32/seccng/cng-portal)


Top level code documentation generated by [MZ-Tools](https://www.mztools.com/) is included for [Microsoft Access](https://htmlpreview.github.io?https://github.com/GustavBrock/VBA.Cryptography/blob/master/documentation/CngCrypt.htm).

Detailed documentation is in-line. 

Full documentation can also be found here:

![EE Logo](images/EE%20Logo.png)

[Next Generation Cryptography](https://www.experts-exchange.com/articles/37111/Utilise-Microsoft's-Next-Generation-Cryptography-CNG-API-in-VBA.html)

Demo applications for *Microsoft Access* and *Microsoft Excel* tested with *Microsoft 365* are located in the *demos* folder. 

---

*If you wish to support my work or need extended support or advice, feel free to:*

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.Cryptography/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)