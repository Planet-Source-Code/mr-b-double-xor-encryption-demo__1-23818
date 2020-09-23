Attribute VB_Name = "FinalXORCrypt"
'========================================
'Coded by: Mr.B
'If you have any comments or suggestions,
'Please e-mail me at: bleetcs@hotmail.com
'                    or blee@tcs.on.ca
'========================================

'description:
'this module demonstrates XOR encryption in a simple way
'each letter of the plain text is converted into ascii
'then it is encrypted with corresponding letter from the password
'which is turned into ascii was well. finally XOR the two numbers.
'the password is recycled. decryption is simply XORing the encrypted
'string. i believe this is good enough to provide PERSONAL security, not
'something like that could be used in a commercial software. : )
'have fun!
'
'for example, if password is "test" and plain text is "this is a test"
'plain text:    t h i s _ i s _ a _ t e s t (_ = space)
'password:      t e s t t e s t t e s t t e

Option Explicit

Function XOREncrypt(PlainText As String, Password As String) As String
    Dim PTLength As Long        'plain text length
    Dim PWDLength As Integer    'password length
    Dim X As Long               'just a variable
    
    PTLength = Len(PlainText) - 1   'get plain text length. you'll see why we subtract
                                    '1 from the length later. well because i used mod to
                                    'find the corresponding letters
    PWDLength = Len(Password)       'get password length
    XOREncrypt = ""                 'initialize the function
    
    'now loop through the plain text and encrypt it
    For X = 0 To PTLength
        'convert a letter of plain text to ascii
        'XOR it with a letter from password, which is also
        'turned into ascii
        'use mod to find the right corresponding letter from
        'password
        XOREncrypt = XOREncrypt + CStr(Asc(Mid$(PlainText, X + 1, 1)) Xor Asc(Mid$(Password, (X Mod PWDLength) + 1, 1))) + " "
    Next X
    
    'trim it since a space added at the end of the string
    XOREncrypt = Trim$(XOREncrypt)
End Function

Function XORDecrypt(Cipher As String, Password As String) As String
    Dim TempArray As Variant    'we are going to dump the cipher here
                                'using split
    Dim X As Long               'just a variable
    Dim PWDLength As Integer    'length of the password
    
    'i have not came up with error when we have wrong passwords
    'so just resume. the output will be null
    On Error Resume Next
    
    TempArray = Split(Cipher, " ")  'split the cipher into an array
    PWDLength = Len(Password)       'get password length
    XORDecrypt = ""                 'initialize the function
    
    For X = 0 To UBound(TempArray)
        'XOR the cipher with the corresponding letter of password.
        'use of mod is the same as encryption.
        'finally turn the output into string
        XORDecrypt = XORDecrypt + Chr(Int(TempArray(X)) Xor Asc(Mid$(Password, (X Mod PWDLength) + 1, 1)))
    Next X
End Function
