Attribute VB_Name = "XORCrypto"
'Original Author or Source: VBxtras' VBHowTo Issue #33 - January 10, 2000
'Demonstrates XOR encryption and decryption in a unique way

Option Explicit

Const ENCRYPT_OFFSET As Integer = 3
Const ENCRYPT_BASE As Integer = 7

Function EncryptString(ByVal sSource As String) As String
Dim sEncrypted As String
Dim nLength As Long
Dim nLoop As Long
Dim nTemp As Integer

    nLength = Len(sSource)
    sEncrypted = Space$(nLength)
    For nLoop = 1 To nLength
        nTemp = Asc(Mid$(sSource, nLoop, 1))
        If nLoop Mod 2 Then
            nTemp = nTemp - ENCRYPT_OFFSET
        Else
            nTemp = nTemp + ENCRYPT_OFFSET
        End If
        nTemp = nTemp Xor (ENCRYPT_BASE - ENCRYPT_OFFSET)
        Mid$(sEncrypted, nLoop, 1) = Chr$(nTemp)
    Next
    EncryptString = sEncrypted
End Function

Function DecryptString(ByVal sSource As String) As String
Dim sDecrypted As String
Dim nLength As Long
Dim nLoop As Long
Dim nTemp As Integer

    nLength = Len(sSource)
    sDecrypted = Space$(nLength)
    For nLoop = 1 To nLength
        nTemp = Asc(Mid$(sSource, nLoop, 1)) Xor (ENCRYPT_BASE - ENCRYPT_OFFSET)
        If nLoop Mod 2 Then
            nTemp = nTemp + ENCRYPT_OFFSET
        Else
            nTemp = nTemp - ENCRYPT_OFFSET
        End If
        Mid$(sDecrypted, nLoop, 1) = Chr$(nTemp)
    Next
    DecryptString = sDecrypted
End Function
