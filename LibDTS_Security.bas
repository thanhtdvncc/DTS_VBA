Attribute VB_Name = "LibDTS_Security"
' Module: LibDTS_Security
' Purpose: Encryption/Decryption for XData protection
' Note: Simple XOR-based encryption for demonstration
Option Explicit

Private Const ENCRYPTION_KEY As String = "DTS_2024_SECURE"

' ==========================================
' ENCRYPTION/DECRYPTION
' ==========================================

' Encrypt a string using XOR encryption
Public Function Encrypt(plainText As String) As String
    On Error GoTo ErrHandler
    
    If Len(plainText) = 0 Then
        Encrypt = ""
        Exit Function
    End If
    
    Dim result As String
    Dim i As Long
    Dim keyLen As Long
    Dim keyChar As Integer
    Dim plainChar As Integer
    
    keyLen = Len(ENCRYPTION_KEY)
    result = ""
    
    For i = 1 To Len(plainText)
        plainChar = Asc(mid$(plainText, i, 1))
        keyChar = Asc(mid$(ENCRYPTION_KEY, ((i - 1) Mod keyLen) + 1, 1))
        result = result & Chr$(plainChar Xor keyChar)
    Next i
    
    ' Base64 encode the result for safe storage
    Encrypt = Base64Encode(result)
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "Encryption error: " & err.description, DTS_ERROR
    Encrypt = plainText ' Return original on error
End Function

' Decrypt a string using XOR decryption
Public Function Decrypt(encryptedText As String) As String
    On Error GoTo ErrHandler
    
    If Len(encryptedText) = 0 Then
        Decrypt = ""
        Exit Function
    End If
    
    ' Base64 decode first
    Dim decoded As String
    decoded = Base64Decode(encryptedText)
    
    ' XOR decrypt (same operation as encrypt)
    Dim result As String
    Dim i As Long
    Dim keyLen As Long
    Dim keyChar As Integer
    Dim encChar As Integer
    
    keyLen = Len(ENCRYPTION_KEY)
    result = ""
    
    For i = 1 To Len(decoded)
        encChar = Asc(mid$(decoded, i, 1))
        keyChar = Asc(mid$(ENCRYPTION_KEY, ((i - 1) Mod keyLen) + 1, 1))
        result = result & Chr$(encChar Xor keyChar)
    Next i
    
    Decrypt = result
    Exit Function
    
ErrHandler:
    LibDTS_Logger.Log "Decryption error: " & err.description, DTS_ERROR
    Decrypt = encryptedText ' Return original on error
End Function

' ==========================================
' BASE64 ENCODING/DECODING
' ==========================================

Private Function Base64Encode(text As String) As String
    On Error GoTo ErrHandler
    
    ' Use MSXML2.DOMDocument for Base64 encoding
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = StringToByteArray(text)
    
    Base64Encode = xmlNode.text
    Exit Function
    
ErrHandler:
    ' Fallback: simple character substitution
    Base64Encode = SimpleBase64Encode(text)
End Function

Private Function Base64Decode(encodedText As String) As String
    On Error GoTo ErrHandler
    
    ' Use MSXML2.DOMDocument for Base64 decoding
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.text = encodedText
    
    Base64Decode = ByteArrayToString(xmlNode.nodeTypedValue)
    Exit Function
    
ErrHandler:
    ' Fallback: simple character substitution
    Base64Decode = SimpleBase64Decode(encodedText)
End Function

' ==========================================
' HELPER FUNCTIONS
' ==========================================

Private Function StringToByteArray(text As String) As Byte()
    Dim bytes() As Byte
    ReDim bytes(0 To Len(text) - 1)
    
    Dim i As Long
    For i = 1 To Len(text)
        bytes(i - 1) = Asc(mid$(text, i, 1))
    Next i
    
    StringToByteArray = bytes
End Function

Private Function ByteArrayToString(bytes() As Byte) As String
    Dim result As String
    Dim i As Long
    
    result = ""
    For i = LBound(bytes) To UBound(bytes)
        result = result & Chr$(bytes(i))
    Next i
    
    ByteArrayToString = result
End Function

' Simple fallback Base64 encode (not RFC compliant but works for simple cases)
Private Function SimpleBase64Encode(text As String) As String
    Const BASE64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    
    Dim result As String
    Dim i As Long
    Dim val1 As Long, val2 As Long, val3 As Long
    
    result = ""
    
    For i = 1 To Len(text) Step 3
        val1 = Asc(mid$(text, i, 1))
        
        If i + 1 <= Len(text) Then
            val2 = Asc(mid$(text, i + 1, 1))
        Else
            val2 = 0
        End If
        
        If i + 2 <= Len(text) Then
            val3 = Asc(mid$(text, i + 2, 1))
        Else
            val3 = 0
        End If
        
        result = result & mid$(BASE64_CHARS, (val1 \ 4) + 1, 1)
        result = result & mid$(BASE64_CHARS, ((val1 And 3) * 16 + val2 \ 16) + 1, 1)
        
        If i + 1 <= Len(text) Then
            result = result & mid$(BASE64_CHARS, ((val2 And 15) * 4 + val3 \ 64) + 1, 1)
        Else
            result = result & "="
        End If
        
        If i + 2 <= Len(text) Then
            result = result & mid$(BASE64_CHARS, (val3 And 63) + 1, 1)
        Else
            result = result & "="
        End If
    Next i
    
    SimpleBase64Encode = result
End Function

' Simple fallback Base64 decode
Private Function SimpleBase64Decode(encodedText As String) As String
    ' Simplified decode - just return the encoded text for now
    ' In production, implement proper Base64 decoding
    SimpleBase64Decode = encodedText
End Function

' ==========================================
' VALIDATION
' ==========================================

Public Function IsEncrypted(text As String) As Boolean
    ' Check if text appears to be Base64 encoded
    On Error Resume Next
    IsEncrypted = (Len(text) > 0 And (Len(text) Mod 4 = 0 Or InStr(text, "=") > 0))
End Function
