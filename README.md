<div align="center">

## LZSS Compress/Decompress


</div>

### Description

This is a standard LZSS compression/decompression engine. It is written in VB for learning purposes, and should be converted to C/C++ if it is to be used with large amounts of data. It uses a dictionary compression algorithm (like ZIP,ARJ and others) and works the best on data with a lot of repetitions.
 
### More Info
 
sCompData - the string to be compressed, sDecompData - the string to be decompressed

Should be obvious


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jesper Soderberg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jesper-soderberg.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jesper-soderberg-lzss-compress-decompress__1-1821/archive/master.zip)





### Source Code

```
Option Explicit
Public Function sCompress(sCompData As String) As String
 Dim lDataCount As Long
 Dim lBufferStart As Long
 Dim lMaxBufferSize As Long
 Dim sBuffer As String
 Dim lBufferOffset As Long
 Dim lBufferSize As Long
 Dim sDataControl As String
 Dim bDataControlChar As Byte
 Dim lControlCount As Long
 Dim bControlPos As Byte
 Dim bCompLen As Long
 Dim lCompPos As Long
 Dim bMaxCompLen As Long
 lMaxBufferSize = 65535
 bMaxCompLen = 255
 lBufferStart = 0
 sDataControl = ""
 bDataControlChar = 0
 bControlPos = 0
 lControlCount = 0
 If Len(sCompData) > 4 Then
 sCompress = Left(sCompData, 4)
 For lDataCount = 5 To Len(sCompData)
  If lDataCount > lMaxBufferSize Then
  lBufferSize = lMaxBufferSize
  lBufferStart = lDataCount - lMaxBufferSize
  Else
  lBufferSize = lDataCount - 1
  lBufferStart = 1
  End If
  sBuffer = Mid(sCompData, lBufferStart, lBufferSize)
  If Len(sCompData) - lDataCount < bMaxCompLen Then bMaxCompLen = Len(sCompData) - lDataCount
  lCompPos = 0
  For bCompLen = 3 To bMaxCompLen Step 3
  If bCompLen > bMaxCompLen Then
   bCompLen = bMaxCompLen
  End If
  lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen), 0)
  If lCompPos = 0 Then
   If bCompLen > 3 Then
   While lCompPos = 0
    lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen - 1), 0)
    If lCompPos = 0 Then bCompLen = bCompLen - 1
   Wend
   End If
   bCompLen = bCompLen - 1
   Exit For
  End If
  Next
  If bCompLen > bMaxCompLen And lCompPos > 0 Then
  bCompLen = bMaxCompLen
  lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen), 0)
  End If
  If lCompPos > 0 Then
  lBufferOffset = lBufferSize - lCompPos + 1
  sCompress = sCompress & Chr((lBufferOffset And &HFF00) / &H100) & Chr(lBufferOffset And &HFF) & Chr(bCompLen)
  lDataCount = lDataCount + bCompLen - 1
  bDataControlChar = bDataControlChar + 2 ^ bControlPos
  Else
  sCompress = sCompress & Mid(sCompData, lDataCount, 1)
  End If
  bControlPos = bControlPos + 1
  If bControlPos = 8 Then
  sDataControl = sDataControl & Chr(bDataControlChar)
  bDataControlChar = 0
  bControlPos = 0
  End If
  lControlCount = lControlCount + 1
 Next
 If bControlPos <> 0 Then sDataControl = sDataControl & Chr(bDataControlChar)
 sCompress = Chr((lControlCount And &H8F000000) / &H1000000) & Chr((lControlCount And &HFF0000) / &H10000) & Chr((lControlCount And &HFF00) / &H100) & Chr(lControlCount And &HFF) & Chr((Len(sDataControl) And &H8F000000) / &H1000000) & Chr((Len(sDataControl) And &HFF0000) / &H10000) & Chr((Len(sDataControl) And &HFF00) / &H100) & Chr(Len(sDataControl) And &HFF) & sDataControl & sCompress
 Else
 sCompress = sCompData
 End If
End Function
Public Function sDecompress(sDecompData As String) As String
 Dim lControlCount As Long
 Dim lControlPos As Long
 Dim bControlBitPos As Byte
 Dim lDataCount As Long
 Dim lDataPos As Long
 Dim lDecompStart As Long
 Dim lDecompLen As Long
 If Len(sDecompData) > 4 Then
 lControlCount = Asc(Left(sDecompData, 1)) * &H1000000 + Asc(Mid(sDecompData, 2, 1)) * &H10000 + Asc(Mid(sDecompData, 3, 1)) * &H100 + Asc(Mid(sDecompData, 4, 1))
 lDataCount = Asc(Mid(sDecompData, 5, 1)) * &H1000000 + Asc(Mid(sDecompData, 6, 1)) * &H10000 + Asc(Mid(sDecompData, 7, 1)) * &H100 + Asc(Mid(sDecompData, 8, 1)) + 9
 sDecompress = Mid(sDecompData, lDataCount, 4)
 lDataCount = lDataCount + 4
 bControlBitPos = 0
 lControlPos = 9
 For lDataPos = 1 To lControlCount
  If 2 ^ bControlBitPos = (Asc(Mid(sDecompData, lControlPos, 1)) And 2 ^ bControlBitPos) Then
  lDecompStart = Len(sDecompress) - (CLng(Asc(Mid(sDecompData, lDataCount, 1))) * &H100 + CLng(Asc(Mid(sDecompData, lDataCount + 1, 1)))) + 1
  lDecompLen = Asc(Mid(sDecompData, lDataCount + 2, 1))
  sDecompress = sDecompress & Mid(sDecompress, lDecompStart, lDecompLen)
  lDataCount = lDataCount + 3
  Else
  sDecompress = sDecompress & Mid(sDecompData, lDataCount, 1)
  lDataCount = lDataCount + 1
  End If
  bControlBitPos = bControlBitPos + 1
  If bControlBitPos = 8 Then
  bControlBitPos = 0
  lControlPos = lControlPos + 1
  End If
 Next
 Else
 sDecompress = sDecompData
 End If
End Function
'Put a two command buttons (Command1 and Command2) on to a form and paste the following on to it as well:
Option Explicit
Private Const sFileName = "c:\compressthis.exe" ' the file to be compressed
Private Sub Command1_Click() 'Compress the file
 Dim sReturn As String
 Dim sFileData As String
 Open sFileName For Binary As #1
  sFileData = Input(LOF(1), #1)
 Close #1
 sReturn = sCompress(sFileData)
 Debug.Print Len(sReturn), Len(sFileData)
 Open Left(sFileName, Len(sFileName) - 3) & "wnc" For Output As #1
  Print #1, sReturn;
 Close #1
End Sub
Private Sub Command2_Click() 'Decompress the file
 Dim sReturn As String
 Dim sFileData As String
 Open Left(sFileName, Len(sFileName) - 4) & ".wnc" For Binary As #1
  sFileData = Input(LOF(1), #1)
  sReturn = sDecompress(sFileData)
 Close #1
 Debug.Print Len(sReturn), Len(sFileData)
 Open Left(sFileName, Len(sFileName) - 4) & "2" & Right(sFileName, 4) For Output As #1
  Print #1, sReturn;
 Close #1
End Sub
```

