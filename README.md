<div align="center">

## FAST Tail Function for VB


</div>

### Description

After hunting for this functionality since I started using VB, I've finally figured out how to *quickly* read the last few lines in a file. This function fills a dynamic array with the last X number of lines in a file you specify. I use it to monitor apachewin32 logs on another server. This is much faster than using 'line input'. I know there may be more optimized ways of doing this, I just thought I should get it out there.
 
### More Info
 
Assumes a publicly declared dynamic array.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[pt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pt.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pt-fast-tail-function-for-vb__1-29692/archive/master.zip)





### Source Code

```
Public Function Tail(fName As String, NumOfLines As Integer, ArrayName() As String)
Dim ff As Integer
Dim raw As String
Dim lines() As String 'used to hold the lines of the text file
Dim lStart As Integer 'switch to LONG if you have over 65k lines.
If Not fileExist(fName) Then
MsgBox "File Not Found - " & vbCrLf & fName, vbCritical, "Error"
Exit Function
End If
ff = FreeFile
Open fName For Binary As #ff
raw = String$(LOF(ff), 32)
Get #ff, 1, raw
Close #ff
lines() = Split(raw, vbNewLine) 'this assumes that the data is stored in individual lines.
ReDim ArrayName(NumOfLines)
If NumOfLines > UBound(lines) Then NumOfLines = UBound(lines)
lStart = UBound(lines) - NumOfLines
For i = 1 To NumOfLines
ArrayName(i) = lines(lStart + i)
Next i
End Function
'and the bonus 'FILEEXIST' function:
Public Function fileExist(filename As String) As Boolean
 Dim l As Long
 On Error Resume Next
 l = FileLen(filename)
 fileExist = Not (Err.Number > 0)
 On Error GoTo 0
End Function
```

