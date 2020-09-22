<div align="center">

## ckReplace \(now with invalid char stripping\)


</div>

### Description

For use with MS Access databases mostly. - this function allows you to with strip characters from a string, replace characters in a string with other characters or strip/replace all non-alpha characters (not printable) from the string.
 
### More Info
 
strIN is the string you wish to modify

'StripChar is the character you wish to remove/replace

'ReplaceChar is the character to use in "Stripchar"s place.

'Only strIN is required.

Returns the submitted string with the modifications made as a string:

'ckReplace("This is a test"," ","") returns "Thisisatest"

'ckReplace("This is a test","i","x") returns "Thxs xs a test"

'ckReplace("Sometext%MoreText") where the % represents some non printing character (like a line feed or someting - would return "SometextMoreText"

'ckReplace("Sometext%MoreText",""," ") where the % represents some non printing character (like a line feed or someting - would return "Sometext MoreText"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chad M\. Kovac](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chad-m-kovac.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VBA MS Access
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chad-m-kovac-ckreplace-now-with-invalid-char-stripping__1-33359/archive/master.zip)





### Source Code

```
Function ckReplace(StrIN As String, Optional StripChar As String = "", Optional ReplaceChar As String = "") As String
 Dim x As Integer
 x = 1
 If StripChar <> "" Then
  Do Until x <= 0 Or StripChar = ReplaceChar
   x = InStr(1, StrIN, StripChar)
   If x > 0 Then StrIN = left$(StrIN, x - 1) & ReplaceChar & Right$(StrIN, Len(StrIN) - (x - 1) - Len(StripChar))
  Loop
 Else
  For x = 1 To Len(StrIN)
   If x > Len(StrIN) Then Exit For
   If Asc(Mid$(StrIN, x, 1)) < 32 Or Asc(Mid$(StrIN, x, 1)) > 126 Then
    StrIN = left$(StrIN, x - 1) & ReplaceChar & Right$(StrIN, Len(StrIN) - (x - 1) - 1)
    If ReplaceChar = "" Then x = x - 1
   End If
  Next
 End If
 ckReplace = StrIN
End Function
```

