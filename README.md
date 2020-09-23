<div align="center">

## URLDecode Function


</div>

### Description

Decodes a URLEncoded string
 
### More Info
 
sEncodedURL - Encoded String to Decode

Decoded String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Markus Diersbock](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/markus-diersbock.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/markus-diersbock-urldecode-function__1-44365/archive/master.zip)





### Source Code

```
Public Function URLDecode(sEncodedURL As String) As String
 On Error GoTo Catch
 Dim iLoop As Integer
 Dim sRtn As String
 Dim sTmp As String
 If Len(sEncodedURL) > 0 Then
 ' Loop through each char
 For iLoop = 1 To Len(sEncodedURL)
 sTmp = Mid(sEncodedURL, iLoop, 1)
 sTmp = Replace(sTmp, "+", " ")
 ' If char is % then get next two chars
 ' and convert from HEX to decimal
 If sTmp = "%" and LEN(sEncodedURL) + 1 > iLoop + 2 Then
 sTmp = Mid(sEncodedURL, iLoop + 1, 2)
 sTmp = Chr(CDec("&H" & sTmp))
 ' Increment loop by 2
 iLoop = iLoop + 2
 End If
 sRtn = sRtn & sTmp
 Next iLoop
 URLDecode = sRtn
 End If
Finally:
 Exit Function
Catch:
 URLDecode = ""
 Resume Finally
End Function
```

