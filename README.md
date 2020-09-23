<div align="center">

## Compression, uncompression using RLE\-algorithm


</div>

### Description

Compresses strings, most effective on bitmap files
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[jouni\.vuorio@vtoy\.fi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jouni-vuorio-vtoy-fi.md)
**Level**          |Unknown
**User Rating**    |4.1 (161 globes from 39 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jouni-vuorio-vtoy-fi-compression-uncompression-using-rle-algorithm__1-745/archive/master.zip)





### Source Code

```
'Copyright 1997 Jouni vuorio
public function compress()
On Error Resume Next
For TT = 1 To Len(Text1)
sana1 = Mid(Text1, TT, 1)
sana2 = Mid(Text1, TT + 1, 1)
sana3 = Mid(Text1, TT + 2, 1)
X = 1
If Not sana1 = sana2 Then löyty = 2
If sana1 = sana2 Then
If sana1 = sana3 Then
löyty = 1
End If
End If
If löyty = 1 Then
alku:
X = X + 1
merkki = Mid(Text1, TT + X + 1, 1)
If merkki = sana1 Then GoTo alku
sana = Chr(255) & Chr(X - 1) & sana1
TT = TT + X
End If
If löyty = 2 Then sana = sana1
Text = Text & sana
Next
Text1 = Text
end function
public function uncompress()
On Error Resume Next
For TT = 1 To Len(Text1)
sana1 = Asc(Mid(Text1, TT, 1))
sana2 = Asc(Mid(Text1, TT + 1, 1))
sana3 = Asc(Mid(Text1, TT + 2, 1))
sana4 = Asc(Mid(Text1, TT - 1, 1))
If sana1 = 255 Then
For TT6 = 1 To sana2
sana = sana & Chr(sana3)
Next
sana1 = ""
sana2 = ""
End If
If sana = "" Then
If Not sana4 = 255 Then
sana = Chr(sana1)
End If
End If
Text = Text & sana
sana = ""
Next
Text1 = Text
end function
'comments to jouni.vuorio@vtoy.fi
```

