<div align="center">

## Decimal To Roman Converter


</div>

### Description

This code takes an input decimal number and converts it to roman numerals. Will work for any number inputted until the output string becomes too big (and testing seems to show this doesnt actuually happen)! This is the shortest piece of code to o this (to my knowledge), as all the others I've seen are useless....
 
### More Info
 
sDecNum - the decimal number you wish to convert.

ToRoman - the string returned from the function of the roman numeral.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[C M Buckley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/c-m-buckley.md)
**Level**          |Beginner
**User Rating**    |2.8 (14 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/c-m-buckley-decimal-to-roman-converter__1-39833/archive/master.zip)





### Source Code

```
Const sMatrix As String = "I~V~X~L~C~D~M"
Private Function ToRoman(ByVal sDecNum As String) As String
  If sDecNum <> "0" And sDecNum <> vbNullString Then
    Dim sNumArray() As String
    If Len(sDecNum) > 3 Then ToRoman = String(Mid(sDecNum, 1, Len(sDecNum) - 3), "M")
    If Len(sDecNum) > 2 Then ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum) - 2, 1), 4)
    If Len(sDecNum) > 1 Then ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum) - 1, 1), 2)
    ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum), 1), 0)
  Else: ToRoman = "No Roman value for 0"
  End If
End Function
Private Function GiveLetters(ByVal sInput As String, ByVal iArrStart As Integer) As String
  Dim sLetterArray() As String
  sLetterArray() = Split(sMatrix, "~")
  Select Case sInput
    Case 4: GiveLetters = sLetterArray(iArrStart) & sLetterArray(iArrStart + 1)
    Case 5: GiveLetters = sLetterArray(iArrStart + 1)
    Case 9: GiveLetters = sLetterArray(iArrStart) & sLetterArray(iArrStart + 2)
    Case 6 To 8: GiveLetters = sLetterArray(iArrStart + 1) & String(sInput - 5, sLetterArray(iArrStart))
    Case Else: GiveLetters = GiveLetters + String(sInput, sLetterArray(iArrStart))
  End Select
End Function
Private Sub Command1_Click()
  Dim sRoman As String
  sRoman = ToRoman(2002)
End Sub
```

