<div align="center">

## Forget NetZero, get Net Hero \(Updated\)


</div>

### Description

### Functions updated. Now they're way faster and smaller. Took out Select Case and added InStr. ##

Ever wanted to use NetZero? Without the ZeroPort software?

Well now it's possible. NetHero will convert any existing NetZero username and password to a username and password that will work with Dial-Up Networking.
 
### More Info
 
Username, Password

Function to use:

zEncryptUsername

zEncryptPassword

Modified Username, Encrypted Password


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Arachnid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arachnid.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/arachnid-forget-netzero-get-net-hero-updated__1-9641/archive/master.zip)





### Source Code

```
Const Strin1 = "`-=~!@#$%^&*()_+[]\{}|;':" & """" & _
  ",./<>?abcdefghijklmnopqrstuvwxyzABCDEFG" & _
  "HIJKLMNOPQRSTUVWXYZ0123456789"
Const Strin2 = "GFEDCBAzyxwvutsrqponmlkjihgfed" & _
  "cba?></.," & """" & ":';|}{\][+_)(*&^%$#@" & _
  "!~=-`9876543210ZYXWVUTSRQPONMLKJIH"
Function Convert(Character)
  Qt = Chr(34)
  Chr1 = InStr(1, Strin1, Character)
  Chr2 = Mid(Strin2, Chr1, 1)
  Convert = Chr2
End Function
Function CharacterToNum(Character)
  CharacterToNum = InStr(1, Strin1, Character)
End Function
Function NumToCharacter(TheNumber)
  NumToCharacter = Mid(Strin1, TheNumber, 1)
End Function
Function zEncryptPassword(Password)
  For i = 0 To Len(Password) - 1
    TheCur = Mid(Password, i + 1, 1)
    Asdf = CharacterToNum(TheCur)
    Asdf = Asdf - i
    Asdf2 = NumToCharacter(Asdf)
    Asdf3 = Convert(Asdf2)
    SomeString = SomeString + Asdf3
  Next i
  zEncryptPassword = "0" & SomeString & "1"
End Function
Function zEncryptUsername(Username)
  zEncryptUsername = "2.2.2:" & Username & "@netzero.net"
End Function
```

