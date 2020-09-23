<div align="center">

## Validate Date


</div>

### Description

Checks and makes sure that the user enters in a valid date. Make a text box on a form and call it txtDate
 
### More Info
 
I made this in a few minutes, after looking for a simple code so that I can make sure that if the user enters in bogus or an incorect form it will either clear it and throw a message or correct it and allow it to pass.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Powers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-powers.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-powers-validate-date__1-25180/archive/master.zip)





### Source Code

```
sub validDate()
 If txtDate.Text <> "" Then
  If IsDate(txtDate) = False Then
    MsgBox "You have entered an Invalid Date in Last Contacted please use MM/DD/YYYY", , "Invalid Entry"
    txtDate.text= ""
  Else
    txtDate.Text = Format(txtDate.Text, "General Date")
  End If
Else
  txtDate.Text = Date
End If
end sub
```

