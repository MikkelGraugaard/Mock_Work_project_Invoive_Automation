```{r}
Option Explicit

Sub extract()

    'Setting ScreenUpdating to False in order to optimize the macro
    Application.ScreenUpdating = False
    
'Defining work sheet and longs
 Dim s1 As Worksheet
 Set s1 = Sheets("Ark1")
 Dim i As Long, lr As Long, lr2 As Long
 
 'Setting lr to loop through the C column as it is here the invoice text is stored
 lr = s1.Range("C" & Rows.Count).End(xlUp).Row
 For i = 1 To lr
 
 'lr2 is looking in cell C for the category specified in cell R1
 'If the C cell contains the category, it returns a number above 0 which activates the if statement
 'It then copies the related cost stored in cell R (one row above)
 'And pastes the cost to the next empty row under the specified category
 lr2 = s1.Range("R" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("R1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("R" & lr2 + 1)
 End If
 
 lr2 = s1.Range("S" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("S1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("S" & lr2 + 1)
 End If
 
 lr2 = s1.Range("T" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("T1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("T" & lr2 + 1)
 End If
 
 lr2 = s1.Range("U" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("U1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("U" & lr2 + 1)
 End If
 
 lr2 = s1.Range("V" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("V1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("V" & lr2 + 1)
 End If
 
 lr2 = s1.Range("W" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("W1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("W" & lr2 + 1)
 End If
 
 lr2 = s1.Range("X" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("X1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("X" & lr2 + 1)
 End If
 
 lr2 = s1.Range("Y" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("Y1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("Y" & lr2 + 1)
 End If
 
 lr2 = s1.Range("Z" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("Z1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("Z" & lr2 + 1)
 End If
 
 lr2 = s1.Range("AA" & Rows.Count).End(xlUp).Row
 If InStr(1, s1.Range("C" & i), s1.Range("AA1")) > 0 Then
 s1.Range("I" & i - 1).Copy s1.Range("AA" & lr2 + 1)
 End If
 
 
 Next i
 
    Application.ScreenUpdating = True

End Sub
```
