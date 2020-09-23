<div align="center">

## Stealing another window


</div>

### Description

This code will allow you to take a program running and put the program in your form.
 
### More Info
 
You may not be able to end it in the program so you may have to do a ctrl+alt+del and end it manualy..havnt figured it out yet...

Also may not be reposistionable in the form...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Khris](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/khris.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/khris-stealing-another-window__1-64956/archive/master.zip)





### Source Code

```
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_VISIBLE = &H10000000
Private Sub Form_Load()
Dim Handle As Long, Ret As Long
'look for the window handle
Handle = FindWindow(vbNullString, "EDHacks.com - Mozilla Firefox")'This is where you put the title of the program/window.
Ret = SetWindowLong(Handle, GWL_STYLE, WS_VISIBLE Or WS_CLIPSIBLINGS)
'This is where the program will be brought into the form.
SetParent Handle, Me.hwnd
End Sub
```

