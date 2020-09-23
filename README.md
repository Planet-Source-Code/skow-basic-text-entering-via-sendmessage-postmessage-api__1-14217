<div align="center">

## Basic Text Entering via SendMessage / Postmessage API


</div>

### Description

This is a *baisc* sample of how to use the PostMessage API calls (same as SendMessage) to send text to a Textbox (or any hWnd actually).

Note: I tracked down the 3 WM_ messages through SPY++ and Notepad. Differnt applications may require more/less WM_ messages. A VB Textbox only needs WM_KEYDOWN but I added the other code incase. =].. Sorry about the cheap Variable names, this was programmed at 5am :)
 
### More Info
 
To make life easy, Assume you have a Form with a Text box on it (Text2) which will have the text to SEND to another text box (which we will put on the form so the hWnd is easy to get: Text1)

May have some issues if you put in random hWnds.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SKoW](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/skow.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/skow-basic-text-entering-via-sendmessage-postmessage-api__1-14217/archive/master.zip)

### API Declarations

```
' I like Postmessage over Sendmessage
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
```


### Source Code

```

Sub SendText(hWnd As Long, Text As String)
 If hWnd = 0 Then MsgBox "no hWnd supplied": Exit Sub
' got hWnd, start sending messages
Dim zwParam As Long ' so no dupe deffs use z infrount =]
Dim zlParam As Long
Dim xwParam As Long ' used for WM_CHAR
For I = 1 To Len(Text)
  ' First, get the lParam for WM_KEYDOWN
  zwParam = GetVKCode(Mid$(Text, I, 1))
  xwParam = zwParam And &H20 ' wants Hex20 added to it so A7 goes to C7 and 15 -> 35 (hex values)
  zlParam = GetScanCode(Mid$(Text, I, 1))
  PostMessage hWnd, WM_KEYDOWN, zwParam, zlParam
  ' Used in notepad, doesn't seem to be used in this example
  'PostMessage hWnd, WM_CHAR, xwParam, zlParam
  ' Used in notepad, but doubles the chars in this example..
  'zlParam = zlParam And &HC0000000 ' wants hex-C (7x0's) added.
  'PostMessage hWnd, WM_KEYUP, zwParam, zlParam
  DoEvents
Next
End Sub
Function GetVKCode(ByVal Char As String) As Long
 On Error Resume Next
 Char = UCase(Left$(Char, 1))
 GetVKCode = Asc(Char)
End Function
Function GetScanCode(bChar As String) As Long
' To get scancodes:
' Start SPY++ on Notepad
'Type in all chars and then stop SPY++ logging. It will tell you all scancodes
' recorded during the logging.. long but ah well..
' Note: Scancode 1E = &H1E0001,  30 = &H300001
'
 Select Case LCase$(Left$(bChar, 1))
  Case "a"
    GetScanCode = &H1E0001
  Case "b"
    GetScanCode = &H300001
  Case "c"
    GetScanCode = &H2E0001
  Case "d"
    GetScanCode = &H200001
  Case "e"
    GetScanCode = &H120001
  Case "f"
    GetScanCode = &H210001
  Case "g"
    GetScanCode = &H220001
  Case "h"
    GetScanCode = &H230001
  Case "i"
    GetScanCode = &H170001
  Case "j"
    GetScanCode = &H240001
  Case "k"
    GetScanCode = &H250001
  Case "l"
    GetScanCode = &H260001
  Case "m"
    GetScanCode = &H320001
  Case "n"
    GetScanCode = &H310001
  Case "o"
    GetScanCode = &H180001
  Case "p"
    GetScanCode = &H190001
  Case "q"
    GetScanCode = &H100001
  Case "r"
    GetScanCode = &H130001
  Case "s"
    GetScanCode = &H1F0001
  Case "t"
    GetScanCode = &H140001
  Case "u"
    GetScanCode = &H160001
  Case "v"
    GetScanCode = &H2F0001
  Case "w"
    GetScanCode = &H110001
  Case "x"
    GetScanCode = &H2D0001
  Case "y"
    GetScanCode = &H150001
  Case "z"
    GetScanCode = &H2C0001
  Case Else
    GetScanCode = 0 ' no scode at the mo =(
  End Select
End Function
```

