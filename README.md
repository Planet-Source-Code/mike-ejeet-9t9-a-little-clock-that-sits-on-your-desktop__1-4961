<div align="center">

## A little clock that sits on your desktop


</div>

### Description

This gives the code for a little arlarm clock that sits on your desktop. The soul purpose of this application is to give the current time. Very Simple.
 
### More Info
 
1 timer- Interval 500 and name timer 1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike\-Ejeet 9t9](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-ejeet-9t9.md)
**Level**          |Beginner
**User Rating**    |2.3 (9 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\)
**Category**       |[Jokes/ Humor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/jokes-humor__1-40.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-ejeet-9t9-a-little-clock-that-sits-on-your-desktop__1-4961/archive/master.zip)





### Source Code

```
I(n the form itself enter this code:
Private Sub Form_Click()
  AlarmTime = InputBox("Enter alarm time", "VB Alarm", AlarmTime)
  If AlarmTime = "" Then Exit Sub
  If Not IsDate(AlarmTime) Then
    MsgBox "The time you entered was not valid."
  Else                  ' String returned from InputBox is a valid time,
    AlarmTime = CDate(AlarmTime)    ' so store it as a date/time value in AlarmTime.
  End If
End Sub
**********In the timer enter this code:*****************
Private Sub Timer1_Timer()
Static AlarmSounded As Integer
  If lblTime.Caption <> CStr(Time) Then
    ' It's now a different second than the one displayed.
    If Time >= AlarmTime And Not AlarmSounded Then
      Beep
      MsgBox "Alarm at " & Time
      AlarmSounded = True
    ElseIf Time < AlarmTime Then
      AlarmSounded = False
    End If
    If WindowState = conMinimized Then
      ' If minimized, then update the form's Caption every minute.
      If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
    Else
      ' Otherwise, update the label Caption in the form every second.
      lblTime.Caption = Time
    End If
  End If
End Sub
```

