<div align="center">

## ULTIMATE  CD Function


</div>

### Description

This Code opens and closes the cd-rom drive at a given interval. This can be used for security purposes and just for laughs. (though cruel laughs they would be ) This code is great for beginners. It shows how to use a timer in a very simple way.
 
### More Info
 
Needs two command buttons (optional but highly recommended) and a timer named Timer1.

Opens and closes cd door.

None unless you make it so that the timer will never be disabled.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[zOnE](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zone.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zone-ultimate-cd-function__1-6179/archive/master.zip)

### API Declarations

```
Declare Function mciSendString Lib "MMSystem" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal wReturnLength As Integer, ByVal hCallback As Integer) As Long
```


### Source Code

```
' This should work with ALL versions of VB, BUT
' It was only tested with VB4 (16-Bit). I will
' Be sure to test it on VB6. Just Follow the code
' Below
Make a timer and name it Timer1
 Set its Enabled property to False
Now set its Interval Property to the time you want the action to occur.
Make 2 command buttons.
Label 1 of them ON
 and
the other OFF.
'In Timer1 Place the following code
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
retvalue = mciSendString("set CDAudio door closed", returnstring, 127,0)
' In Command1 labeled On place the following code
Timer1 = Enabled
' In Command2 labeled OFF place the following code
Timer1 = Disbaled
```

