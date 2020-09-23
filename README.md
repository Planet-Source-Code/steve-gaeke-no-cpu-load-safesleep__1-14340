<div align="center">

## No CPU Load \- SafeSleep


</div>

### Description

The *BEST* way to safely Sleep/Pause without taxing the CPU.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Gaeke](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-gaeke.md)
**Level**          |Intermediate
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-gaeke-no-cpu-load-safesleep__1-14340/archive/master.zip)





### Source Code

<b>
The two functions below both sleep for the specified number of seconds. However, the popular code seen in the BusySleep procedure actually causes the CPU load to stay near 100% until complete. The SafeSleep routine pauses without taxing the CPU and stays at nearly 0% CPU.
Both functions take a single value so you can sleep for fractions of a second.</b>
<br><br>
<tt>
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long<br><br>
Public Sub SafeSleep(ByVal inWaitSeconds As Single)<br>
<nbsp><nbsp> Const WAIT_OBJECT_0 As Long = 0<br>
<nbsp><nbsp> Const WAIT_TIMEOUT As Long = &H102<br><br>
<nbsp><nbsp> Dim lastTick As Single<br>
<nbsp><nbsp> Dim timeout As Long<br>
<nbsp><nbsp> timeout = inWaitSeconds * 1000<br>
<nbsp><nbsp> lastTick = Timer<br><br>
<nbsp><nbsp> Do<br>
<nbsp><nbsp><nbsp><nbsp> Select Case MsgWaitForMultipleObjects(0, 0, False, timeout, 255)<br>
<nbsp><nbsp><nbsp><nbsp> Case WAIT_OBJECT_0<br>
<nbsp><nbsp><nbsp><nbsp>  DoEvents<br>
<nbsp><nbsp><nbsp><nbsp>  timeout = ((inWaitSeconds) - (Timer - lastTick)) * 1000<br>
<nbsp><nbsp><nbsp><nbsp>  If timeout < 0 Then timeout = 0<br><br>
<nbsp><nbsp><nbsp><nbsp> Case Else<br>
<nbsp><nbsp><nbsp><nbsp>  Exit Do<br><br>
<nbsp><nbsp><nbsp><nbsp> End Select<br><br>
<nbsp><nbsp> Loop While True<br><br>
End Sub<br><br>
Public Sub BusySleep(ByVal inWaitSeconds As Single)<br>
<nbsp><nbsp> Dim lastTick As Single<br><br>
<nbsp><nbsp> lastTick = Timer<br><br>
<nbsp><nbsp> Do<br>
<nbsp><nbsp><nbsp><nbsp> DoEvents<br><br>
<nbsp><nbsp> Loop While (Timer - lastTick) < inWaitSeconds<br><br>
End Sub<br>
</tt>

