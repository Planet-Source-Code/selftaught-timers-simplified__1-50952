<div align="center">

## Timers, simplified


</div>

### Description

Now updated with some comments!

This is a simple timer module with a standard implementation class. The module uses only a single timer resource, but it can fire events to a nearly unlimited number of subscribers. The worst precision that I have gotten is just under 2 hundredths of a second. The code uses the system time to determine when to fire, so there is no uptime limitation as with GetTickCount.
 
### More Info
 
Implement the iTimer interface on any of your classes, and then call:

SetATimer Me, (Interval), (Tag)

where interval is in milliseconds and Tag is an identifier for the timer to be returned with each fire. You can set multiple timers for one object without creating any new objects.

Uses an illegal object reference to allow the object to terminate without forcing you to have a dispose-type method. This works just fine, as long as you call KillTimers Me in the terminate event of your class.

This code kills the timer after each entry to TimerProc. It goes through all of the subscribers and decides when the next one should be, and sets a timer for the appropriate duration. This allows many timers with different intervals to be fired, never using more than a single timer resource.


<span>             |<span>
---                |---
**Submitted On**   |2004-01-13 16:34:08
**By**             |[selftaught](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/selftaught.md)
**Level**          |Beginner
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Timers\_\_si1694001132004\.zip](https://github.com/Planet-Source-Code/selftaught-timers-simplified__1-50952/archive/master.zip)








