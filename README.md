<div align="center">

## Avoid multiple instances by using Mutexes \(much stronger than App\.PrevInstance\)


</div>

### Description

If you want your program to have only one instance, then there exists VB's in-built way to do it, the App.PrevInstance property. But it does not work if you copy your exe file elsewhere and then run (at least in my machine, Win98). So, here is another approach that guarantees only one instance of your app whether it is copied to different paths or renamed unless any serious error has occurred. You have to compile it to see the effect. And it will also work if the application has crashed for some reason. The code is fairly commented, hope it will help somebody. Please report bugs or any problem in this method of avoiding multiple instances. I will appreciate comments greatly.######Special thanks to LiTe for first pointing out that there is a problem in my previous method. However, it is solved now.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |1999-01-01 11:29:06
**By**             |[Isbat Sakib](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/isbat-sakib.md)
**Level**          |Intermediate
**User Rating**    |4.9 (49 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Avoid\_mult1794929172004\.zip](https://github.com/Planet-Source-Code/isbat-sakib-avoid-multiple-instances-by-using-mutexes-much-stronger-than-app-previnstance__1-56211/archive/master.zip)








