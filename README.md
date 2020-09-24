<div align="center">

## Using sytem wide hooks ,spyworks\(cross task subclassing \) clone


</div>

### Description

Using system wide hooks in VB. Showing you how you can add menu's to other applications and react to their click button.With system wide hooks the limitations of modifying only those applications run in the same process ceases to exist. Do vote for me if u like the code
 
### More Info
 
Here is the source code of the dll made in c++

#include <windows.h>

int prevmnu;

int prevmsg;

int currmnu;

int currmsg;

HWND hndl;

LRESULT CALLBACK CallWndProc(

int nCode,   // hook code

WPARAM wParam, // type of hook about to be called

LPARAM lParam  // address of structure with debugging information

)

{

if (nCode>=0)

{

HWND trt=FindWindow(NULL,"Untitled - NotePad");

CWPSTRUCT *wt=(CWPSTRUCT*)lParam;

if (trt==wt->hwnd)

{

if(((wt->message) < 289 ) ||((wt->message) > 289 ) )

{

prevmsg=currmsg;

prevmnu=currmnu;

currmsg=wt->message ;

currmnu=LOWORD(wt->wParam);

}

if(currmsg==293)

if(prevmsg==WM_MENUSELECT)

if (prevmnu==35){

hndl=GetWindow(trt,GW_CHILD);

MessageBox(0,"You clicked a menu made by the dude","The dude rocks!!!",0);

if(hndl>0)

SendMessage(hndl,WM_SETTEXT,NULL,(LPARAM)"This was printed by the dude's program!!!!");

hndl=0;

}	}	}

return CallNextHookEx(NULL, nCode, wParam, lParam);

}


<span>             |<span>
---                |---
**Submitted On**   |2001-06-20 17:33:34
**By**             |[Venkat Mani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/venkat-mani.md)
**Level**          |Intermediate
**User Rating**    |4.7 (42 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Using syte213486192001\.zip](https://github.com/Planet-Source-Code/venkat-mani-using-sytem-wide-hooks-spyworks-cross-task-subclassing-clone__1-24209/archive/master.zip)








