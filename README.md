# PRO-BOT

Extends the capabilities of your PC to automatically control your RCX (or vice-versa) beyond imagination! Compatible with all versions of the RCX, including v1.0, v1.5, and now v2.0.

### What is PRO-BOT?
PRO-BOT is a program editor for the Lego © RCX programmable brick. Using PRO-BOT, you can build programs, and download them to the RCX. You can also send immediate commands (ex. PlaySystemSound, or PBTurnOff). You can also retrieve information from the RCX using commands such as Poll, or MemMap. PRO-BOT is compatible with the Lego firmware firmw0309.lgo and firmw0328.log. Complete documentation is available on the Lego Mindstorms web site.

### Do you need PRO-BOT?
It all depends on what are your objectives: If you want to build autonomous robots, robots which can run totally independently of the PC computer, robots whose program is fully located in the RCX brick, then you don't need PRO-BOT. Any other environment (such as RCXCC, BrickCommand, Gordon's Brick, NQC, etc.) do have all the same capabilities as PRO-BOT. However, PRO-BOT is directly aimed at keeping a constant link between the PC and the brick. If you wish to constantly exchange information or programs between the two computers, then you absolutely need PRO-BOT.

### What’s more in PRO-BOT?
PRO-BOT has unique features: it can interact with programs located on your computer. Now, with PRO-BOT, your computer can shut down the RCX, or the RCX can shut down your computer! (even without touching it, of course) PRO-BOT implements instructions that are intended for the computer only. Such commands include sending information to files (such as PollTo) and taking decisions based on the content of a file.

### Examples of programs using PRO-BOT.
With PRO-BOT, you can control your robot from your computer (see keyboard.rcp program). The RCX has a complete program to face most of the situation; on the computer, based on the keyboard, you have a program that change the variable 1 of the RCX. RCX uses this variable to decide its action.

As another example, suppose your robot is a small vacuum programmed to wandered by night. Your vacuum must withdraw when light is coming. If it gets stuck, it request (through your computer) to send a email. Next thing you know is that you receive a e-mail from your vacuum!

Finally, genetic algorithm is now a reality. Your computer can generate new programs for your RCX based on cross-over principles between efficient programs, and send them to the RCX for a test. After a trial period, the computer evaluates the how good the RCX program is, and breed together the best fitted.


### Philosophy of PRO-BOT.
PRO-BOT is intended to be the simplest possible. There is no test panel, there is no button to download the firmware. Everything has to be done through programs, and there is a program for everything. This makes easier to send listings to other users. Further, you can user other commands to download firmware (such as firmdl.c and save ½ Kbytes) without problem.

### Download and installation of PRO-BOT
* You do not need SPIRIT.ocx anymore on your system. A replacement, phantom.dll (by Fenestra inc., freeware) is automatically installed by PRO-BOT.
* If you don't have the previous version of PRO-BOT, you may need the Visual Basic Run-Time library: MSVBVM50.exe (included in the GitHub release). Download it to your desktop, and run it.
* To install PRO-BOT, simply download, unzip, and run install.exe.
* Sample files. Download it to your desktop and unzip them under the Program Files\PRO-BOT 2000 directory.
* The source code of PRO-BOT for Visual Basic 6 is available.

### Resources
* [Examples](http://www.mapageweb.umontreal.ca/cousined/lego/4-RCX/examples/index.html):  Follow this link for examples of some of the more elaborate projects completed with PRO-BOT. All these examples are included with the sample files that you can download above.
* [History](http://www.mapageweb.umontreal.ca/cousined/lego/4-RCX/history/history.html):  Follow this link for a history of the versions of PRO-BOT.

* * *

Originally available at http://www.mapageweb.umontreal.ca/cousined/lego/4-RCX/PRO-BOT/

Special thanks from author Denis Cousineau to Brandon Yates for many suggestions and reporting some bugs.
