Attribute VB_Name = "loadString"
Option Explicit

Rem numbre of elements in the language
Public Const nbrtoken = 91
Public Const nbreparam = 22

Rem comment separators
Public Const CommentOpening = "{"
Public Const CommentClosing = "}"

Rem /////////////////////////
Rem memorizing all the symbols of the language, and help
Rem along with the number of parameters required.
Rem see loadstrings() for initializing these values
Rem /////////////////////////
Public paramlist(nbreparam) As String
Public paramhelp(nbreparam) As String

Public tokenlist(nbrtoken) As String
Public tokencolor(nbrtoken) As Single
Public tokenhelp(nbrtoken) As String
Public tokenparam(nbrtoken) As String

Rem tokencolor is also token category:
Rem 1:  comment
Rem 2:  preprocessor
Rem 3:  section
Rem 4:  Host based commands
Rem 4.1   sending information
Rem 4.2   retrieving information
Rem 4.3   host only
Rem 5:  RCX commands
Rem 5.1   immediate command
Rem 5.2   downloadable command
Rem 5.3   immediate/downloadable command

Sub LoadParam()
Rem f
  paramlist(1) = "f"
  paramhelp(1) = "'Filename'"
Rem q=w=f   : quoted file or name
  paramlist(2) = "q"
  paramhelp(2) = "'QuotedText'"
Rem w
  paramlist(3) = "w"
  paramhelp(3) = "ValueOrString"
Rem v       : variable nbre, (0..31)
  paramlist(4) = "v"
  paramhelp(4) = "VarNo"
Rem s       : source  (0..15)
  paramlist(5) = "s"
  paramhelp(5) = "SourceNo"
Rem P       : Program no (0..4)
  paramlist(6) = "P"
  paramhelp(6) = "ProgNo"
Rem n       : nbre fct of s
  paramlist(7) = "n"
  paramhelp(7) = "Number"
Rem m       : motor list/sensor list (0..2)
  paramlist(8) = "m"
  paramhelp(8) = "MotorList"
Rem t       : time (0..60)
  paramlist(9) = "t"
  paramhelp(9) = "Minutes"
Rem H       : Hours
  paramlist(10) = "H"
  paramhelp(10) = "Hour"
Rem l       : number (-32000..32000)
  paramlist(11) = "l"
  paramhelp(11) = "Integer"
Rem 7       : number (0..7)
  paramlist(12) = "7"
  paramhelp(12) = "Value"
Rem 4       : timer number (0..3)
  paramlist(13) = "4"
  paramhelp(13) = "TimerNo"
Rem 2       : short (0..255)
  paramlist(14) = "2"
  paramhelp(14) = "number"
Rem b       : binary value (0..1)
  paramlist(15) = "b"
  paramhelp(15) = "Boolean(0,1)"
Rem 6       : sensor type
  paramlist(16) = "6"
  paramhelp(16) = "Type"
Rem o       : sensor mode
  paramlist(17) = "o"
  paramhelp(17) = "Mode"
Rem y       : sensor slope
  paramlist(18) = "y"
  paramhelp(18) = "Slope"
Rem R       : relational operator (0..3)
  paramlist(19) = "R"
  paramhelp(19) = "RelOperator"
Rem i
  paramlist(20) = "i"
  paramhelp(20) = "Identifier"
Rem d
  paramlist(21) = "d"
  paramhelp(21) = "DelayInTensOfMs"
Rem Z
  paramlist(22) = "Z"
  paramhelp(22) = "Option"

End Sub

Sub LoadStrings()
  Rem load strings load all the elements of the
  Rem language in arrays; that way, using
  Rem inList(i) (see syntax) it is easy to
  Rem decide if a word is valid or not, also
  Rem make easier showing help.
  Rem /////////////////////////
  Dim i As Integer
  i = 0

  tokenlist(i) = ""
  tokenhelp(i) = ""
  tokenparam(i) = ""
  tokencolor(i) = 0
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- PREPROCESSOR -" + CommentClosing
  tokenhelp(i) = "Before syntax is analysed, Preprocessing commands are analyzed and substituted.  There is two preprocessing pass: Pass1 analyze all the Insert and Define commands.  Then, Pass2 analyzes all the Declare commands.  Due to this order, there can be no Declared symbols in Define commands. Check the option 'see compiling' to see the result of preprocessing."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "Insert"
  tokenhelp(i) = "insert the content of a RCP file."
  tokenparam(i) = "f"
  tokencolor(i) = 2
  i = i + 1
  
  tokenlist(i) = "Define"
  tokenhelp(i) = "Create symbolic names.  for example, you can give names to your filename with DEFINE file ('c:\results.txt').  This command is mostly useful if your definition span more than one elements, for example: DEFINE = (,2,) is very useful in the IF command."
  tokenparam(i) = "i(w)"
  tokencolor(i) = 2
  i = i + 1
  
  tokenlist(i) = "Declare"
  tokenhelp(i) = "Declare is similar to Define except that it is context-sensitive.  For example, you can relate the identifier FOO to variable 0.  After that, you don't need to specify the source anymore, it will be provided by the preprocessor when requires. Example is SetVar(foo, con, 1) and Poll(foo).  In the second case, the source type will be provided automatically. Declare are analyzed on Pass2."
  tokenparam(i) = "isn"
  tokencolor(i) = 2
  i = i + 1
  
  tokenlist(i) = "%Built-in strings%"
  tokenhelp(i) = "Built-in strings are predefined strings that can be used at run-time in your programs.  They are %ComPort% (Communication port connected to the IR), %WorkingPath% (Path of the current RCP program), %ProgName% (your rcp program name), and %AppPath% (path of PRO-RCX 2000)"
  tokenparam(i) = ""
  tokencolor(i) = 2
  i = i + 1
  
  tokenlist(i) = "?Suser-defined?"
  tokenhelp(i) = "User-defined strings receive a value at preprocessing. An InputBox with a question appear during preprocessing, and the answer is feed in the Declare. S (optional) means that a string, quoted value is returned."
  tokenparam(i) = ""
  tokencolor(i) = 2
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- SECTIONS -" + CommentClosing
  tokenhelp(i) = "Sections are used to separate commands activated on 'Execute' (F5) from commands activated by events such as Event or TimeOut.  Commands not in any section will be executed on all occasion."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "Main"
  tokenhelp(i) = "All commands preceeded by the Main() keyword are compiled when 'Execute' (F5) is used."
  tokenparam(i) = "()"
  tokencolor(i) = 3
  i = i + 1
  
  tokenlist(i) = "OnTimeOut"
  tokenhelp(i) = "This section is analyzed when a TimeOut event is issued.  Use SetTimeOut to initialize a timer.  TimeOut event is host-based, i.e. the PC computer generates TimeOut on its own."
  tokenparam(i) = "()"
  tokencolor(i) = 3
  i = i + 1
  
  tokenlist(i) = "SetTimeOut"
  tokenhelp(i) = "Define interval at which timeOut are issued and the section OnTimeOut executed. The section is executed with the host computer precision, and if a timeOut event is not completed, other timeOut will be skipped."
  tokenparam(i) = "(d)"
  tokencolor(i) = 3
  i = i + 1

  tokenlist(i) = "ClearTimeOut"
  tokenhelp(i) = "Stop generating host's time out events."
  tokenparam(i) = "()"
  tokencolor(i) = 3
  i = i + 1
  
  tokenlist(i) = "OnEvent"
  tokenhelp(i) = "This section is analyzed when the variable Var_Number is changed. OnEvent are RCX-generated events, when the variable is changed, and must be initialized using SetEvent (in the SPRIRIT.ocx doc, it is refered to OnVariableChange)."
  tokenparam(i) = "(v)"
  tokencolor(i) = 3
  i = i + 1
  
  tokenlist(i) = "SetEvent"
  tokenhelp(i) = "Define what element of the RCX to spy.  As soon as this element change, the section OnEvent is executed.  On the current version of the Spirit.ocx, only variable 0 can be spied."
  tokenparam(i) = "(s,n,d)"
  tokencolor(i) = 3
  i = i + 1
  
  tokenlist(i) = "ClearEvent"
  tokenhelp(i) = "Stop spying the element. ClearEvent is used in conjunction with the section 'OnEvent'.  All your program should finish at some point with a ClearEvent command to make sure the PC won't continue to spy the RCX after you left."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 3
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX <-> Host -" + CommentClosing
  tokenhelp(i) = "These commands are intended to download and start a firmware in the RCX and to retrieve information from the RCX."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "DownloadFirmware"
  tokenhelp(i) = "Send the content of the firmware stored in 'file' toward the RCX.  This is necessary in order to program the RCX."
  tokenparam(i) = "(f)"
  tokencolor(i) = 4.1
  i = i + 1
  
  tokenlist(i) = "UnlockFirmware"
  tokenhelp(i) = "This command send the password to the RCX so that it can be controlled by your host computer.  The password is 'Do you byte, when I knock?'"
  tokenparam(i) = "(q)"
  tokencolor(i) = 4.1
  i = i + 1

  tokenlist(i) = "UnlockPBrick"
  tokenhelp(i) = "Used to retrieve the ROM version and the firmware version loaded in the PBrick.  This command can only succeed if the UnlockFirmware command succeeded."
  tokenparam(i) = "()"
  tokencolor(i) = 4.2
  i = i + 1

  tokenlist(i) = "PBBattery"
  tokenhelp(i) = "Use to retrieve Battery level. Level indicated in milliVolts. Use >> or >>> to redirect the result."
  tokenparam(i) = "()"
  tokencolor(i) = 4.2
  i = i + 1

  tokenlist(i) = "SendPCMessage"
  tokenhelp(i) = "Use the PC to Send a message to the RCX.  Messages are number between 0 and 255, and can be read on the RCX by using the source 15: PBMessage."
  tokenparam(i) = "(2)"
  tokencolor(i) = 4.1
  i = i + 1
  
  tokenlist(i) = "Poll"
  tokenhelp(i) = "Use to retrieve info about one element of the RCX. Use > or >> to redirect the result."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 4.2
  i = i + 1

  tokenlist(i) = "UploadDatalog"
  tokenhelp(i) = "Use to retrieve the Datalog contained on the RCX. The first number define the starting point in the log, the second, the length to retrieve. Use >> or >>> to redirect the information (the upload is NOT limited in size). UploadDatalog(0,1) retrieve the position to be filled next in the datalog."
  tokenparam(i) = "(2,2)"
  tokencolor(i) = 4.2
  i = i + 1
  
  tokenlist(i) = "MemMap"
  tokenhelp(i) = "Used to retrieve the MemMap memory map of your RCX . Use >> or >>> to redirect the result."
  tokenparam(i) = "()"
  tokencolor(i) = 4.2
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- HOST  INPUT/OUTPUT COMMANDS -" + CommentClosing
  tokenhelp(i) = "The host computer can execute specific commands.  Some occurs only on the PC, and do not affect the RCX, such as START or PAUSE.  Other can retrieve information from the RCX, such as POLL."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "Options"
  tokenhelp(i) = "Set options of PRO-RCX. They are i) FORMAT, the format of the retrieved information from UploadDataLog, PBBattery, Poll, and MemMap. Format are VERBOSE (default; text is included), RAW (only numbers), and SHORT (only the value loggued -for UploadDatalog only); ii) START, which defines if the Communication port should be MUTECOM (muted while a program is started) or KEEPCOM (communications should continue, default). SHOWERRORS, is ACTIVE (show spifit.ocx errors; default) or INHIBITED."
  tokenparam(i) = "(Z,i)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = ">>"
  tokenhelp(i) = ">> is used to redirect information to a file or to the display.  It can be used after MemMap, PBBattery, Poll, or UploadDatalog. To send to the display, use the special file '*'."
  tokenparam(i) = "f"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = ">>>"
  tokenhelp(i) = ">>> is used to append information to the end of a file or to the display.  It can be used after MemMap, PBBattery, Poll, or UploadDatalog. Using the special file '*' does not append anything.  Use MsgBoxFrom to see the content of a file."
  tokenparam(i) = "f"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "Beep"
  tokenhelp(i) = "Makes a beep on the computer."
  tokenparam(i) = "()"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "InputBox"
  tokenhelp(i) = "Display a question and give space for the user to respond.  Use >> or >>> to redirect the result to a file."
  tokenparam(i) = "(q)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "MsgBox"
  tokenhelp(i) = "Display a user-define message on the computer screen or use >> or >>> to redirect the result.."
  tokenparam(i) = "(q)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "MsgBoxFrom"
  tokenhelp(i) = "display the content of a user-defined file on the computer screen or use >> or >>> to redirect the result."
  tokenparam(i) = "(f)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "Returned"
  tokenhelp(i) = "Reads the first line of the file 'Result.log', and compares the content with ValueOrString.  If equal, execute the following commands, otherwise skip to the next RETURNED or RETURNED END command.  Example: START('menu.exe') RETURNED '1' PLAYTONE(..) RETURNED END"
  tokenparam(i) = "w"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "Start"
  tokenhelp(i) = "Start an application. The application *must* finish with creating a file named 'Result.log'."
  tokenparam(i) = "(f)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "Pause"
  tokenhelp(i) = "The PC computer make a pause for the delay (in tens of ms) mentionned.  Useful to give downloads time to finish."
  tokenparam(i) = "(d)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "PBTxPower"
  tokenhelp(i) = "Set the operating range of the IR tower (0=short, 1=long)."
  tokenparam(i) = "(b)"
  tokencolor(i) = 4.3
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX INITIALISATION -" + CommentClosing
  tokenhelp(i) = "Initialization commands set or clear the value of an element of the RCX."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "SetDatalog"
  tokenhelp(i) = "Initialize the size of the datalog in the RCX memory.  There must be enough free memory left (check with MemMap)."
  tokenparam(i) = "(l)"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "ClearSensorValue"
  tokenhelp(i) = "Erase the content of a sensor."
  tokenparam(i) = "(n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetSensorType"
  tokenhelp(i) = "Define the sensor's type (see Pop-up menu for a list)."
  tokenparam(i) = "(n,6)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetSensorMode"
  tokenhelp(i) = "Define the sensor's mode (see Pop-up menu for a list)."
  tokenparam(i) = "(n,o,y)"
  tokencolor(i) = 5.3
  i = i + 1

  tokenlist(i) = "ClearTimer"
  tokenhelp(i) = "Reset one RCX timer to zero. Timer number are 0 to 3."
  tokenparam(i) = "(n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetWatch"
  tokenhelp(i) = "Set the time (hours, minute) seen on the display of the RCX."
  tokenparam(i) = "(H,t)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SelectDisplay"
  tokenhelp(i) = "Show what to display on the RCX. 0 is the Watch, 1 to 3 are the sensors, and 4 to 6 are the outputs."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX INPUT-OUTPUT -" + CommentClosing
  tokenhelp(i) = "Commands that actually make the RCX do something by acting on its ouptuts (A, B, or C) or speaker or its memory (datalog)."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "PBPowerDownTime"
  tokenhelp(i) = "Set the time in minutes after which the RCX automatically shuts downé  Set to 0, the RCX will never shut down automatically."
  tokenparam(i) = "(t)"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "PBTurnOff"
  tokenhelp(i) = "Turn off the RCX now."
  tokenparam(i) = "()"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "AlterDir"
  tokenhelp(i) = "Change direction of rotation of motors listed."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "Float"
  tokenhelp(i) = "Stop rotation of motors listed."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1

  tokenlist(i) = "Off"
  tokenhelp(i) = "Stop rotation of motors listed."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1

  tokenlist(i) = "On"
  tokenhelp(i) = "Start the rotation of motors listed."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetFwd"
  tokenhelp(i) = "Change direction of rotation of motors listed to forward."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetRwd"
  tokenhelp(i) = "Change direction of rotation of motors listed to reverse."
  tokenparam(i) = "(m)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetPower"
  tokenhelp(i) = "Set the speed of rotation of motors listed."
  tokenparam(i) = "(m,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "Wait"
  tokenhelp(i) = "Wait for 10 ms times the number in the element of the RCX."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "DatalogNext"
  tokenhelp(i) = "Add another value to the RCX log."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 5.3
  i = i + 1

  tokenlist(i) = "PlayTone"
  tokenhelp(i) = "Play a tone of the frequency in the first number for a duration contained in the second number."
  tokenparam(i) = "(l,d)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "PlaySystemSound"
  tokenhelp(i) = "Play one of the 5 pre-recorded RCX sound."
  tokenparam(i) = "(n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "ClearPBMessage"
  tokenhelp(i) = "Clear the last message received."
  tokenparam(i) = "()"
  tokencolor(i) = 5.2
  i = i + 1
    
  tokenlist(i) = "SendPBMessage"
  tokenhelp(i) = "Send a message given by the element of the RCX to another RCX.  Messages are number from 0 to 255. Use source 15 to read messages."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 5.2
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX PROGRAM STRUCTURE -" + CommentClosing
  tokenhelp(i) = "These commands are used to define the content of programs that will be downloaded to the RCX."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "SelectPrgm"
  tokenhelp(i) = "Set the program currently used (from 0 to 4). Programs are zones that can contain up to 10 tasks and up to 8 subs."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "DeleteTask"
  tokenhelp(i) = "Erase the content of a task from the RCX memory."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "DeleteAllTasks"
  tokenhelp(i) = "Erase the content of all tasks."
  tokenparam(i) = "()"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "DeleteSub"
  tokenhelp(i) = "Erase the content of a sub from the RCX memory."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "DeleteAllSubs"
  tokenhelp(i) = "Erase the content of all subs."
  tokenparam(i) = "()"
  tokencolor(i) = 5.1
  i = i + 1
  
  tokenlist(i) = "StartTask"
  tokenhelp(i) = "Start running the task."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "StopTask"
  tokenhelp(i) = "Stop running the task."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "StopAllTasks"
  tokenhelp(i) = "Stop running all tasks."
  tokenparam(i) = "()"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "GoSub"
  tokenhelp(i) = "Run a subprocedure."
  tokenparam(i) = "(P)"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "BeginOfTask"
  tokenhelp(i) = "Start the definition of a task. Task number ranges from 0 to 9."
  tokenparam(i) = "(P)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "EndOfTask"
  tokenhelp(i) = "End the definition of a task. Please add a delay after the completion of a task to allow the RCX to receive the packet (using Pause or MsgBox for example)."
  tokenparam(i) = "()"
  tokencolor(i) = 4.1
  i = i + 1
  
  tokenlist(i) = "BeginOfSub"
  tokenhelp(i) = "Start the definition of a Sub.  Sub number ranges from 0 to 7."
  tokenparam(i) = "(P)"
  tokencolor(i) = 4.3
  i = i + 1
  
  tokenlist(i) = "EndOfSub"
  tokenhelp(i) = "End the definition of a sub. Please add a delay after the completion of a sub to allow the RCX to receive the packet (using MsgBox for example)."
  tokenparam(i) = "()"
  tokencolor(i) = 4.1
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX FLOW OF COMMANDS -" + CommentClosing
  tokenhelp(i) = "Structural commands for loops, condition, and operations."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "If"
  tokenhelp(i) = "Test the condition and execute the following commands if true."
  tokenparam(i) = "(s,n,R,s,n)"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "Else"
  tokenhelp(i) = "If the if is false, execute the command that follow."
  tokenparam(i) = ""
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "EndIF"
  tokenhelp(i) = "End of the If structure."
  tokenparam(i) = "()"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "Loop"
  tokenhelp(i) = "Repeat the number of time specified in the element of the RCX.  If you specify a loop of 2,0, it will be an infinit loop."
  tokenparam(i) = "(s,n)"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "EndLoop"
  tokenhelp(i) = "End of the Loop structure"
  tokenparam(i) = "()"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "While"
  tokenhelp(i) = "Repeat while the condition is true."
  tokenparam(i) = "(s,n,R,s,n)"
  tokencolor(i) = 5.2
  i = i + 1
  
  tokenlist(i) = "EndWhile"
  tokenhelp(i) = "End of the While structure"
  tokenparam(i) = "()"
  tokencolor(i) = 5.2
  i = i + 1
  
Rem /////////////////////////////////////////////
Rem /////////////////////////////////////////////

  tokenlist(i) = CommentOpening + "- RCX ARITHMETIC -" + CommentClosing
  tokenhelp(i) = "Calculations."
  tokenparam(i) = ""
  tokencolor(i) = 1
  i = i + 1
  
  tokenlist(i) = "AbsVar"
  tokenhelp(i) = "Give to variable the absolute value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "AndVar"
  tokenhelp(i) = "And variable and the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "DivVar"
  tokenhelp(i) = "Divide variable the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "MulVar"
  tokenhelp(i) = "Multiply variable the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "OrVar"
  tokenhelp(i) = "Or variable and the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SetVar"
  tokenhelp(i) = "Give to variable the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SgnVar"
  tokenhelp(i) = "Give to variable the signe of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SubVar"
  tokenhelp(i) = "Substract to variable the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  tokenlist(i) = "SumVar"
  tokenhelp(i) = "Add to variable the value of the element of the RCX."
  tokenparam(i) = "(v,s,n)"
  tokencolor(i) = 5.3
  i = i + 1
  
  Rem don't forget to update the const nbertoken
  Rem at the begining of this module
End Sub

