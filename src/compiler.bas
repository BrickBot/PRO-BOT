Attribute VB_Name = "compiler"
Option Explicit

Rem ==================================
Rem event related variables
Rem ==================================
Public AnEvent As Boolean      'OnEvent()
Public ATimeOut As Boolean     'OnTimeout()
Public AStart As Boolean       'Start()
Public AnEvent1 As Integer, AnEvent2 As Integer
Public DatalogSize As Integer  'not used...

Rem ==================================
Rem global variables
Rem ==================================
Public OptionShowErrors As Boolean    'InhibitSpiritErrors
Public OptionFormat As String   'type of output
Public OptionStart As String    'should start mute the com port?
Public Output As String         'for pipes commands

Sub interpret(text As String)
Dim inBegin As Boolean, bas As Integer
Dim token As String, i As Long, j As Variant
Dim k As Long, result As String, temp As String
Dim k2 As Integer, l As Integer, temp2 As String
Dim id As Double, file As String
Dim klong1 As Long, m As Integer, TotalLength As Integer
Dim start As Integer, length As Integer, plus As Integer

On Error Resume Next

inBegin = False
OptionFormat = "verbose"
OptionStart = "keepcom"
ChDir CurrentPath

With FMainform.Lego
i = 1
Do While (i <= Len(text))
 progress.ProgressBar.Value = i / Len(text)
 token = GetToken(text, i)
 
 Rem check that only downloadable commands are in programs
 If inBegin Then
   j = InList(token, tokenlist(), nbrtoken)
   If tokencolor(j) <> 5.2 And tokencolor(j) <> 5.3 And Mid(token, 1, 5) <> "endof" Then
     Warning 32, 0, token
     Exit Sub
   End If
 End If
 
 Rem interpret each command in turn
 Select Case token
  
  Case ">>"
    If Output = "" Then
      Warning 54, 0, ""
    Else
      temp = sGetToken(text, i)
      If temp = "*" Then
        MsgBox Output, vbOKOnly, App.Title
      Else
        Open temp For Output As #1
        Print #1, Output
        Close #1
      End If
    End If
    Output = ""
    
  Case ">>>"
    If Output = "" Then
      Warning 54, 0, ""
    Else
      temp = sGetToken(text, i)
      If temp = "*" Then
        MsgBox Output, vbOKOnly, App.Title
      Else
        Open temp For Append As #1
        Print #1, Output
        Close #1
      End If
    End If
    Output = ""
    
  Case "sendpcmessage"
    j = nGetToken(text, i)
    sendpc (j)
    
  Case "options"
    result = GetToken(text, i)
    If result = "format" Then
      OptionFormat = GetToken(text, i)
    ElseIf result = "start" Then
      OptionStart = GetToken(text, i)
    ElseIf result = "showerrors" Then
      If GetToken(text, i) = "inhibited" Then
        OptionShowErrors = False
      End If
    Else
      Warning 60, 0, result
      Exit Sub
    End If

  Case "downloadfirmware"
    file = sGetToken(text, i)
    If Dir(file) = "" Then
      Warning 51, 0, file
      Exit Sub
    End If
    j = .DownloadFirmware(file)
  
  Case "unlockfirmware"
    j = .UnlockFirmware(sGetToken(text, i))
  
  Case "unlockpbrick"
    j = .UnlockPBrick()
    MsgBox j, vbOKOnly, "Result of UnlockPBrick"
  
  Case "memmap"
    j = .MemMap()
    If j(0) <> 0 Then
      Output = ""
      Select Case OptionFormat
        Case "verbose"
          Rem the 8 subs of each 5 programs
          Output = Output & "====Size of the Subs=====" & Chr(13) & Chr(10)
          bas = j(1)
          For l = 0 To 4
            Output = Output + "Prg: " & l + 1 & ": "
            For k = 0 To 7
              k2 = j(l * 8 + k + 2) - bas
              bas = j(l * 8 + k + 2)
              Output = Output & k2 & ", "
            Next k
            Output = Output & Chr(13) & Chr(10)
          Next l
          Rem the 10 tasks of each 5 programs
          Output = Output & "====Size of the Tasks=====" & Chr(13) & Chr(10)
          bas = j(41)
          For l = 0 To 4
            Output = Output + "Prg: " & l + 1 & ": "
            For k = 0 To 9
              k2 = j(l * 10 + k + 42) - bas
              bas = j(l * 10 + k + 42)
              Output = Output & k2 & ", "
            Next k
            Output = Output & Chr(13) & Chr(10)
          Next l
          Rem the datalog
          Output = Output & "====Size of the Datalog=====" & Chr(13) & Chr(10)
          Output = Output & (j(92) - j(91)) / 3 - 1 & " items on  " & (j(93) - j(91)) / 3 - 1 & " (" & j(92) - j(91) - 3 & " bytes out of " & j(93) - j(91) - 3 & ")" & Chr(13) & Chr(10)
          Output = Output & "====Remaining space=====" & Chr(13) & Chr(10)
          Output = Output & j(94) - j(93) & " out of 6144."
        Case "raw", "short"
          For k = 1 To 94
            Output = Output & j(k) & " "
          Next k
        Case Else
          Output = "Unknown format"
      End Select
    Else
      MsgBox "Error in the transmission."
    End If
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) <> ">>" Then
      Warning 53, 0, token
    End If

  Case "inputbox"
    Output = InputBox(sGetToken(text, i))
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) <> ">>" Then
      Warning 53, 0, token
    End If
    
  Case "msgbox"
    DoEvents        'in case a message is coming
    Output = sGetToken(text, i)
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) = ">>" Then
      Warning 55, 0, ""
    Else
      j = MsgBox(Output)
    End If
  
  Case "msgboxfrom"
    DoEvents
    Output = ""
    temp = sGetToken(text, i)
    If Dir(CurrentPath + "\" + temp) = "" Then
      Warning 62, 0, temp
    Else
      Open temp For Input As #1
      Do While Not EOF(1)
        Line Input #1, temp2
        Output = Output + temp2 & Chr(13) & Chr(10)
      Loop
      Close #1
      Rem look forward for the presence of redirection
      k = i
      result = GetToken(text, k)
      If Mid(result, 1, 2) = ">>" Then
        Warning 55, 0, ""
      Else
        j = MsgBox(Output, vbOKOnly, "Content of file " & temp)
      End If
    End If

  Case "pause"
    PauseElapsed = False
    FMainform.Timer2.Interval = nGetToken(text, i) * 10
    Do
      DoEvents
    Loop Until PauseElapsed
    
  Case "pbbattery"
    k2 = .PBBattery()
    Select Case OptionFormat
      Case "verbose"
        Output = Int(k2 / 1000) & " volts"
      Case "raw", "short"
        Output = k2
      Case Else
        Output = "Unknown format"
    End Select
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) <> ">>" Then
      Warning 53, 0, token
    End If
  
  Case "poll"
    Output = LTrim(RTrim(.Poll(nGetToken(text, i), nGetToken(text, i))))
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) <> ">>" Then
      Warning 53, 0, token
    End If

  Case "returned"
    result = ""
    If Dir(CurrentPath + "\result.log") <> "" Then
      Open CurrentPath + "\result.log" For Input As #1
      Line Input #1, result
      Close #1
    End If
    Do
      token = GetToken(text, i)
      If token <> "end" And result <> Mid(token, 2, Len(token) - 2) Then
        Do While GetToken(text, i) <> "returned" And i <= Len(text)
        Loop
      End If
    Loop Until token = "end" Or result = Mid(token, 2, Len(token) - 2) Or i > Len(text)

  Case "start"
    FMainform.StartOn.Visible = True
    Rem erasing previous results
    If Dir(CurrentPath + "\result.log") <> "" Then
      Kill CurrentPath + "\result.log"
    End If
    Rem if mutecom on
    If OptionStart = "mutecom" Then
      FMainform.Lego.CloseComm
      FMainform.startstatus.Visible = True
    End If
    Rem start the named program
    temp = sGetToken(text, i)
    id = Shell(temp, vbNormalFocus)
    DoEvents
    
    If id <> 0 Then
      Rem waiting for target file
      Do
        DoEvents
      Loop Until Dir(CurrentPath + "\result.log") <> ""
      Rem remove event on label
      FMainform.StartOn.Visible = False
      Rem if muted, restart com
      If OptionStart = "mutecom" Then
        FMainform.Lego.InitComm
        FMainform.startstatus.Visible = False
      End If
    Else
      Warning 62, 0, temp
    End If
    
  Case "uploaddatalog"
    Output = ""
    start = nGetToken(text, i)
    length = nGetToken(text, i)
    TotalLength = length
    If length Mod 50 = 0 Then plus = 0 Else plus = 1
    If length > 50 Then
      FMainform.lblhelp = "Retrieving... "
      FMainform.lblhelp.Refresh
    End If

    For m = 1 To Int(length / 50) + plus
      If TotalLength > 50 Then
        FMainform.lblhelp = FMainform.lblhelp & (m - 1) * 50 + start & "... "
        FMainform.lblhelp.Refresh
      End If
      j = .UploadDatalog((m - 1) * 50 + start, min(50, length))
      length = length - 50

      If IsArray(j) Then
        Select Case OptionFormat
          Case "verbose"
            For k = LBound(j, 2) To UBound(j, 2)
              Output = Output & Format(j(0, k), "@@") & " " & Format(j(1, k), "@@") & ":" & " " & Format(j(2, k), "@@@@@@") & "   "
              If ((k - LBound(j, 2) + 1) Mod 5 = 0) Then
                Output = Output & Chr(13) & Chr(10)
              End If
            Next k
          Case "raw"
             For k = LBound(j, 2) To UBound(j, 2)
               Output = Output & j(0, k) & " " & j(1, k) & " " & j(2, k) & " "
             Next k
          Case "short"
             For k = LBound(j, 2) To UBound(j, 2)
               Output = Output & j(2, k) & " "
             Next k
          Case Else
             Output = "Unknown format: " & OptionFormat
        End Select
      Else
        Warning 52, 0, "UploadDatalog"
        Exit For
      End If
    Next m
    Rem look forward for the presence of redirection
    k = i
    temp = GetToken(text, k)
    If Mid(temp, 1, 2) <> ">>" Then
      Warning 53, 0, token
    End If
  
  Case "setdatalog"
    DatalogSize = nGetToken(text, i)
    j = .SetDatalog(DatalogSize)
  
  Case "clearevent"
    AnEvent = False
    FMainform.EventOn.Visible = False
    j = .ClearEvent(nGetToken(text, i), nGetToken(text, i))
  
  Case "setevent"
    AnEvent = True
    FMainform.EventOn.Visible = True
    k2 = nGetToken(text, i)
    AnEvent1 = nGetToken(text, i)
    AnEvent2 = nGetToken(text, i)
    j = .SetEvent(k2, AnEvent1, AnEvent2)

  Case "clearsensorvalue"
    j = .ClearSensorValue(nGetToken(text, i))
  
  Case "setsensortype"
    j = .SetSensorType(nGetToken(text, i), nGetToken(text, i))
  
  Case "setsensormode"
    j = .SetSensorMode(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "cleartimer"
    j = .ClearTimer(nGetToken(text, i))
  
  Case "cleartimeout"
    ATimeOut = False
    FMainform.TimeoutOn.Visible = False
    FMainform.Timer1.Interval = 0
  
  Case "settimeout"
    ATimeOut = True
    FMainform.TimeoutOn.Visible = True
    FMainform.Timer1.Interval = nGetToken(text, i) * 10
  
  Case "setwatch"
    j = .SetWatch(nGetToken(text, i), nGetToken(text, i))
  
  Case "alterdir"
    j = .AlterDir(GetToken(text, i))
  
  Case "float"
    j = .Float(GetToken(text, i))
  
  Case "on"
    j = .On(GetToken(text, i))
  
  Case "off"
    j = .Off(GetToken(text, i))
  
  Case "setfwd"
    j = .SetFwd(GetToken(text, i))
  
  Case "setrwd"
    j = .SetRwd(GetToken(text, i))
  
  Case "wait"
    j = .Wait(nGetToken(text, i), nGetToken(text, i))
  
  Case "setpower"
    j = .SetPower(GetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "datalognext"
    j = .DatalogNext(nGetToken(text, i), nGetToken(text, i))
  
  Case "pbpowerdowntime"
    j = .PBPowerdownTime(nGetToken(text, i))
  
  Case "pbturnoff"
    j = .PBTurnOff()
  
  Case "pbtxpower"
    j = .PBTxPower(nGetToken(text, i))
  
  Case "playtone"
    j = .PlayTone(nGetToken(text, i), nGetToken(text, i)) * 10
  
  Case "playsystemsound"
    j = .PlaySystemSound(nGetToken(text, i))
  
  Case "selectdisplay"
    j = .SelectDisplay(nGetToken(text, i), nGetToken(text, i))
  
  Case "sendpbmessage"
    j = .SendPBMessage(nGetToken(text, i), nGetToken(text, i))
  
  Case "clearpbmessage"
    j = .ClearPBMessage()
  
  Case "selectprgm"
    j = .SelectPrgm(nGetToken(text, i))
  
  Case "deletetask"
    j = .DeleteTask(nGetToken(text, i))
  
  Case "deletealltasks"
    j = .DeleteAllTasks()
  
  Case "deletesub"
    j = .DeleteSub(nGetToken(text, i))
  
  Case "deleteallsubs"
    j = .DeleteAllSubs()
  
  Case "starttask"
    j = .StartTask(nGetToken(text, i))
  
  Case "stoptask"
    j = .StopTask(nGetToken(text, i))
  
  Case "stopalltasks"
    j = .StopAllTasks()
  
  Case "gosub"
    j = .GoSub(nGetToken(text, i))
  
  Case "loop"
    j = .Loop(nGetToken(text, i), nGetToken(text, i))
  
  Case "endloop"
    j = .EndLoop()
  
  Case "while"
    j = .While(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "endwhile"
    j = .EndWhile()
  
  Case "if"
    j = .If(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "else"
    j = .Else
  
  Case "endif"
    j = .EndIf
  
  Case "beginoftask"
    inBegin = True
    j = .BeginOfTask(nGetToken(text, i))
  
  Case "endoftask"
    inBegin = False
    j = .EndOfTask()
  
  Case "endoftasknodownload"
    inBegin = False
    j = .EndOfTaskNoDownload()
  
  Case "beginofsub"
    inBegin = True
    j = .BeginOfSub(nGetToken(text, i))
  
  Case "endofsub"
    inBegin = False
    j = .EndOfSub()
  
  Case "endofsubnodownload"
    inBegin = False
    j = .EndOfSubNoDownload()
  
  Case "setvar"
    j = .SetVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "sumvar"
    j = .SumVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "subvar"
    j = .SubVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "divvar"
    j = .DivVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "mulvar"
    j = .MulVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "sgnvar"
    j = .SgnVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "absvar"
    j = .AbsVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "andvar"
    j = .AndVar(nGetToken(text, i), nGetToken(text, i), nGetToken(text, i))
  
  Case "beep"
    Beep
 
 End Select
Loop
End With

End Sub

Sub sendpc(message As Integer)
 Dim wrapper As String, messcode As String
 Dim content As String, checksum As String
 
With FMainform.MSComm1
  .CommPort = FMainform.Combo1.ListIndex
  .Settings = "2400,O,8,1"
  Rem close lego port and open comm port
  FMainform.Lego.CloseComm
  .PortOpen = True
  Rem wrapper
  wrapper = Chr$(&H55) + Chr$(&HFF) + Chr$(&H0)
  Rem message code
  messcode = Chr$(&HF7) + Chr$(&H8)
  Rem content of the message
  content = Chr$(Val(message) Mod 256) + Chr$((Val(message) Mod 256) Xor &HFF)
  checksum = Chr$((Val(message) + &HF7) Mod 256) + Chr$(((Val(message) + &HF7) Mod 256) Xor &HFF)
  Rem send full message
  .Output = wrapper + messcode + content + checksum
  Rem close com port and reopen lego port
  .PortOpen = False
  FMainform.Lego.InitComm
End With

End Sub
