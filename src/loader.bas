Attribute VB_Name = "Loader"
Option Compare Text
Option Explicit

Rem recorded configuration variables; see loadcolors()
Public Const delay = 10000    'slow down the loading at startup
Public ConfigF(6) As String   'fonts
Public ConfigO(5) As String   'options
Public ConfigP(5) As String   'position of the window

Public DefineMain As Integer
Public EditorHasFocus As Boolean
Public PauseElapsed As Boolean
Public PopupBoolean As Boolean
Public FMainform As Form

Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Rem main procedure that starts the programs
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Sub Main()
  Dim i As Integer
  Dim fichier As String
  
  Rem show splash window and load main window
  Set FMainform = New FrmMain
  frmSplash.Show
  frmSplash.Refresh
  Load FMainform
  Load preproc
    
  frmSplash.lblLicenseTo = "Loading strings..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  LoadStrings
  LoadParam
  SetLanguage
  SetDefaults
  
  frmSplash.lblLicenseTo = "Loading colors..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  LoadCodemax
  SetCommondialog
   
  frmSplash.lblLicenseTo = "Loading list..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  SetList
  SetFile
    
  frmSplash.lblLicenseTo = "Loading config..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  LoadOptions
    
  frmSplash.lblLicenseTo = "Loading popup menu..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  LoadPosition
  SetPopup
    
  frmSplash.lblLicenseTo = "Opening communication port..."
  frmSplash.Refresh: For i = 1 To delay: Next i
  FMainform.Lego.ComPortNo = FMainform.Combo1.ListIndex
  FMainform.Lego.InitComm
    
Debug.Print FMainform.Lego.ComPortNo

  Rem display the main form
  FMainform.Show
  FMainform.CodeMax1(0).SetFocus
  Unload frmSplash
  
  If Command <> "" Then
    If Dir(FMainform.CommonDialog1.filename) <> "" Then
      fichier = Mid(Command, 2, Len(Command) - 2)
      FMainform.CodeMax1(0).openfile fichier
      FMainform.SSTab2.Caption = Mid(fichier, SearchBackward("\", fichier) + 1, 999)
      FilePath(0) = fichier
    End If
  End If
  
End Sub
  
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Rem Set* procedures
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Sub SetDefaults()
  OptionShowErrors = True 'inhibitSpiritErrors
  AnEvent = False   'an Event is on?
  ATimeOut = False  'a time-out event is on?
  AStart = False    'a Start is running?
End Sub

Sub SetLanguage()
  Dim str2 As String, str3 As String
  Dim str4 As String, str5 As String
  Dim i As Integer
  
  Dim lang As CodeMaxCtl.Language
  Set lang = New CodeMaxCtl.Language
  lang.CaseSensitive = True
  
  lang.CaseSensitive = False
  For i = 1 To nbrtoken
    If tokencolor(i) < 3 Then
      str2 = str2 + tokenlist(i) + Chr(10)
    ElseIf tokencolor(i) < 4 Then
      str3 = str3 + tokenlist(i) + Chr(10)
    ElseIf tokencolor(i) < 5 Then
      str4 = str4 + tokenlist(i) + Chr(10)
    ElseIf tokencolor(i) < 6 Then
      str5 = str5 + tokenlist(i) + Chr(10)
    End If
  Next i
  lang.ScopeKeywords1 = ""
  lang.ScopeKeywords2 = ""
  lang.Operators = str2
  lang.Keywords = str3
  lang.TagElementNames = str4 + Chr(10) + str5
  
  lang.Style = cmLangStyleProcedural
  lang.SingleLineComments = ""
  lang.MultiLineComments1 = "{"
  lang.MultiLineComments2 = "}"
  lang.StringDelims = Chr(34) + Chr(10) + "'"
    
  Dim globals As CodeMaxCtl.globals
  Set globals = New CodeMaxCtl.globals
  Call globals.RegisterLanguage("PRO-RCX", lang)
  
  Const HOTKEYF_CONTROL As Integer = 2
  Dim k As HotKey
  Set k = New HotKey
  Let k.Modifiers2 = 0
  Let k.Modifiers1 = HOTKEYF_CONTROL
  
  Rem create new commands for the menus intercepted in registeredcommand
  Call globals.RegisterCommand(1010, "SelectAllText", "Select all text")
  Let k.VirtKey1 = "A"
  globals.RegisterHotKey k, 1010
  
  Call globals.RegisterCommand(1011, "FindText", "Find text")
  Let k.VirtKey1 = "F"
  globals.RegisterHotKey k, 1011
  
  Call globals.RegisterCommand(1012, "FindAndReplace", "Find and Replace text")
  Let k.VirtKey1 = "R"
  globals.RegisterHotKey k, 1012
  
  Call globals.RegisterCommand(1000, "NewFile", "Create a new file")
  k.VirtKey1 = "N"
  globals.RegisterHotKey k, 1000
  
  Call globals.RegisterCommand(1001, "OpenFile", "Open a file")
  k.VirtKey1 = "O"
  globals.RegisterHotKey k, 1001
  
  Call globals.RegisterCommand(1002, "SaveFile", "Save the current editor")
  k.VirtKey1 = "S"
  globals.RegisterHotKey k, 1002
  
  Call globals.RegisterCommand(1003, "SaveAllFile", "Save all the editors")
  k.VirtKey1 = "L"
  globals.RegisterHotKey k, 1003
  
  Call globals.RegisterCommand(1004, "Print", "Print current editor")
  k.VirtKey1 = "P"
  globals.RegisterHotKey k, 1004

  FMainform.CodeMax1(0).Language = "PRO-RCX"
  FMainform.CodeMax1(1).Language = "PRO-RCX"
  FMainform.CodeMax1(2).Language = "PRO-RCX"
  
End Sub

Sub SetFile()
  Dim i As Integer
  StartPath = App.path
  For i = 0 To 2
    saved(i) = True
    FilePath(i) = StartPath + "\noname" & i & ".rcp"
  Next i
  DefineMain = -1
End Sub

Sub SetCommondialog()
  FMainform.CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlCFBoth Or cdlCCRGBInit
  FMainform.CommonDialog1.Filter = "RCP files (*.rcp)|*.rcp|all files (*.*)|*.*"
  FMainform.CommonDialog1.FilterIndex = 0
  FMainform.CommonDialog1.CancelError = True
End Sub

Sub SetList()
  Dim i As Integer
  For i = 1 To nbrtoken
    FMainform.List1.AddItem (tokenlist(i))
  Next i
  
End Sub

Sub SetPopup()
  Dim ligne As String, cont As String, r As Long
  Dim temp As String
  Dim categorie As Integer, no As Integer

  If Dir(StartPath + "\rcxdefine.rcp") <> "" Then
    Open StartPath + "\rcxdefine.rcp" For Input As #1
    Rem skip the first comments at the begining
    Do While Not EOF(1) And Mid(ligne, 1, 8) <> CommentOpening + "======="
      Line Input #1, ligne
    Loop
    Do While Not EOF(1)
      Line Input #1, ligne
        If Mid(ligne, 2, 4) = "====" Then
        ElseIf Mid(ligne, 1, 1) = CommentOpening Then
          no = 0
          categorie = categorie + 1
          If categorie = 11 Then
            Close #1
            Exit Sub
          End If
          FMainform.catego(categorie).Caption = Mid(ligne, 2, Len(ligne) - 3)
        Else
          r = 1
          temp = GetToken(ligne, r)
          Select Case categorie
            Case 1: If no <> 0 Then Load FMainform.empty1(no)
              FMainform.empty1(no).Caption = UCase(GetToken(ligne, r))
            Case 2: If no <> 0 Then Load FMainform.empty2(no)
              FMainform.empty2(no).Caption = UCase(GetToken(ligne, r))
            Case 3: If no <> 0 Then Load FMainform.empty3(no)
              FMainform.empty3(no).Caption = UCase(GetToken(ligne, r))
            Case 4: If no <> 0 Then Load FMainform.empty4(no)
              FMainform.empty4(no).Caption = UCase(GetToken(ligne, r))
            Case 5: If no <> 0 Then Load FMainform.empty5(no)
              FMainform.empty5(no).Caption = UCase(GetToken(ligne, r))
            Case 6: If no <> 0 Then Load FMainform.empty6(no)
              FMainform.empty6(no).Caption = UCase(GetToken(ligne, r))
            Case 7: If no <> 0 Then Load FMainform.empty7(no)
              FMainform.empty7(no).Caption = UCase(GetToken(ligne, r))
            Case 8: If no <> 0 Then Load FMainform.empty8(no)
              FMainform.empty8(no).Caption = UCase(GetToken(ligne, r))
            Case 9: If no <> 0 Then Load FMainform.empty9(no)
              FMainform.empty9(no).Caption = UCase(GetToken(ligne, r))
            Case 10: If no <> 0 Then Load FMainform.empty10(no)
              FMainform.empty10(no).Caption = UCase(GetToken(ligne, r))
          End Select
          no = no + 1
        End If
    Loop
    Close #1
    PopupBoolean = True
  Else
    MsgBox "RCXDefine.rcp could not be found; popup menu not available", vbOKOnly, "Error loading popup menu"
    PopupBoolean = False
  End If
End Sub

Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Rem diagnostic tool
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Sub DoDiagnostic()
  Dim j As Integer
  
  FMainform.DiagCheck(0).Visible = False
  FMainform.DiagCheck(1).Visible = False
  FMainform.DiagCheck(2).Visible = False
  FMainform.DiagCheck(3).Visible = False
  FMainform.diagno(0).Visible = False
  FMainform.diagno(1).Visible = False
  FMainform.diagno(2).Visible = False
  FMainform.diagno(3).Visible = False
  FMainform.Refresh
  
  If AnEvent Then
    j = FMainform.Lego.ClearEvent(AnEvent1, AnEvent2)
    FMainform.EventOn.Visible = False
    AnEvent = False
  End If

  If FMainform.Lego.TowerAndCableConnected Then
    FMainform.DiagCheck(0).Visible = True
  Else
    FMainform.diagno(0).Visible = True
  End If
  
  If FMainform.Lego.TowerAlive Then
     FMainform.DiagCheck(1).Visible = True
  Else
    FMainform.diagno(1).Visible = True
  End If
  
  If FMainform.Lego.PBAliveOrNot Then
    FMainform.DiagCheck(2).Visible = True
  Else
    FMainform.diagno(2).Visible = True
    FMainform.diagno(3).Visible = True
    Exit Sub
  End If
  
  If Mid(FMainform.Lego.UnlockPBrick, 7, 2) <> "00" Then
    FMainform.DiagCheck(3).Visible = True
  Else
    FMainform.diagno(3).Visible = True
  End If
  
End Sub

Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Rem Save* and Load* for recorded configuration
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Sub SaveCodemax()
  SaveSetting App.Title, "Font", "Bold", ConfigF(1)
  SaveSetting App.Title, "Font", "Italic", ConfigF(2)
  SaveSetting App.Title, "Font", "Name", ConfigF(3)
  SaveSetting App.Title, "Font", "Size", ConfigF(4)
  SaveSetting App.Title, "Font", "Strikethru", ConfigF(5)
  SaveSetting App.Title, "Font", "Underline", ConfigF(6)
  End Sub

Sub LoadCodemax()
  Dim i As Integer
  
  ConfigF(1) = GetSetting(App.Title, "Font", "Bold", "0")
  ConfigF(2) = GetSetting(App.Title, "Font", "Italic", "0")
  ConfigF(3) = GetSetting(App.Title, "Font", "Name", "Times New Roman")
  ConfigF(4) = GetSetting(App.Title, "Font", "Size", "12")
  ConfigF(5) = GetSetting(App.Title, "Font", "Strikethru", "0")
  ConfigF(6) = GetSetting(App.Title, "Font", "Underline", "0")
  For i = 0 To 2
    FMainform.CodeMax1(i).Font.Bold = CBool(ConfigF(1))
    FMainform.CodeMax1(i).Font.Italic = CBool(ConfigF(2))
    FMainform.CodeMax1(i).Font.Name = ConfigF(3)
    FMainform.CodeMax1(i).Font.Size = CInt(ConfigF(4))
    FMainform.CodeMax1(i).Font.Strikethrough = CBool(ConfigF(5))
    FMainform.CodeMax1(i).Font.Underline = CBool(ConfigF(6))
  Next i
End Sub

Sub SavePosition()
  SaveSetting App.Title, "Position", "Maximized", ConfigP(1)
  SaveSetting App.Title, "Position", "top", ConfigP(2)
  SaveSetting App.Title, "Position", "left", ConfigP(3)
  SaveSetting App.Title, "Position", "height", ConfigP(4)
  SaveSetting App.Title, "Position", "width", ConfigP(5)
  
End Sub

Sub LoadPosition()
  ConfigP(1) = GetSetting(App.Title, "Position", "Maximized", "0")
  ConfigP(2) = GetSetting(App.Title, "Position", "top", "50")
  ConfigP(3) = GetSetting(App.Title, "Position", "left", "50")
  ConfigP(4) = GetSetting(App.Title, "Position", "height", "8500")
  ConfigP(5) = GetSetting(App.Title, "Position", "width", "6200")
  FMainform.WindowState = ConfigP(1)
  FMainform.Top = ConfigP(2)
  FMainform.Left = ConfigP(3)
  FMainform.Height = ConfigP(4)
  FMainform.Width = ConfigP(5)
End Sub

Sub SaveOptions()
  SaveSetting App.Title, "Options", "ComPort", ConfigO(1)
  SaveSetting App.Title, "Options", "Compiling", ConfigO(2)
  SaveSetting App.Title, "Options", "HelpOn", ConfigO(3)
  SaveSetting App.Title, "Options", "AutoSave", ConfigO(4)
  SaveSetting App.Title, "Options", "Load_RCXDefine.rcp", ConfigO(5)
End Sub

Sub LoadOptions()
  ConfigO(1) = GetSetting(App.Title, "Options", "ComPort", "1")
  ConfigO(2) = GetSetting(App.Title, "Options", "Compiling", "1")
  ConfigO(3) = GetSetting(App.Title, "Options", "HelpOn", "1")
  ConfigO(4) = GetSetting(App.Title, "Options", "AutoSave", "0")
  ConfigO(5) = GetSetting(App.Title, "Options", "Load_RCXDefine.rcp", "1")
  FMainform.Combo1.ListIndex = Val(ConfigO(1))
  FMainform.Check1.Value = ConfigO(2)
  FMainform.Check2.Value = ConfigO(3)
  FMainform.Check3.Value = ConfigO(4)
  FMainform.Check4.Value = ConfigO(5)
End Sub

Rem ////////////////////////////////////////
Rem a general purpose sub used to send
Rem error messages or warning to the user.
Rem no is the error number, further can be
Rem another information,
Rem ////////////////////////////////////////

Sub Warning(no As Integer, further As Integer, token As String)
  Dim ErrorType As String
  Dim ErrorTxt As String
  
  Rem note: if the warning is critical; don't forget to add exit sub where it occured
  Rem determine the category of error
  Select Case no
    Case 0:        ErrorType = "Preprocessor completed with success!"
    Case 1 To 19:  ErrorType = "Preprocessor error #" & no & ": "
    Case 20 To 29: ErrorType = "Syntax error #" & no & ":       "
    Case 30 To 49: ErrorType = "Compiler error #" & no & ":     "
    Case 50 To 69: ErrorType = "Run-time error #" & no & ":     "
    Case 70 To 89: ErrorType = "Event warning #" & no & ":      "
    Case 90 To 99: ErrorType = "Phantom.dll error #" & no & ":  "
  End Select
  
  Rem determine specific error message
  Select Case no
   Case 0:  ErrorTxt = ""
   
   Case 1:  ErrorTxt = "Define not ended by ): " & token & "."
   Case 2:  ErrorTxt = "Missing identifier in the " & further & "th DEFINE."
   Case 3:  ErrorTxt = "Insert file missing or not quoted: " & token & "."
   Case 4:  ErrorTxt = "Inserted file name does not exist: " & token & "."
   Case 5:  ErrorTxt = "Could not find RCXDefine.rcp in the path: " & token & ". Continuing..."
   Case 6:  ErrorTxt = "Value improperly quoted in the " & further & "th DECLARE: " & token & "."
   Case 7:  ErrorTxt = "STRING Value should be quoted in the " & further & "th DECLARE: " & token & "."
   Case 8:  ErrorTxt = "Source type missing (number between 1 and 16): " & token & "."
   Case 9:  ErrorTxt = "File " & token & " already inserted (cyclic insert?).  Skipping..."
   Case 10: ErrorTxt = token & " not an identifier starting with a letter in the " & further & "th Define. Cyclic define?"
   Case 11: ErrorTxt = token & " not an identifier starting with a letter in the " & further & "th Declare. Cyclic declare?"
   Case 12: ErrorTxt = "'" & token & "' not a numerical value."
   
   Case 20: ErrorTxt = "Unknown command " & token & "."
   Case 21: ErrorTxt = "Expected separator: " & token & "."
   Case 22: ErrorTxt = "Expected ' or " + Chr(34) + ": " & token & "."
   Case 23: ErrorTxt = "Expected number: " & token & "."
   Case 24: ErrorTxt = "No ( found in Declare.  Found " & token & "."

   Case 32: ErrorTxt = "No immediat or host command allowed inside of a sub/task: " & token & "."
   
   Case 50: ErrorTxt = "RCX not connected; rerun diagnostic."
   Case 51: ErrorTxt = "Firmware not found: " & token & "."
   Case 52: ErrorTxt = "Transmission of information failed during " & token & ". Check diagnostic."
   Case 53: ErrorTxt = "No >> or >>> for the information retrieved from " & token & ". Continuing..."
   Case 54: ErrorTxt = "No Output to redirect.  Use MemMap or another output generating command."
   Case 55: ErrorTxt = ">> or >>> used.  The actual output will not be send to MsgBox. Continuing..."
   Case 60: ErrorTxt = "Option " & token & " unknown.  Please check spelling."
   Case 61: ErrorTxt = "Section OnEvent's variable is not the one use in the SetEvent instruction: " & token & "."
   Case 62: ErrorTxt = "File " & token & " not found. Skip..."
   
   Case 70: ErrorTxt = "A preceding SetEvent is still on and is going to be stop."
   Case 71: ErrorTxt = "A preceding SetTimeOut is still on and is going to be stop."
   Case 72: ErrorTxt = "A preceding Start is still on and is going to be stop."
   
   Case 90: ErrorTxt = "Error no " & further & ": " & token
   
   Case Else: ErrorTxt = "Unknown error"
  End Select

  Rem displaying the error message
  FMainform.lblhelp = FMainform.lblhelp + ErrorType + ErrorTxt + Chr(13) + Chr(10)
End Sub

Function min(X As Integer, Y As Integer) As Integer
  If X < Y Then min = X Else min = Y
End Function


