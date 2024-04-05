Attribute VB_Name = "preprocessor"
Option Explicit

Rem various token separator
Public Const setletter = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_<>=&#%?"
Public Const setnumber = "0123456789"
Public Const setseparator = " (),.=-*^$@!\|]{}[;:/`~"   'and tab, return, etc
Public Const setspaces = " " 'and tab, return, etc

Rem global variable to report preprocessing errors
Public HasError As Boolean

Rem lists for the defined value
Public definedlist(1000) As String
Public definedas(1000) As String
Public nodefine As Integer
Public relatedlist(1000) As String
Public relatedas(1000) As String
Public relatedtype(1000) As String
Public norelate As Integer
Public noinsert As Integer
Public insertlist(1000) As String

Rem /////////////////////////
Rem boolean functions to know in
Rem what category falls a char
Rem /////////////////////////
Function IsSeparator(car As String) As Boolean
  Rem identifies if the car is a separator of the language
  If InStr(setseparator + Chr(vbKeyTab) + Chr(vbKeyReturn) + Chr(10), car) Then
    IsSeparator = True
  Else
    IsSeparator = False
  End If
End Function

Function Spaces(car As String) As Boolean
  If InStr(setspaces + Chr(vbKeyTab) + Chr(vbKeyReturn) + Chr(10) + Chr(13), car) Then
    Spaces = True
  Else
    Spaces = False
  End If
End Function

Function Letter(car As String) As Boolean
  If InStr(setletter, UCase(car)) Then
    Letter = True
  Else
    Letter = False
  End If
End Function

Function Number(car As String) As Boolean
  If InStr(setnumber, car) Then
    Number = True
  Else
    Number = False
  End If
End Function

Function Numbers(mot As String) As Boolean
  Dim i As Integer
  
  For i = 1 To Len(mot)
    If Not Number(Mid(mot, i, 1)) Then
      Numbers = False
      Exit Function
    End If
  Next i
  Numbers = True
End Function

Function Quoted(token As String) As Boolean
  Dim car1 As String, car2 As String
  car1 = Mid(token, 1, 1)
  car2 = Right(token, 1)
  
  If (car1 = "'") And (car2 = "'") Then
    Quoted = True
  ElseIf (car1 = Chr(34)) And (car2 = Chr(34)) Then
    Quoted = True
  Else
    Quoted = False
  End If
End Function

Function Questioned(token As String) As Boolean
  Dim car1 As String, car2 As String
    
  If (Mid(token, 1, 1) = "?") And (Right(token, 1) = "?") Then
    Questioned = True
  Else
    Questioned = False
  End If
End Function

Rem /////////////////////////
Rem look for an element in a list
Rem having a maximum of noelement elements
Rem /////////////////////////
Function InList(element As String, list() As String, noelement As Integer) As Integer
  Dim i As Integer
  
  Rem if found, return the position of element is the list
  For i = 1 To noelement
    If UCase(list(i)) = UCase(element) Then
      InList = i
      Exit Function
    End If
  Next
  InList = 0
End Function

Rem /////////////////////////
Rem text analyzer that extract "token"
Rem that is, the smallest unit of the language
Rem token are words, numbers,
Rem quoted text, and separators
Rem Comments are discarded.
Rem /////////////////////////
Function nGetToken(texte As String, ByRef i As Long) As Integer
  nGetToken = CInt(GetToken(texte, i))
End Function

Function sGetToken(texte As String, ByRef i As Long) As String
  Dim temp As String
  
  temp = GetToken(texte, i)
  sGetToken = Mid(temp, 2, Len(temp) - 2)
End Function

Function GetToken(ByRef texte As String, ByRef i As Long) As String
  Dim token As String
  
  Rem skip spaces and comment if any
  Do While (Spaces(Mid(texte, i, 1)) Or Mid(texte, i, 1) = CommentOpening) And i <= Len(texte)
    Rem skip spaces
    Do While Spaces(Mid(texte, i, 1)) And i <= Len(texte)
      i = i + 1
    Loop
    Rem skip comments
    If Mid(texte, i, 1) = CommentOpening Then
      Do
        i = i + 1
      Loop Until Mid(texte, i, 1) = CommentClosing Or i > Len(texte)
      i = i + 1
    End If
  Loop
  
  Rem read between single or double (chr(34)) quote as one token
  If Mid(texte, i, 1) = "'" Then
    Do
      token = token + Mid(texte, i, 1)
      i = i + 1
    Loop Until Mid(texte, i, 1) = "'" Or i > Len(texte)
    token = token + Mid(texte, i, 1)
    i = i + 1
    GetToken = token
    Exit Function
  ElseIf Mid(texte, i, 1) = Chr(34) Then
    Do
      token = token + Mid(texte, i, 1)
      i = i + 1
    Loop Until Mid(texte, i, 1) = Chr(34) Or i > Len(texte)
    token = token + Mid(texte, i, 1)
    i = i + 1
    GetToken = token
    Exit Function
  ElseIf Mid(texte, i, 1) = "?" Then
    Do
      token = token + Mid(texte, i, 1)
      i = i + 1
    Loop Until Mid(texte, i, 1) = "?" Or i > Len(texte)
    token = token + Mid(texte, i, 1)
    i = i + 1
    GetToken = token
    Exit Function
  Else
    Rem start collecting a token
    If Number(UCase(Mid(texte, i, 1))) Then
      Rem case it is a number
      Do While Number(Mid(texte, i, 1)) And i <= Len(texte)
        token = token + Mid(texte, i, 1)
        i = i + 1
      Loop
    ElseIf Letter(UCase(Mid(texte, i, 1))) Then
      Do While (Letter(Mid(texte, i, 1)) Or Number(Mid(texte, i, 1))) And i <= Len(texte)
        token = token + Mid(texte, i, 1)
        i = i + 1
      Loop
    Else
      token = Mid(texte, i, 1)
      i = i + 1
    End If
  End If
  GetToken = LCase(token)
End Function

Function LookUp(texte As String, ByRef i As Long) As String
  Dim tok As String, k As Integer, j As Long
    
  j = i
  tok = GetToken(texte, j)
  k = InList(tok, definedlist(), nodefine)
  If k <> 0 Then
    texte = Mid(texte, 1, i - 1) + definedas(k) + Mid(texte, j, Len(texte))
    tok = LookUp(texte, i)
  Else
    i = j
  End If
  LookUp = tok
End Function

Rem /////////////////////////
Rem pass0, pass1, pass2 are the
Rem preprocessing functions
Rem each receive a string as input, and
Rem produce a string as output, the
Rem treated program.
Rem errors empties the result, and call Warning
Rem /////////////////////////
Function PassSection(ByVal Text0 As String, section As String) As String
  Rem retrieve the correct section from the program
  Dim i As Long, token As String, actualsection As String
  Dim ligne As String
  
  i = 1
  PassSection = ""
  HasError = False
  actualsection = ""
  Do While i <= Len(Text0)
    token = GetToken(Text0, i)
    If token = "main" Or token = "onevent" Or token = "ontimeout" Then
      actualsection = token
      If token = "onevent" Then
        token = GetToken(Text0, i)
        If token <> AnEvent1 Then
          Warning 61, 0, token
          Exit Function
        End If
      End If
    End If
    If actualsection = "" Or actualsection = section Then
      PassSection = PassSection + token + " "
    End If
  Loop
  
End Function

Sub Pass0()
  nodefine = 4
  noinsert = 0
  norelate = 0
  definedlist(1) = "%comport%"
  definedas(1) = "'" & FMainform.Combo1.text & "'"
  definedlist(2) = "%workingpath%"
  definedas(2) = "'" & CurrentPath & "'"
  definedlist(3) = "%apppath%"
  definedas(3) = "'" & StartPath & "'"
  definedlist(4) = "%ProgName%"
  definedas(4) = "'" & FMainform.SSTab2.Caption & "'"
End Sub

Function Pass1(ByVal CodeMax1 As String) As String
  Dim i As Long, token As String, temp As String
  Dim j As Integer, k As Integer
  Dim ligne As String, pos As Long, file As String
  
  i = 1
  HasError = False
  Pass1 = CommentOpening + "--Pass 1- Inserting files and definitions--" + CommentClosing + Chr(13) + Chr(10)
  
  Do While i <= Len(CodeMax1)
    Rem showing progress
    If Int(i / (Len(CodeMax1) + 1) * 10) > j Then
      j = j + 1
      progress.ProgressBar.Value = i / Len(CodeMax1)
    End If
    
    pos = i
    token = LookUp(CodeMax1, i)
    Select Case token
      Case "insert":
        Rem inserting insert files
        token = LookUp(CodeMax1, i)
        If Not Quoted(token) Then
          Warning 3, 0, token
          Pass1 = Pass1 + token + "<--Error!"
          HasError = True
          Exit Function
        End If
        Rem memorizing this insert to avoid cyclic insert
        If InList(token, insertlist(), noinsert) Then
          Warning 9, 0, token
        Else
          Rem put the file in the list
          noinsert = noinsert + 1
          insertlist(noinsert) = token
          Rem check that the file exist
          If Dir(Mid(token, 2, Len(token) - 2)) = "" Then
            Warning 4, 0, token
            Pass1 = Pass1 + token + "<--Error!"
            HasError = True
            Exit Function
          End If
          Rem insert the file into pass1
          file = ""
          Open Mid(token, 2, Len(token) - 2) For Input As #1
          Do While Not EOF(1)
            Line Input #1, ligne
            file = file + ligne + Chr(13) + Chr(10)
          Loop
          Close #1
          CodeMax1 = Mid(CodeMax1, 1, pos - 1) + file + Mid(CodeMax1, i, Len(CodeMax1) - i + pos)
          i = pos
        End If
      
      Case "define":
        token = LookUp(CodeMax1, i)
        If token = "" Then
          Warning 2, nodefine, ""
          Pass1 = Pass1 + "define " + token + "<--Error!"
          HasError = True
          Exit Function
        End If
        If Not Letter(Mid(token, 1, 1)) Then
          Warning 10, nodefine, token
          Pass1 = Pass1 + "define " + token + "<--Error!"
          HasError = True
          Exit Function
        End If
        nodefine = nodefine + 1
        definedlist(nodefine) = token
        token = LookUp(CodeMax1, i)
        If token <> "(" Then
          Warning 24, 0, token
          HasError = True
          Exit Function
        End If
        Rem reading what's in between parenthesis
        token = ""
        file = LookUp(CodeMax1, i)
        Do While file <> ")" And i < Len(CodeMax1)
          If Questioned(file) Then
            temp = InputBox(Mid(file, 3, Len(file) - 3), App.Title)
            If LCase(Mid(file, 2, 1)) = "s" Then
              temp = "'" + temp + "'"
            End If
            token = Merge(token, temp)
          Else
            token = Merge(token, file)
          End If
          file = LookUp(CodeMax1, i)
        Loop
        If file <> ")" Then
          Warning 1, 0, file
          HasError = True
          Exit Function
        End If
        Rem record the definition
        definedas(nodefine) = token
      
      Case Else:
        Pass1 = Pass1 + token + " "
    End Select
  Loop
  Pass1 = Pass1 + Chr(13) + Chr(10) + CommentOpening + "--Pass 1 completed succesfully--" + CommentClosing

End Function

Function Pass2(ByVal Pass1 As String) As String
  Dim token As String, i As Long
  Dim j As Integer, param As String
  Dim k As Integer, pos As Integer
  Dim r As Long, temp As String
  
  i = 1
  HasError = False
  Pass2 = CommentOpening + "--Pass 2- Inserting declarations--" + CommentClosing + Chr(13) + Chr(10)
  Do While i <= Len(Pass1)
    If Int(i / (Len(Pass1) + 1) * 10) > j Then
      j = j + 1
      progress.ProgressBar.Value = i / Len(Pass1)
    End If
    
    token = GetToken(Pass1, i)
    
    Rem declare
    If token = "declare" Then
      norelate = norelate + 1
      
      token = GetToken(Pass1, i)
      relatedlist(norelate) = token
      If Not Letter(Mid(token, 1, 1)) Then
        Warning 11, norelate, token
        Pass2 = Pass2 + "declare " + token + "<--Error!"
        HasError = True
        Exit Function
      End If
      If Quoted(token) Then
        Warning 6, norelate, relatedlist(norelate)
        Pass2 = Pass2 + "declare " + token + "<--Error!"
        HasError = True
        Exit Function
      End If
      
      token = GetToken(Pass1, i)
      If Not Numbers(token) Then
        Warning 8, 0, token
        Pass2 = Pass2 + "declare " + relatedlist(norelate) + relatedtype(norelate) + "<--Error!"
        HasError = True
        Exit Function
      End If
      relatedtype(norelate) = token
      If Quoted(relatedtype(norelate)) Then
        Warning 6, norelate, relatedtype(norelate)
        Pass2 = Pass2 + "declare " + relatedlist(norelate) + relatedtype(norelate) + "<--Error!"
        HasError = True
        Exit Function
      Else
        relatedtype(norelate) = CInt(relatedtype(norelate))
      End If
      
      Rem the value of the declare
      token = GetToken(Pass1, i)
      relatedas(norelate) = token
      If Not Numbers(token) Then
        Warning 12, 0, token
        Pass2 = Pass2 + "declare " + relatedlist(norelate) + " " + relatedtype(norelate) + " " + token + "<--Error!"
        HasError = True
        Exit Function
      End If
      If Quoted(token) And relatedtype(norelate) <> 16 Then
        Warning 6, norelate, relatedlist(norelate)
        Pass2 = Pass2 + "declare " + relatedlist(norelate) + " " + relatedtype(norelate) + " " + token + "<--Error!"
        HasError = True
        Exit Function
      ElseIf (Not Quoted(token)) And relatedtype(norelate) = 16 Then
        Warning 7, norelate, relatedlist(norelate)
        Pass2 = Pass2 + "declare " + relatedlist(norelate) + " " + relatedtype(norelate) + " " + token + "<--Error!"
        HasError = True
        Exit Function
      End If
      Rem record the token number
    
    Else
      k = InList(token, tokenlist(), nbrtoken)
      If k <> 0 Then
        Rem a command starts here
        param = tokenparam(k)
        pos = 0
        Pass2 = Pass2 + token + " "
      Else
        Rem remaining parameters of a command
        pos = pos + 1
        k = InList(token, relatedlist(), norelate)
        If k <> 0 Then
          If Mid(param, pos, 1) = "s" Then
            Pass2 = Pass2 + relatedtype(k) + " , " + relatedas(k) + " "
            pos = pos + 2
          Else
            Pass2 = Pass2 & relatedas(k) & " "
          End If
        Else
          Pass2 = Pass2 + token + " "
        End If
      End If
    End If
  Loop
  Pass2 = Pass2 + Chr(13) + Chr(10) + CommentOpening + "--Pass 2 completed succesfully--" + CommentClosing

End Function

Function Syntax_Check(ByVal Pass2 As String) As String
  Dim token As String, i As Long, k As Integer
  Dim j As Integer
  
  i = 1
  HasError = False
  Do While i <= Len(Pass2)
    If Int(i / (Len(Pass2) + 1) * 10) > j Then
      j = j + 1
      progress.ProgressBar.Value = i / Len(Pass2)
    End If
    token = GetToken(Pass2, i)
    Syntax_Check = Syntax_Check + token + " "
    If token <> "" Then
      k = InList(token, tokenlist(), nbrtoken)
      If k = 0 Then
        Warning 20, 0, token
        Syntax_Check = Syntax_Check + token + "<--Error!"
        HasError = True
        Exit Function
      End If
      For j = 1 To Len(tokenparam(k))
        token = GetToken(Pass2, i)
        Select Case Mid(tokenparam(k), j, 1)
          Case "f", "q"
            If Not Quoted(token) Then
              Warning 22, 0, token
              Syntax_Check = Syntax_Check + token + "<--Error!"
              HasError = True
              Exit Function
            End If
            Syntax_Check = Syntax_Check + token + " "
          Case "v", "s", "n", "P", "t", "H", "l", "d", "7", "4", "2", "b", "6", "o", "y", "R"
            If Not Numbers(token) Then
              Warning 23, 0, token
              Syntax_Check = Syntax_Check + token + "<--Error!"
              HasError = True
              Exit Function
            End If
            Syntax_Check = Syntax_Check + token + " "
          Case "m", "i", "Z", "w"
            Rem string, nothing to do
            Syntax_Check = Syntax_Check + token + " "
          Case Else
            If token <> Mid(tokenparam(k), j, 1) Then
              Warning 21, 0, token
              Syntax_Check = Syntax_Check + token + "<--Error!"
              HasError = True
              Exit Function
            End If
        End Select
      Next j
    End If
  Loop

End Function

Sub SetProgress()
  Dim i As Integer
  
  progress.Top = FMainform.Top + FMainform.Height / 2 - progress.Height / 2
  progress.Left = FMainform.Left + FMainform.Width / 2 - progress.Width / 3
  For i = 0 To 2
    progress.check(i).Visible = False
    progress.status(i).Font.Bold = False
  Next i
  progress.Visible = True

End Sub

Sub ShowProgress(i As Integer)
  If i > 1 Then
    progress.status(i - 2).Font.Bold = False
    progress.check(i - 2).Visible = True
  End If
  progress.status(i - 1).Font.Bold = True
  progress.Refresh

End Sub

Function pass0to3(texte As String) As String
  Dim pass1result As String
  Dim pass2result As String
  Dim pass3result As String
  Dim i As Integer
  
  FMainform.lblhelp = ""
  SetProgress
  
  Rem if autoload RCXDefine.rcp
  If FMainform.Check4.Value = 1 Then
    If DefineMain = -1 Then
      If UCase(FilePath(FMainform.SSTab2.Tab)) <> UCase(StartPath + "\RCXDefine.rcp") Then
        Rem add insert rcxdefine.rcp
        If Dir(StartPath + "\RCXDefine.rcp") <> "" Then
          texte = "insert '" + StartPath + "\RCXDefine.rcp'" + texte
        Else
          Warning 5, 0, StartPath
        End If
      End If
    Else
      If UCase(FilePath(DefineMain)) <> UCase(StartPath + "\RCXDefine.rcp") Then
        Rem add insert rcxdefine.rcp
        If Dir(StartPath + "\RCXDefine.rcp") <> "" Then
          texte = "insert '" + StartPath + "\RCXDefine.rcp'" + texte
        Else
          Warning 5, 0, StartPath
        End If
      End If
    End If
  End If
  
  Rem built-in strings
  Pass0
  
  Rem pass 1- manage insert instructions
  ShowProgress (1)
  pass1result = Pass1(texte)
  If HasError Then 'error
    preproc.Pass1 = pass1result + Chr(13) + Chr(10) + CommentOpening + "-Pass 1 interrupted with error(s)-" + CommentClosing
    preproc.Pass2 = CommentOpening + "-no Pass 2 -" + CommentClosing
    progress.Visible = False
    Exit Function
  End If
  preproc.Pass1 = pass1result
  
  Rem pass 2- read define instructions
  ShowProgress (2)
  pass2result = Pass2(pass1result)
  If HasError Then
    preproc.Pass2 = pass2result + Chr(13) + Chr(10) + CommentOpening + "-Pass 2 interrupted with error(s)-" + CommentClosing
    progress.Visible = False
    Exit Function
  End If
  preproc.Pass2 = pass2result
  
  Rem pass 3- syntax check and compile
  ShowProgress (3)
  pass0to3 = Syntax_Check(pass2result)
  If HasError Then
    preproc.Pass3 = pass0to3 + Chr(13) + Chr(10) + CommentOpening + "-Syntax check interrupted with error(s)-" + CommentClosing
    progress.Visible = False
    Exit Function
  End If
  
  progress.Visible = False
  preproc.Pass3 = pass0to3
  Warning 0, 0, ""          'Success!
End Function

Function Merge(txt1 As String, txt2 As String) As String
  If Quoted(txt1) And Quoted(txt2) Then
    Merge = "'" & Mid(txt1, 2, Len(txt1) - 2) + Mid(txt2, 2, Len(txt2) - 2) & "'"
  Else
    Merge = txt1 + txt2
  End If
End Function
