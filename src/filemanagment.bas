Attribute VB_Name = "filemanagment"
Option Explicit

Rem location of files
Public FilePath(2) As String      ' each file`s location
Public StartPath As String        ' starting location
Public CurrentPath As String      ' current rcp program's location
Public saved(3) As Boolean        ' saved file?

Sub puttokenhelp(signe As Integer)
  Dim k As Integer
  Dim s As String, i As Integer
  
  With FMainform
    .lblsyntax = tokenlist(signe)
    For i = 1 To Len(tokenparam(signe))
      s = Mid(tokenparam(signe), i, 1)
      k = InList(s, paramlist(), nbreparam)
      If k <> 0 Then
        .lblsyntax = .lblsyntax + " " + paramhelp(k)
      Else
        .lblsyntax = .lblsyntax + " " + s
      End If
    Next i

    .lblhelp = tokenhelp(signe)
    Select Case tokencolor(signe)
      Case 1: .lblcategory.Visible = False
        For i = 0 To 2
          .thetype(i).Visible = False
        Next i
      Case 2: .lblcategory.Visible = True
        .lblcategory.Caption = "Preprocessor command"
        For i = 0 To 2
          .thetype(i).Visible = False
        Next i
      Case 3: .lblcategory.Visible = True
        .lblcategory.Caption = "Section identificator"
        For i = 0 To 2
          .thetype(i).Visible = False
        Next i
      Case 4 To 4.3: .lblcategory.Visible = True
        .lblcategory.Caption = "Host command"
        For i = 0 To 2
          .thetype(i).Visible = True
          .thetype(i).Value = 0
        Next i
        .thetype(0).Caption = "Send information"
        .thetype(1).Caption = "Retrieve information"
        .thetype(2).Caption = "Host only"
      
        If tokencolor(signe) = 4.1 Then
          .thetype(0).Value = 1
        ElseIf tokencolor(signe) = 4.2 Then
          .thetype(1).Value = 1
        ElseIf tokencolor(signe) = 4.3 Then
          .thetype(2).Value = 1
        End If
      Case 5 To 5.3: .lblcategory.Visible = True
        .lblcategory.Caption = "RCX command"
        For i = 0 To 1
          .thetype(i).Visible = True
          .thetype(i).Value = 0
        Next i
        .thetype(2).Visible = False
        .thetype(0).Caption = "Immediate command"
        .thetype(1).Caption = "Downloadable command"
        
        If tokencolor(signe) = 5.1 Then
          .thetype(0).Value = 1
        ElseIf tokencolor(signe) = 5.2 Then
          .thetype(1).Value = 1
        ElseIf tokencolor(signe) = 5.3 Then
          .thetype(0).Value = 1
          .thetype(1).Value = 1
        End If
    End Select
  End With
End Sub

Function saveAsfile(zone As String, fichier As String)
  
  FMainform.CommonDialog1.filename = fichier
  FMainform.CommonDialog1.ShowSave
  If FMainform.CommonDialog1.filename <> "" Then
    FMainform.CodeMax1(zone).savefile FMainform.CommonDialog1.filename, False
    FMainform.SSTab2.Caption = FMainform.CommonDialog1.FileTitle
    FilePath(FMainform.SSTab2.Tab) = FMainform.CommonDialog1.filename
  End If
End Function

Function savefile(zone As Integer, fichier As String)
  FMainform.CodeMax1(zone).savefile fichier, False
End Function

Sub openfile(zone As Integer)
  Dim fichier As String
  Dim ligne As String
  On Error GoTo errend
  
  FMainform.CommonDialog1.ShowOpen
  If Dir(FMainform.CommonDialog1.filename) <> "" Then
    fichier = FMainform.CommonDialog1.filename
    If FMainform.CommonDialog1.FileTitle <> "" Then
      FMainform.CodeMax1(zone).openfile fichier
      FMainform.SSTab2.Caption = FMainform.CommonDialog1.FileTitle
      FilePath(FMainform.SSTab2.Tab) = FMainform.CommonDialog1.filename
    End If
  Else
    MsgBox "File does not exist"
  End If
  Exit Sub

errend:

End Sub

Sub DoNew()
  On Error GoTo errend
  If Not saved(FMainform.SSTab2.Tab) Then
    If MsgBox("Would you like to save " & FMainform.SSTab2.Caption & "?", vbYesNo) = vbYes Then
      If Mid(FMainform.SSTab2.Caption, 1, 6) = "noname" Then
        saveAsfile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      Else
        savefile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      End If
    End If
  End If

  FMainform.CodeMax1(FMainform.SSTab2.Tab).text = ""
  FilePath(FMainform.SSTab2.Tab) = "noname" & FMainform.SSTab2.Tab & ".rcp"
  FMainform.SSTab2.Caption = FilePath(FMainform.SSTab2.Tab)
  saved(FMainform.SSTab2.Tab) = True
  Exit Sub
  
errend:

End Sub

Sub DoOpen()
  On Error GoTo errend

  If Not saved(FMainform.SSTab2.Tab) Then
    If MsgBox("Would you like to save " & FMainform.SSTab2.Caption & "?", vbYesNo) = vbYes Then
      If Mid(FMainform.SSTab2.Caption, 1, 6) = "noname" Then
        saveAsfile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      Else
        savefile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      End If
    End If
  End If
  
  openfile (FMainform.SSTab2.Tab)
  saved(FMainform.SSTab2.Tab) = True
  Exit Sub
  
errend:
  MsgBox "Error: cannot open file."
End Sub

Sub DoSave()
  On Error GoTo errend

  If Not saved(FMainform.SSTab2.Tab) Then
    If Mid(FMainform.SSTab2.Caption, 1, 6) = "noname" Then
      saveAsfile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      If DefineMain <> -1 Then
        FMainform.SSTab2.Caption = FMainform.SSTab2.Caption + " *"
      End If
    Else
      savefile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
    End If
    saved(FMainform.SSTab2.Tab) = True
  End If
  Exit Sub
  
errend:
  MsgBox "Error: cannot save file."
End Sub

Sub DoSaveAs()
  On Error GoTo errend
  saveAsfile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
  
  If DefineMain <> -1 Then
    FMainform.SSTab2.Caption = FMainform.SSTab2.Caption + " *"
  End If
  Exit Sub
  
errend:
  MsgBox "Error: cannot save as file."
End Sub

Sub DoSaveAll()
  Dim i As Integer
  
  On Error GoTo errend
  For i = 0 To 2
    FMainform.SSTab2.Tab = i
    If Not saved(FMainform.SSTab2.Tab) Then
      If Mid(FMainform.SSTab2.Caption, 1, 6) = "noname" Then
        saveAsfile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      Else
        savefile FMainform.SSTab2.Tab, FilePath(FMainform.SSTab2.Tab)
      End If
    End If
    saved(FMainform.SSTab2.Tab) = True
  Next i
  
  Exit Sub
  
errend:

End Sub

Sub DoPrint()

Rem  Dim LineWidth As Long
Rem  FMainform.CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
Rem  FMainform.CommonDialog1.ShowPrinter
Rem  LineWidth = WYSIWYG_RTF(FMainform.codemax1(FMainform.SSTab2.Tab), 1440, 1440) '1440 Twips=1 Inch
Rem  ' Print the contents of the RichTextBox with a one inch margin
Rem  PrintRTF FMainform.codemax1(FMainform.SSTab2.Tab), 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch
Rem  FMainform.CodeMax1(FMainform.SSTab2.Tab).PrintContents

End Sub

Function SearchBackward(car As String, text As String) As Integer
  Dim i As Integer

  For i = Len(text) To 1 Step -1
    If car = Mid(text, i, 1) Then
      SearchBackward = i
      Exit Function
    End If
  Next i
  SearchBackward = 0
End Function

Sub DoCurrentPath()
  Dim i As Integer
  i = SearchBackward("\", FilePath(FMainform.SSTab2.Tab))
  If i > 0 Then
    CurrentPath = Mid(FilePath(FMainform.SSTab2.Tab), 1, i - 1)
  Else
    CurrentPath = App.path
  End If
  FMainform.path = "Working path: " & Chr(13) & Chr(10) & _
  Left(CurrentPath, 3) & "..." & Right(CurrentPath, 17)
  FMainform.path.ToolTipText = CurrentPath
End Sub
