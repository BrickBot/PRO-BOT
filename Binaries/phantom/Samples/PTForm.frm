VERSION 5.00
Object = "{C6114D03-59EB-48D0-96E6-A27A8A65F021}#1.0#0"; "PHANTOM.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PTForm 
   Caption         =   "PhantomTester"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "PTFORM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   6255
      Left            =   6600
      TabIndex        =   35
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   11033
      _Version        =   393216
      Appearance      =   1
      Max             =   11000
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnDownloadProgram 
      Caption         =   "Download Program"
      Height          =   675
      Left            =   3360
      TabIndex        =   34
      Top             =   4455
      Width           =   1455
   End
   Begin VB.ComboBox ComPortList 
      Height          =   315
      ItemData        =   "PTFORM.frx":030A
      Left            =   3120
      List            =   "PTFORM.frx":030C
      TabIndex        =   33
      Text            =   "Combo1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnDownloadFirmware 
      Caption         =   "Download Firmware"
      Height          =   675
      Left            =   3360
      TabIndex        =   32
      Top             =   5220
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sounds"
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   2655
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btnPlaySystemSound 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Motors"
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   1380
      Width           =   4575
      Begin MSComctlLib.Slider PowerLevel 
         Height          =   375
         Left            =   1920
         TabIndex        =   37
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
      End
      Begin VB.ComboBox MotorList 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Text            =   "Motor 0"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnMotorOn 
         Caption         =   "Motor On"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnMotorOff 
         Caption         =   "Motor Off"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   900
         Width           =   1455
      End
      Begin VB.CommandButton btnMotorRev 
         Caption         =   "Motor Rev"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton btnSetPower 
         Caption         =   "Set Power"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Power Level"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   2040
         Width           =   2295
      End
   End
   Begin VB.CommandButton btnGetShortTermRetransStatistics 
      Caption         =   "STRetransStatistics"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton btnGetLongTermRetransmitStatistics 
      Caption         =   "LTRetransStatistics"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton btnTowerAndCableConnected 
      Caption         =   "Tower And Cable"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton btnTowerAlive 
      Caption         =   "Tower Alive"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox ResultWindow 
      Height          =   6255
      Left            =   7140
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   135
      Width           =   1815
   End
   Begin VB.CommandButton btnDisConnect 
      Caption         =   "DisConnect"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton btnDataLog 
      Caption         =   "Data Log"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton btnPBTurnOff 
      Caption         =   "PowerOff"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton btnMemMap 
      Caption         =   "MemMap"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton btnPBAliveOrNot 
      Caption         =   "Brick Alive"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sensors"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   3135
      Begin VB.CheckBox Check1 
         Caption         =   "Sensor 3"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sensor 2"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sensor 1"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnStartMonitorSen 
         Caption         =   "Monitor Sensors"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnStopMonitorSen 
         Caption         =   "Stop Monitor"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
   End
   Begin PHANTOMLibCtl.PhantomCtrl Phantom 
      Left            =   4320
      Top             =   600
      ComPortNo       =   0
      LinkType        =   0
      Brick           =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brick Alive = "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "PTForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConnect_Click()
    Phantom.Brick = RCX2
    Phantom.InitComm
    Label1.Caption = "Brick Alive = " & Phantom.PBAliveOrNot
    ProgressBar2.Value = Phantom.PBBattery
    
End Sub

Private Sub btnDownloadProgram_Click()
  'simple motor on wait off test program
  Phantom.SelectPrgm 3
  
  Phantom.BeginOfTask 0
    Phantom.PlaySystemSound SOUND_SWEEPDOWN
    Phantom.Wait SRC_CON, 50    'wait 0.5 sec
    'drive forward for 2 sec
    Phantom.SetPower "motor0motor2", SRC_CON, 7
    Phantom.SetFwd "motor0motor2"
    Phantom.On "motor0motor2"
    Phantom.Wait SRC_CON, 200    'wait 2 sec
    'change direction and drive 2 sec
    Phantom.SetRwd "motor0motor2"
    Phantom.Wait SRC_CON, 200    'wait 2 sec
    Phantom.Off "motor0motor2"
    'play buildin sound
    Phantom.PlaySystemSound SOUND_SWEEPUP
  Phantom.EndOfTask
  Phantom.BeginOfSub 0
    Phantom.PlaySystemSound SOUND_SWEEPUP
  
  Phantom.EndOfSub

End Sub

Private Sub btnSetPower_Click()
    Phantom.SetPower MotorList.Text, 2, PowerLevel.Value
End Sub

Private Sub btnDataLog_Click()
    Dim i As Integer
    Phantom.SetDatalog 15
    Phantom.DatalogNext 0, 3
    Phantom.DatalogNext 1, 0
    Phantom.DatalogNext 9, 0
    Phantom.DatalogNext 14, 0
    ResultWindow.Text = ""
    Dim arr As Variant
    arr = Phantom.UploadDatalog(0, 15)
    If IsArray(arr) Then
        For i = LBound(arr, 2) To UBound(arr, 2)
        ResultWindow.Text = ResultWindow.Text & "Type: " + Str(arr(0, i)) + " No. " + Str(arr(1, i)) + " Value: " + Str(arr(2, i)) & vbCrLf
        Next i
    Else
        MsgBox "Upload NOT a valid array"
    End If

End Sub

Private Sub btnDisConnect_Click()
    Phantom.CloseComm
End Sub

Private Sub btnDownloadFirmware_Click()
    CommonDialog1.filename = ""
    CommonDialog1.Filter = "Firmware (*.lgo)|*.lgo"
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
        btnDownloadFirmware.Enabled = False
        Phantom.DownloadFirmware CommonDialog1.filename
    End If
End Sub

Private Sub btnTowerAndCableConnected_Click()
    Label1.Caption = "Tower And Cable = " & Phantom.TowerAndCableConnected
End Sub

Private Sub btnTowerAlive_Click()
    Label1.Caption = "Tower Alive = " & Phantom.TowerAlive
End Sub

Private Sub btnGetShortTermRetransStatistics_Click()
    Dim Stat As Variant
    Dim i As Long
    Dim val1 As Integer
    Dim val2 As Integer
    Stat = Phantom.GetShortTermRetransStatistics()
    ResultWindow.Text = ""
    For i = LBound(Stat, 2) To UBound(Stat, 2)
        val1 = Stat(0, i)
        val2 = Stat(1, i)
        ResultWindow.Text = ResultWindow.Text & Str(val1) + " : " + Str(val2) & vbCrLf
    Next i
End Sub

Private Sub btnGetLongTermRetransmitStatistics_Click()
    Dim Stat As Variant
    Dim i As Long
    Dim val1 As Integer
    Dim val2 As Integer
    Stat = Phantom.GetLongTermRetransmitStatistics()
    ResultWindow.Text = ""
    For i = LBound(Stat, 2) To UBound(Stat, 2)
        val1 = Stat(0, i)
        val2 = Stat(1, i)
        ResultWindow.Text = ResultWindow.Text & Str(val1) + " : " + Str(val2) & vbCrLf
    Next i
End Sub

Private Sub btnStartMonitorSen_Click()
    Dim i As Integer
    For i = 0 To 2
        Debug.Print Phantom.SetEvent(SRC_SENBOOL, i, 300)
    Next i
End Sub

Private Sub btnStopMonitorSen_Click()
    Dim i As Integer
    For i = 0 To 2
        Debug.Print Phantom.ClearEvent(SRC_SENBOOL, i)
    Next i
End Sub


Private Sub btnPBAliveOrNot_Click()
    Label1.Caption = "Brick Alive = " & Phantom.PBAliveOrNot
End Sub

Private Sub btnMemMap_Click()
    ShowMap (Phantom.MemMap)
End Sub

Private Sub btnPBTurnOff_Click()
    Debug.Print Phantom.PBTurnOff
End Sub

Private Sub btnPlaySystemSound_Click(Index As Integer)
    Phantom.PlaySystemSound Index
End Sub

Private Sub btnMotorOn_Click()
     Phantom.On MotorList.Text
End Sub

Private Sub btnMotorOff_Click()
     Phantom.Off MotorList.Text
End Sub

Private Sub btnMotorRev_Click()
    Phantom.AlterDir MotorList.Text
End Sub

Private Sub Command1_Click()
  Phantom.AboutBox
End Sub

Private Sub ComPortList_Click()
    Phantom.ComPortNo = ComPortList.ListIndex
End Sub

Private Sub Form_Load()
    MotorList.AddItem "Motor 0"
    MotorList.AddItem "Motor 1"
    MotorList.AddItem "Motor 2"
    ComPortList.AddItem "USB"
    ComPortList.AddItem "COM1"
    ComPortList.AddItem "COM2"
    ComPortList.AddItem "COM3"
    ComPortList.AddItem "COM4"
    ComPortList.ListIndex = Phantom.ComPortNo
End Sub

Private Sub Form_Resize()
    ResultWindow.Move ResultWindow.Left, ResultWindow.Top, Abs(PTForm.ScaleWidth - (ResultWindow.Left + ResultWindow.Top)), Abs(PTForm.ScaleHeight - (ResultWindow.Top + ResultWindow.Top))
End Sub

Private Sub Phantom_AsyncronBrickError(ByVal Number As Long, ByVal Description As String)
   MsgBox Description
End Sub

Private Sub Phantom_DownloadDone(ByVal ErrorCode As Long, ByVal DownloadNo As Long)
    btnDownloadFirmware.Enabled = True

End Sub

Private Sub Phantom_DownloadProgress(ByVal BytesDownloaded As Long)
    ProgressBar1.Value = BytesDownloaded
End Sub

Private Sub Phantom_DownloadStatus(ByVal timeInMS As Long, ByVal sizeInBytes As Long, ByVal taskNo As Long)
    ProgressBar1.Min = 0
    If sizeInBytes <> 0 Then
        ProgressBar1.Max = sizeInBytes
    End If
    ProgressBar1.Value = 0
End Sub


Private Sub Phantom_InputChange(ByVal Number As Long, ByVal Value As Long)
    Check1(Number).Value = Value
    Debug.Print "Phantom_InputChange", Number, Value
End Sub

Private Sub Phantom_PBMessage(ByVal Number As Integer)
    Debug.Print Number
End Sub

Private Sub Phantom_VariableChange(ByVal Number As Long, ByVal Value As Long)
    Debug.Print "Phantom_VariableChange", Number, Value
End Sub

Private Sub ShowMap(Map As Variant)
    Me.Show
    Dim Slot As Integer
    Dim Task As Integer
    Dim Subr As Integer
    Dim offset As Integer
    Dim size As Integer
    Dim outstring As String
    Dim strtlen As Integer
       
    ResultWindow.Text = ""
    For Slot = 0 To 4
        outstring = "Program: " & Str(Slot) & vbCrLf
        strtlen = Len(outstring)
        For Task = 0 To 9
            offset = ((10 * Slot) + Task) + 41
            If Map(offset) <> Map(offset + 1) Then
                outstring = outstring & "  Task: " & Str(Task) & " Start: " & Str(Map(offset)) & " Length: " & Str(Map(offset + 1) - Map(offset)) & vbCrLf
            End If
        Next Task
        
        For Subr = 0 To 7
            offset = ((8 * Slot) + Subr) + 1
            size = Map(offset + 1) - Map(offset)
            If size > 1 Then
                outstring = outstring & "  Subr: " & Str(Subr) & " Start: " & Str(Map(offset)) & " Length: " & Str(size - 1) & vbCrLf
            End If
        Next Subr
        If Len(outstring) <> strtlen Then
            ResultWindow.Text = ResultWindow.Text & outstring
        End If
    Next Slot
    ResultWindow.Text = ResultWindow.Text & "DataLog" & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "LOG Start:  " & vbTab & Str(Map(91)) & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "LOG Current:" & vbTab & Str(Map(92)) & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "LOG End:    " & vbTab & Str(Map(93)) & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "LOG Length: " & vbTab & Str(Map(93) - Map(91)) & " (" & Str((Map(93) - Map(91)) / 3) & " items)" & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "Memory End: " & vbTab & Str(Map(94)) & vbCrLf
    ResultWindow.Text = ResultWindow.Text & "Free RAM:  " & Str(Map(94) - Map(93)) & ", Free Data:  " & Str(Map(93) - Map(92)) & " (" & Str((Map(93) - Map(92)) / 3) & " items), Logged Data:  " & Str(Map(92) - Map(91)) & " (" & Str((Map(92) - Map(91)) / 3) & " items)"
End Sub

