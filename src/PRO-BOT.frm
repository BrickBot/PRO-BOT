VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "CMAX20.OCX"
Object = "{C6114D03-59EB-48D0-96E6-A27A8A65F021}#1.0#0"; "PHANTOM.DLL"
Begin VB.Form FrmMain 
   Caption         =   "PRO-BOT 2000"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9375
   HelpContextID   =   100
   Icon            =   "PRO-BOT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      HelpContextID   =   240
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   11245
      _Version        =   393216
      TabOrientation  =   3
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Command list"
      TabPicture(0)   =   "PRO-BOT.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List1"
      Tab(0).Control(1)=   "thetype(0)"
      Tab(0).Control(2)=   "thetype(1)"
      Tab(0).Control(3)=   "thetype(2)"
      Tab(0).Control(4)=   "lblcategory"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "PRO-BOT.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lego"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CommonDialog1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Check2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Check3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Combo1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Timer1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Check4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Timer2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Check1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "MSComm1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Diagnostic"
      TabPicture(2)   =   "PRO-BOT.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command10"
      Tab(2).Control(1)=   "DiagCheck(3)"
      Tab(2).Control(2)=   "diagno(3)"
      Tab(2).Control(3)=   "LBLDiag(3)"
      Tab(2).Control(4)=   "startstatus"
      Tab(2).Control(5)=   "path"
      Tab(2).Control(6)=   "StartOn"
      Tab(2).Control(7)=   "TimeoutOn"
      Tab(2).Control(8)=   "DiagCheck(2)"
      Tab(2).Control(9)=   "DiagCheck(1)"
      Tab(2).Control(10)=   "DiagCheck(0)"
      Tab(2).Control(11)=   "Line1(0)"
      Tab(2).Control(12)=   "EventOn"
      Tab(2).Control(13)=   "LBLDiag(2)"
      Tab(2).Control(14)=   "LBLDiag(1)"
      Tab(2).Control(15)=   "LBLDiag(0)"
      Tab(2).Control(16)=   "Line1(1)"
      Tab(2).Control(17)=   "Label4(0)"
      Tab(2).Control(18)=   "Label4(2)"
      Tab(2).Control(19)=   "diagno(0)"
      Tab(2).Control(20)=   "diagno(2)"
      Tab(2).Control(21)=   "diagno(1)"
      Tab(2).ControlCount=   22
      Begin MSCommLib.MSComm MSComm1 
         Left            =   900
         Top             =   4020
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327680
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Check again"
         Height          =   375
         HelpContextID   =   250
         Left            =   -74520
         TabIndex        =   18
         ToolTipText     =   "Perform diagnostic to see if the tower emiter is connected and functionning, and if the Brick is on"
         Top             =   1500
         WhatsThisHelpID =   210
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   4935
         HelpContextID   =   260
         Left            =   -74880
         TabIndex        =   1
         ToolTipText     =   "List of all the commands available. ENTER paste one command into your editor zone; F2 toggle between your editor and this list"
         Top             =   120
         WhatsThisHelpID =   220
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show compiling"
         Height          =   255
         HelpContextID   =   240
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "If check, will show the result of the syntax check in a second window.  Errors will be indicated where they occured"
         Top             =   840
         WhatsThisHelpID =   230
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Left            =   960
         Top             =   4680
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Insert RCXDefine.rcp"
         Height          =   255
         HelpContextID   =   240
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "If check, will automatically includes all the defintions included in the file 'RCXDefine.rcp'"
         Top             =   1560
         WhatsThisHelpID =   250
         Width           =   1935
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Fon&ts/Colors"
         Height          =   375
         HelpContextID   =   240
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "Set the font used in the editor zones"
         Top             =   2040
         WhatsThisHelpID =   260
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   4680
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         HelpContextID   =   240
         ItemData        =   "PRO-BOT.frx":035E
         Left            =   720
         List            =   "PRO-BOT.frx":0371
         TabIndex        =   3
         ToolTipText     =   "Define in which port the InfraRed tower emitter is connected"
         Top             =   480
         WhatsThisHelpID =   270
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Autosave"
         Height          =   255
         HelpContextID   =   240
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "If check, will automatically save all the editor zones that have been modified"
         Top             =   1320
         WhatsThisHelpID =   280
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show help (shift-F1)"
         Height          =   255
         HelpContextID   =   240
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "If check, show a small help zone text under the editor zone"
         Top             =   1080
         WhatsThisHelpID =   290
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1680
         Top             =   4680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontName        =   "Times New Roman"
      End
      Begin VB.CheckBox thetype 
         Caption         =   "Check5"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   16
         Top             =   5520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox thetype 
         Caption         =   "Check5"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox thetype 
         Caption         =   "Check5"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   9
         Top             =   6000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin PHANTOMLibCtl.PhantomCtrl Lego 
         Left            =   1680
         Top             =   3960
         ComPortNo       =   0
         LinkType        =   0
         Brick           =   0
      End
      Begin VB.Image DiagCheck 
         Height          =   225
         Index           =   3
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":0392
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image diagno 
         Height          =   225
         Index           =   3
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":069C
         Stretch         =   -1  'True
         ToolTipText     =   "Check that the display shows four digits, or else use DownloadFirmware."
         Top             =   1200
         Visible         =   0   'False
         WhatsThisHelpID =   300
         Width           =   225
      End
      Begin VB.Label LBLDiag 
         Caption         =   "Firmware present?"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label startstatus 
         Caption         =   "PRO-RCX muted..."
         Height          =   195
         Left            =   -74100
         TabIndex        =   27
         Top             =   4740
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label path 
         Height          =   555
         Left            =   -74760
         TabIndex        =   26
         Top             =   2460
         Width           =   1995
      End
      Begin VB.Label StartOn 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Start on"
         Height          =   615
         Left            =   -74760
         TabIndex        =   25
         ToolTipText     =   "Waiting for file 'result.log' to be created."
         Top             =   4560
         Visible         =   0   'False
         WhatsThisHelpID =   310
         Width           =   615
      End
      Begin VB.Label TimeoutOn 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Timout on"
         Height          =   615
         Left            =   -74760
         TabIndex        =   24
         ToolTipText     =   "A time-out is on (dbl click to cancel)"
         Top             =   3840
         Visible         =   0   'False
         WhatsThisHelpID =   320
         Width           =   615
      End
      Begin VB.Image DiagCheck 
         Height          =   225
         Index           =   2
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":09A6
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image DiagCheck 
         Height          =   225
         Index           =   1
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":0CB0
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image DiagCheck 
         Height          =   225
         Index           =   0
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":0FBA
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -74760
         X2              =   -72960
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label EventOn 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Event on"
         Height          =   615
         Left            =   -74760
         TabIndex        =   22
         ToolTipText     =   "An event is on (dbl click to cancel)"
         Top             =   3120
         Visible         =   0   'False
         WhatsThisHelpID =   330
         Width           =   615
      End
      Begin VB.Label LBLDiag 
         Caption         =   "Tower connected?"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label LBLDiag 
         Caption         =   "Tower alive?"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LBLDiag 
         Caption         =   "PB alive?"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   240
         X2              =   2040
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Options"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -74760
         X2              =   -72960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Diagnostic"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblcategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Status"
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   23
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Image diagno 
         Height          =   225
         Index           =   0
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":12C4
         Stretch         =   -1  'True
         ToolTipText     =   "Verify that the tower is firmly connected and also that you set the right com port."
         Top             =   480
         Visible         =   0   'False
         WhatsThisHelpID =   340
         Width           =   225
      End
      Begin VB.Image diagno 
         Height          =   225
         Index           =   2
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":15CE
         Stretch         =   -1  'True
         ToolTipText     =   "Verify that the RCX is in range of the IR and power is on."
         Top             =   960
         Visible         =   0   'False
         WhatsThisHelpID =   350
         Width           =   225
      End
      Begin VB.Image diagno 
         Height          =   225
         Index           =   1
         Left            =   -73200
         Picture         =   "PRO-BOT.frx":18D8
         Stretch         =   -1  'True
         ToolTipText     =   "Verify that the battery in the IR tower is fully charged and installed correctly."
         Top             =   720
         Visible         =   0   'False
         WhatsThisHelpID =   360
         Width           =   225
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6375
      HelpContextID   =   270
      Left            =   2820
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Editor zones where up to three programs can be worked"
      Top             =   60
      WhatsThisHelpID =   370
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11245
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "noname0.rcp"
      TabPicture(0)   =   "PRO-BOT.frx":1BE2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CodeMax1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "noname1.rcp"
      TabPicture(1)   =   "PRO-BOT.frx":1BFE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CodeMax1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "noname2.rcp"
      TabPicture(2)   =   "PRO-BOT.frx":1C1A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CodeMax1(2)"
      Tab(2).ControlCount=   1
      Begin CodeMaxCtl.CodeMax CodeMax1 
         Height          =   5775
         HelpContextID   =   100
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "PRO-BOT.frx":1C36
         TabIndex        =   30
         Top             =   480
         Width           =   6255
      End
      Begin CodeMaxCtl.CodeMax CodeMax1 
         Height          =   5775
         HelpContextID   =   100
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "PRO-BOT.frx":1D98
         TabIndex        =   31
         Top             =   480
         Width           =   6255
      End
      Begin CodeMaxCtl.CodeMax CodeMax1 
         Height          =   5775
         HelpContextID   =   100
         Index           =   2
         Left            =   -74880
         OleObjectBlob   =   "PRO-BOT.frx":1EFA
         TabIndex        =   29
         Top             =   480
         Width           =   6255
      End
   End
   Begin VB.Label lblsyntax 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   3720
      Width           =   6375
   End
   Begin VB.Label lblhelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   13
      Top             =   4080
      Width           =   6375
   End
   Begin VB.Menu menupopup 
      Caption         =   "menupopup"
      Visible         =   0   'False
      Begin VB.Menu catego 
         Caption         =   "catego1"
         Index           =   1
         Begin VB.Menu empty1 
            Caption         =   "empty1"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego2"
         Index           =   2
         Begin VB.Menu empty2 
            Caption         =   "empty2"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego3"
         Index           =   3
         Begin VB.Menu empty3 
            Caption         =   "empty3"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego4"
         Index           =   4
         Begin VB.Menu empty4 
            Caption         =   "empty4"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego5"
         Index           =   5
         Begin VB.Menu empty5 
            Caption         =   "empty5"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego6"
         Index           =   6
         Begin VB.Menu empty6 
            Caption         =   "empty6"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego7"
         Index           =   7
         Begin VB.Menu empty7 
            Caption         =   "empty7"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego8"
         Index           =   8
         Begin VB.Menu empty8 
            Caption         =   "empty8"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego9"
         Index           =   9
         Begin VB.Menu empty9 
            Caption         =   "empty9"
            Index           =   0
         End
      End
      Begin VB.Menu catego 
         Caption         =   "catego10"
         Index           =   10
         Begin VB.Menu empty10 
            Caption         =   "empty10"
            Index           =   0
         End
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      HelpContextID   =   380
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu separator1 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Saveas 
         Caption         =   "Sav&e as"
      End
      Begin VB.Menu Saveall 
         Caption         =   "Save a&ll"
         Shortcut        =   ^L
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu noname 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      HelpContextID   =   390
      Begin VB.Menu Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu separateur42 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu separateur345 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "Select &all"
         Shortcut        =   ^A
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu find 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu findnext 
         Caption         =   "Find next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu find_replace 
         Caption         =   "Find and &Replace"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu runmenu 
      Caption         =   "&Run"
      HelpContextID   =   400
      Begin VB.Menu check 
         Caption         =   "&Check syntax"
         Shortcut        =   {F4}
      End
      Begin VB.Menu run 
         Caption         =   "&Execute"
         Shortcut        =   {F5}
      End
      Begin VB.Menu adsfasf 
         Caption         =   "-"
      End
      Begin VB.Menu define_main 
         Caption         =   "&Define main file"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "&?"
      HelpContextID   =   410
      Begin VB.Menu help 
         Caption         =   "&Help"
         HelpContextID   =   100
         Shortcut        =   {F1}
      End
      Begin VB.Menu diclaim 
         Caption         =   "&Disclaimer"
         Enabled         =   0   'False
         HelpContextID   =   190
      End
      Begin VB.Menu separator3 
         Caption         =   "-"
      End
      Begin VB.Menu aboutpro 
         Caption         =   "About &PRO-BOT 2000"
      End
      Begin VB.Menu aboutspirit 
         Caption         =   "About &Phantom.dll"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem ========================================
Rem ========================================
Rem Event-managing buttons
Rem ========================================
Rem ========================================

Private Sub StartOn_DblClick()
  If MsgBox("Do you want to interrupt the Start (with value -1)?", vbYesNo, App.Title) = vbYes Then
    Open CurrentPath + "\result.log" For Output As #1
    Print #1, "-1"
    Close #1
  End If
End Sub

Private Sub TimeoutOn_DblClick()
  If MsgBox("Do you want to clear the time-out running every " & Timer1.Interval & "ms?", vbYesNo, App.Title) = vbYes Then
    Timer1.Interval = 0
    TimeoutOn.Visible = False
    ATimeOut = False
  End If
End Sub

Private Sub EventOn_DblClick()
  Dim j As Variant
  If MsgBox("Do you want to clear the event on Var" & AnEvent1 & " running every " & AnEvent2 * 10 & "ms?", vbYesNo, App.Title) = vbYes Then
    j = Lego.ClearEvent(AnEvent1, AnEvent2)
    EventOn.Visible = False
    AnEvent = False
  End If
End Sub

Rem ========================================
Rem ========================================
Rem menu commandes
Rem ========================================
Rem ========================================

Rem ========================================
Rem FILE menu commandes
Rem ========================================
Private Sub New_Click()
  Rem new
  DoNew
  CodeMax1(SSTab2.Tab).SetFocus

End Sub

Private Sub open_Click()
  Rem open
  DoOpen
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Private Sub Save_Click()
  Rem save
  DoSave
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Private Sub Saveall_Click()
  Rem save all file
  Dim currenttab As Integer
  currenttab = SSTab2.Tab
  DoSaveAll
  SSTab2.Tab = currenttab
  CodeMax1(SSTab2.Tab).SetFocus

End Sub

Private Sub Saveas_Click()
  Rem save as
  DoSaveAs
  CodeMax1(SSTab2.Tab).SetFocus

End Sub

Private Sub print_Click()
  Rem print
Rem  DoPrint
Rem  CodeMax1(SSTab2.Tab).PrintContents(Printer.hdc, 0) = 0
MsgBox ("Verifier que print fonctionne")
  CodeMax1(SSTab2.Tab).SetFocus

End Sub

Private Sub exit_Click()
  Unload Me
End Sub

Rem ========================================
Rem EDIT menu commandes
Rem ========================================

Private Sub Undo_Click()
  If CodeMax1(SSTab2.Tab).CanUndo Then
    CodeMax1(SSTab2.Tab).Undo
  End If
End Sub

Private Sub copy_Click()
  If CodeMax1(SSTab2.Tab).CanCopy Then
    CodeMax1(SSTab2.Tab).copy
  End If
End Sub

Private Sub cut_Click()
  If CodeMax1(SSTab2.Tab).CanCut Then
    CodeMax1(SSTab2.Tab).cut
  End If
End Sub

Private Sub paste_Click()
  If CodeMax1(SSTab2.Tab).CanPaste Then
    CodeMax1(SSTab2.Tab).paste
  End If
End Sub

Private Sub selectall_Click()
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdSelectAll
End Sub

Private Sub find_Click()
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdFind
  CodeMax1(SSTab2.Tab).SetFocus
  findnext.Enabled = True
End Sub

Private Sub find_replace_Click()
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdFindReplace
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Private Sub findnext_Click()
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdFindNext
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Rem ========================================
Rem RUN menu commandes
Rem ========================================

Private Sub check_Click()
  Rem syntax
  Dim pass4result As String
    
  Rem autosave here
  If ConfigO(3) Then
    Saveall_Click
  End If
  
  Rem show debug?
  If Check1 Then
    preproc.Visible = True
    preproc.Refresh
  End If
  
  Rem compute the current path
  DoCurrentPath
  
  Refresh
  If DefineMain = -1 Then
    pass4result = pass0to3(CodeMax1(SSTab2.Tab).text)
  Else
    pass4result = pass0to3(CodeMax1(DefineMain).text)
  End If
End Sub

Private Sub run_Click()
  Rem compiler
  Dim pass4result As String
  Dim PassSectionResult As String
  Dim j As Integer
  
  Rem autosave here
  If ConfigO(3) Then
    Saveall_Click
  End If
  
  Rem show debug?
  If Check1 Then
    preproc.Visible = True
    preproc.Refresh
  End If
  
  Rem compute the current path
  DoCurrentPath
  
  Refresh
  If DefineMain = -1 Then
    pass4result = pass0to3(CodeMax1(SSTab2.Tab).text)
  Else
    pass4result = pass0to3(CodeMax1(DefineMain).text)
  End If
  
  Rem if error, stop
  If HasError Then
    Exit Sub
  End If
  
  Rem if not alive, stop
  Command10_Click
  If DiagCheck(2).Visible <> True Then
    Warning 50, 0, ""
    Exit Sub
  End If
  
  Rem pass 0- manage sections
  PassSectionResult = PassSection(pass4result, "main")
  
  Rem verify that no event is on
  If ATimeOut Then
    Warning 71, 0, ""
    Timer1.Interval = 0
    TimeoutOn.Visible = False
    ATimeOut = False
  End If
  
  If AnEvent Then
    Warning 70, 0, ""
    j = Lego.ClearEvent(AnEvent1, AnEvent2)
    EventOn.Visible = False
    AnEvent = False
  End If
    
  Rem pass 5- interpret code
  lblsyntax = "Download information:"
  lblhelp = ""
  interpret PassSectionResult

End Sub

Private Sub define_main_Click()
  If DefineMain = -1 Then
    SSTab2.Caption = SSTab2.Caption + " *"
    DefineMain = SSTab2.Tab
  Else
    SSTab2.Tab = DefineMain
    SSTab2.Caption = Mid(SSTab2.Caption, 1, Len(SSTab2.Caption) - 2)
    DefineMain = -1
  End If
End Sub

Rem ========================================
Rem HELP menu commandes
Rem ========================================

Private Sub help_Click()
  MsgBox "Use F1 key."
End Sub

Private Sub aboutpro_Click()
  frmSplash.Visible = True
  frmSplash.bok.Visible = True
  frmSplash.lblLicenseTo = "Comments are welcome"
End Sub

Private Sub aboutspirit_Click()
  Lego.AboutBox
End Sub

Rem ========================================
Rem ========================================
Rem OPTIONS commandes
Rem ========================================
Rem ========================================

Rem four commands to record
Rem the new options when
Rem the objects are changed

Private Sub Combo1_LostFocus()
  ConfigO(1) = Combo1.ListIndex
  SaveOptions
  Lego.ComPortNo = Combo1.ListIndex
  
Debug.Print Lego.ComPortNo

End Sub

Private Sub Check1_Click()
  ConfigO(2) = Check1.Value
  SaveOptions
End Sub

Private Sub Check2_Click()
  Dim i As Integer
  
  ConfigO(3) = Check2.Value
  SaveOptions
  If ConfigO(3) Then
    For i = 0 To 2
      CodeMax1(i).Height = lblsyntax.Top - CodeMax1(i).Top - 200
    Next i
    SSTab2.Height = lblsyntax.Top - 150
    List1.Height = lblcategory.Top - 100
  Else
    For i = 0 To 2
      CodeMax1(i).Height = lblhelp.Top + lblhelp.Height - CodeMax1(i).Top - 60
    Next i
    SSTab2.Height = lblhelp.Top + lblhelp.Height
    List1.Height = thetype(2).Top + thetype(2).Height
  End If
End Sub

Private Sub Check3_Click()
  ConfigO(4) = Check3.Value
  SaveOptions
End Sub

Private Sub Check4_Click()
  ConfigO(5) = Check4.Value
  SaveOptions
End Sub

Rem ========================================
Rem ========================================
Rem OPTIONS button commandes
Rem ========================================
Rem ========================================

Private Sub Command11_Click()
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdProperties
  CodeMax1(SSTab2.Tab).SetFocus
  ConfigF(1) = CodeMax1(SSTab2.Tab).Font.Bold
  ConfigF(2) = CodeMax1(SSTab2.Tab).Font.Italic
  ConfigF(3) = CodeMax1(SSTab2.Tab).Font.Name
  ConfigF(4) = CodeMax1(SSTab2.Tab).Font.Size
  ConfigF(5) = CodeMax1(SSTab2.Tab).Font.Strikethrough
  ConfigF(6) = CodeMax1(SSTab2.Tab).Font.Underline
  SaveCodemax
End Sub

Rem ========================================
Rem ========================================
Rem buttons commandes
Rem ========================================
Rem ========================================

Private Sub Command10_Click()   'diagnostic button
  Rem perform diagnostic
  Command10.Enabled = False
  DoDiagnostic
  Command10.Enabled = True
End Sub


Private Sub Form_Resize()
  Dim i As Integer

  If ScaleHeight > SSTab2.Top + lblsyntax.Height + lblhelp.Height + 1000 Then
    lblsyntax.Top = ScaleHeight - lblsyntax.Height - lblhelp.Height - 100
    lblhelp.Top = ScaleHeight - lblhelp.Height - 100
    Check2_Click
  End If
  
  If ScaleWidth > SSTab2.Left + 300 Then
    lblsyntax.Width = ScaleWidth - lblsyntax.Left - 100
    lblhelp.Width = ScaleWidth - lblhelp.Left - 100
    SSTab2.Width = ScaleWidth - SSTab2.Left - 100
    For i = 0 To 2
      CodeMax1(i).Width = ScaleWidth - SSTab2.Left - 240
    Next i
  End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Rem shortkeys that moves across objects
  If KeyCode = vbKeyF1 And Shift = 1 Then
    If Check2.Value = 1 Then
      Check2.Value = 0
    Else
      Check2.Value = 1
    End If
    KeyCode = 0
  ElseIf KeyCode = vbKeyF2 Then
    If EditorHasFocus Then
      SSTab1.Tab = 0
      List1.SetFocus
    Else
      CodeMax1(SSTab2.Tab).SetFocus
    End If
    EditorHasFocus = Not EditorHasFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyF6 Then
    If EditorHasFocus Then
      If Shift = 0 Then
        SSTab2.Tab = (SSTab2.Tab + 1) Mod 3
      ElseIf Shift = 1 Then
        SSTab2.Tab = (SSTab2.Tab - 1 + 3) Mod 3
      End If
    Else
      If Shift = 0 Then
        SSTab1.Tab = (SSTab1.Tab + 1) Mod 3
      ElseIf Shift = 1 Then
        SSTab1.Tab = (SSTab1.Tab - 1 + 3) Mod 3
      End If
    End If
    KeyCode = 0
  End If
End Sub

Private Sub form_unload(Cancel As Integer)
  Rem verify that no event is on
  Dim j As Integer
  
  If AStart Then
    Warning 72, 0, ""
    Open CurrentPath + "\result.log" For Output As #1
    Print #1, "-1 Error: User interrupted."
    Close #1
  End If
  
  If ATimeOut Then
    Warning 71, 0, ""
    Timer1.Interval = 0
    TimeoutOn.Visible = False
    ATimeOut = False
  End If
  
  If AnEvent Then
    Warning 70, 0, ""
    j = Lego.ClearEvent(AnEvent1, AnEvent2)
    EventOn.Visible = False
    AnEvent = False
  End If
  
  ConfigP(1) = FMainform.WindowState
  ConfigP(2) = FMainform.Top
  ConfigP(3) = FMainform.Left
  ConfigP(4) = FMainform.Height
  ConfigP(5) = FMainform.Width
  SavePosition
  
  Rem closing comm port
  Lego.CloseComm
  
  Saveall_Click
  Unload preproc
  Unload progress
End Sub

Private Sub Lego_AsyncronBrickError(ByVal Number As Long, ByVal Description As String)
  If OptionShowErrors Then
    Warning 90, (Number), Description
  End If
End Sub

Private Sub Lego_DownloadStatus(ByVal timeInMS As Long, ByVal sizeInBytes As Long, ByVal taskNo As Long)
  Dim thetype As String
  
  If taskNo > 20 Then
    thetype = "Firmware"
  ElseIf taskNo < 10 Then
    thetype = "Task"
  Else
    thetype = "Sub"
  End If
  lblhelp = lblhelp & "Downloading " & thetype & "(" & taskNo & ") will complete in " & timeInMS & " ms; size is " & sizeInBytes & " bytes" & Chr(13) & Chr(10)
End Sub

Private Sub Lego_VariableChange(ByVal Number As Long, ByVal Value As Long)
  Dim PassSectionResult As String
  Dim pass5result As String
  
  Rem if not alive, stop
  If DiagCheck(2).Visible <> True Then
    Warning 50, 0, ""
    Exit Sub
  End If
  
  Rem pass 0- manage sections
  PassSectionResult = PassSection(preproc.Pass3, "onevent")
  
  Rem pass 5- interpret code
  interpret PassSectionResult
  
End Sub

Rem ========================================
Rem ========================================
Rem LIST F2 commandes
Rem ========================================
Rem ========================================
Rem the following are for the list of all the
Rem command.  The list is obtained from tokenlist
Rem if one click, set help; if doubleclick, paste

Private Sub List1_Click()
  Rem place syntax and help in help labels.
  puttokenhelp InList(List1.text, tokenlist(), nbrtoken)
End Sub

Private Sub List1_DblClick()
  CodeMax1(SSTab2.Tab).ReplaceSel List1.text
  CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    List1_DblClick
  End If
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
  CodeMax1(SSTab2.Tab).SetFocus
End Sub

Rem ========================================
Rem ========================================
Rem codemax1 commandes
Rem ========================================
Rem ========================================

Private Sub CodeMax1_PropsChange(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax)
  SaveCodemax
End Sub

Private Sub CodeMax1_GotFocus(Index As Integer)
  EditorHasFocus = True
End Sub

Private Sub CodeMax1_RegisteredCmd(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax, ByVal lCmd As CodeMaxCtl.cmCommand)
  If (lCmd = 1000) Then
    New_Click
  ElseIf (lCmd = 1001) Then
    open_Click
  ElseIf (lCmd = 1002) Then
    Save_Click
  ElseIf (lCmd = 1003) Then
    Saveall_Click
  ElseIf (lCmd = 1004) Then
    print_Click
  ElseIf (lCmd = 1010) Then
    selectall_Click
  ElseIf (lCmd = 1011) Then
    find_Click
  ElseIf (lCmd = 1012) Then
    find_replace_Click
  End If
End Sub

Private Sub CodeMax1_LostFocus(Index As Integer)
  EditorHasFocus = False
End Sub

Private Function CodeMax1_KeyPress(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
  Dim signe As Integer

  If saved(Index) Then
    saved(Index) = False
  End If
  With CodeMax1(Index)
    If IsSeparator(Chr(KeyAscii)) Then
      signe = InList(.CurrentWord, tokenlist(), nbrtoken)
      If signe <> 0 Then
        puttokenhelp signe
      End If
    End If
  End With
End Function

Private Function CodeMax1_MouseDown(Index As Integer, ByVal Control As CodeMaxCtl.ICodeMax, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
  If PopupBoolean And Button = 2 Then
    PopupMenu menupopup
  End If
End Function

Rem ========================================
Rem ========================================
Rem TIMER commandes
Rem ========================================
Rem ========================================

Private Sub Timer1_Timer()
  Dim PassSectionResult As String
  Dim pass5result As String

  Rem if not alive, stop
  If DiagCheck(2).Visible <> True Then
    Warning 50, 0, ""
    Exit Sub
  End If
  
  Rem pass 0- manage sections
  PassSectionResult = PassSection(preproc.Pass3, "ontimeout")
  
  Rem pass 5- interpret code
  interpret PassSectionResult

End Sub

Private Sub Timer2_Timer()
  Timer2.Interval = 0
  PauseElapsed = True
End Sub

Rem ========================================
Rem ========================================
Rem POPUP commandes
Rem ========================================
Rem ========================================

Private Sub empty1_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty1(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty2_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty2(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty3_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty3(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty4_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty4(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty5_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty5(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty6_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty6(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty7_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty7(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty8_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty8(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty9_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty9(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub

Private Sub empty10_Click(Index As Integer)
    CodeMax1(SSTab2.Tab).SelText = empty10(Index).Caption
    CodeMax1(SSTab2.Tab).ExecuteCmd cmCmdWordRight
End Sub
