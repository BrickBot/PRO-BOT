VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form preproc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result of the compile process"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
   HelpContextID   =   440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7223
      _Version        =   327681
      TabOrientation  =   3
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
      TabCaption(0)   =   "Pass 1"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Pass1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pass 2"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pass2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Syntax"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Pass3"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Pass3 
         BackColor       =   &H80000008&
         ForeColor       =   &H80000005&
         Height          =   3615
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Pass2 
         BackColor       =   &H80000008&
         ForeColor       =   &H80000005&
         Height          =   3615
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Pass1 
         BackColor       =   &H80000008&
         ForeColor       =   &H80000005&
         Height          =   3615
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "preproc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  preproc.Visible = False
End Sub
