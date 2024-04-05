VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form progress 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   HelpContextID   =   450
   LinkTopic       =   "Form2"
   ScaleHeight     =   1770
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   1
      End
      Begin VB.Label status 
         Caption         =   "Pass 3: Syntax check"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   5
         Top             =   780
         Width           =   2835
      End
      Begin VB.Image check 
         Height          =   300
         Index           =   2
         Left            =   120
         Picture         =   "progress.frx":0000
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image check 
         Height          =   300
         Index           =   1
         Left            =   120
         Picture         =   "progress.frx":030A
         Stretch         =   -1  'True
         Top             =   420
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image check 
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "progress.frx":0614
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label status 
         Caption         =   "Pass 1: Inserting files/defines..."
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   4
         Top             =   180
         Width           =   2835
      End
      Begin VB.Label status 
         Caption         =   "Pass 2: Managing declarations..."
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   3
         Top             =   480
         Width           =   2835
      End
      Begin VB.Label Label3 
         Caption         =   "Pass 1: Inserting files..."
         Height          =   255
         Left            =   660
         TabIndex        =   2
         Top             =   480
         Width           =   2115
      End
   End
End
Attribute VB_Name = "progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


