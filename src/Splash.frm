VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6630
   ControlBox      =   0   'False
   HelpContextID   =   430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton bok 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Portions Copyright 1997-2000 Barry Allyn.  All rights reserved."
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2580
      Width           =   4335
   End
   Begin VB.Label lblcompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblcompany"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   6
      Tag             =   "Version"
      Top             =   2220
      Width           =   3630
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "See http://Prelude.Psy.UMontreal.ca/~cousined/lego"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Tag             =   "Warning"
      Top             =   2940
      Width           =   6255
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   1260
      Width           =   3765
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CompanyProduct"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Tag             =   "CompanyProduct"
      Top             =   240
      Width           =   2865
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2160
      TabIndex        =   1
      Tag             =   "Product"
      Top             =   660
      Width           =   2865
   End
   Begin VB.Label lblLicenseTo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Tag             =   "LicenseTo"
      Top             =   1740
      Width           =   3885
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   240
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bok_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblCompanyProduct = App.FileDescription
    lblcompany = App.LegalCopyright
    lblProductName.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor
End Sub

