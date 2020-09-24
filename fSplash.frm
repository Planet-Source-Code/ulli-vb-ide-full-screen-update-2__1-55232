VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picMenuDown 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3030
      Picture         =   "fSplash.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMenuUp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2745
      Picture         =   "fSplash.frx":037E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.OptionButton opDummy 
      Height          =   195
      Left            =   4710
      TabIndex        =   1
      Top             =   825
      Width           =   180
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   195
      Picture         =   "fSplash.frx":06C0
      Top             =   188
      Width           =   825
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Full Screen Add-In..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   450
      Width           =   3000
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this form has no code

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 5 + 0 = 5 Lines
