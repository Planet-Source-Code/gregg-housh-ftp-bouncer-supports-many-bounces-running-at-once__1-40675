VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version: 1.0 beta"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "gregg@sc.am"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblHomepage 
      Caption         =   "http://www.cryptim.com/bounce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblLicense 
      Caption         =   "License: Free"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author: Gregg Housh"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image Server 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":058A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTop2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0063C7ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4005
   End
   Begin VB.Label lblTop 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WinFTPBounce"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Display()
    Me.Show
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
    OpenURL "mailto:gregg@sc.am", Me.hwnd
End Sub

Private Sub lblHomepage_Click()
    OpenURL "http://www.cryptim.com/bounce", Me.hwnd
End Sub
