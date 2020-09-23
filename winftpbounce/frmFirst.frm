VERSION 5.00
Begin VB.Form frmFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to wBounce"
   ClientHeight    =   3285
   ClientLeft      =   3300
   ClientTop       =   3210
   ClientWidth     =   6510
   Icon            =   "frmFirst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6510
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   120
      Width           =   975
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         Picture         =   "frmFirst.frx":000C
         ScaleHeight     =   735
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   35
         Left            =   840
         ScaleHeight     =   30
         ScaleWidth      =   5415
         TabIndex        =   5
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label LabelTip 
         BackColor       =   &H80000009&
         Height          =   1815
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "First Time..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   80
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    
    On Error GoTo PROC_ERR
    
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "First" & vbCrLf & "Procedure: " & "Command2_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    LabelTip.Caption = "This is the first time you have run Win FTP Bounce on this system. Please check the settings window, as all the settings have been set to their default values."
End Sub

