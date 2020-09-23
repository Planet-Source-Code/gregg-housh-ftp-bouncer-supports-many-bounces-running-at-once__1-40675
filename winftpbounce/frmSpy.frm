VERSION 5.00
Begin VB.Form frmSpy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spy"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   2655
      Left            =   10
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   630
      Width           =   7425
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSpy.frx":0000
      Top             =   70
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connection Spy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -585
      TabIndex        =   0
      Top             =   300
      Width           =   8025
   End
   Begin VB.Label lblTop1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0051B6F2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FTP Bouncer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   -600
      TabIndex        =   1
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mbDone As Boolean
Private miConnection As Integer

Public Sub AddToLog(ByVal strData As String)
    On Error GoTo PROC_ERR

    txtLog.Text = txtLog.Text & strData
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmSpy" & vbCrLf & "Procedure: " & "RefreshInformation" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub RefreshInformation()
    On Error GoTo PROC_ERR
    
    Dim Client As cClients
    Set Client = gOClients(CStr(miConnection))
    
    txtLog.Text = Client.Log
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmSpy" & vbCrLf & "Procedure: " & "RefreshInformation" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub Display(ByVal iConnection As Integer, ByVal sCaption As String, ByVal Parent As Form)
    On Error GoTo PROC_ERR
    
    miConnection = iConnection
    
    Me.Caption = sCaption
    
    RefreshInformation
    
    Me.Tag = miConnection
    
    Me.Show vbModeless, Parent
    
    Do Until mbDone = True
        DoEvents
    Loop

    mbDone = False

    Unload Me

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmSpy" & vbCrLf & "Procedure: " & "Display" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdClose_Click()
    On Error GoTo PROC_ERR
    
    mbDone = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmSpy" & vbCrLf & "Procedure: " & "cmdClose_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub txtLog_Change()
    On Error GoTo PROC_ERR
    
    txtLog.SelStart = Len(txtLog.Text)

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmSpy" & vbCrLf & "Procedure: " & "txtLog_Change" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub
