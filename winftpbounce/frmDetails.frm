VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Details"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpy 
      Caption         =   "&Spy"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdDisconnectRemove 
      Caption         =   "Disconnect && &Remove IP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblDLB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3960
      TabIndex        =   25
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Bytes DL:"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Uploads:"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblUL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Downloads:"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblDL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblToPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblToHost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblFromIP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "CWD:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblCWD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "User:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblTimeStamp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "TimeStamp:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Site:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblSite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDetails.frx":0000
      Top             =   70
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connection Details"
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
      Width           =   5625
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
      Width           =   5655
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mbDone As Boolean
Private miConnection As Integer

Public Sub RefreshInformation()
    On Error GoTo PROC_ERR
    
    Dim Client As cClients
    Set Client = gOClients(CStr(miConnection))
    With Client
        Select Case frmMain.wsServer(Client.Index).State
            Case 0
                .Status = "Closed"
            Case 1
                .Status = "Open"
            Case 2
                .Status = "Listening"
            Case 3
                .Status = "Connection Pending"
            Case 4
                .Status = "Resolving Host"
            Case 5
                .Status = "Host resolved"
            Case 6
                .Status = "Connecting"
            Case 7
                .Status = "Connected"
            Case 8
                .Status = "Peer is closing connection"
            Case 9
                .Status = "Error"
        End Select
        lblStatus.Caption = .Status
        lblFromIP.Caption = .ClientIP
        lblToHost.Caption = .ServerIP
        lblToPort.Caption = .ServerPort
        'txtServer.Text = .ServerCommands
        'txtClient.Text = .ClientCommands
        'txtLog.Text = .Log
        lblUL.Caption = .Uploads
        lblDL.Caption = .Downloads
        lblDLB.Caption = .Bytes
        lblUser.Caption = .User
        lblCWD.Caption = .CWD
        lblTimeStamp.Caption = .TimeStamp
        lblSite.Caption = .Site
        gOClients.Remove CStr(miConnection)
        gOClients.Add Client, CStr(miConnection)
    End With

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmDetails" & vbCrLf & "Procedure: " & "RefreshInformation" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub Display(ByVal iConnection As Integer, ByVal Parent As Form)
    On Error GoTo PROC_ERR

    miConnection = iConnection
    
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
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmDetails" & vbCrLf & "Procedure: " & "Display" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdClose_Click()
    On Error GoTo PROC_ERR
    
    mbDone = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmDetails" & vbCrLf & "Procedure: " & "cmdClose_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdDisconnect_Click()
    On Error GoTo PROC_ERR
    
    Dim ofrm As Form
    
    frmMain.WriteToLog "Server", "Kicked:" & frmMain.wsClient(miConnection).RemoteHostIP
    frmMain.wsServer(miConnection).Close
    frmMain.wsClient(miConnection).Close
    frmMain.CreateOld miConnection
    Unload frmMain.wsClient(miConnection)
    Unload frmMain.wsServer(miConnection)
    
    For Each ofrm In Forms
        If ofrm.Name = "frmSpy" And ofrm.Tag = CStr(miConnection) Then
            ofrm.mbDone = True
            Unload ofrm
        End If
    Next ofrm
    
    mbDone = True
    
PROC_EXIT:
    Set ofrm = Nothing
        
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmDetails" & vbCrLf & "Procedure: " & "cmdDisconnect_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdSpy_Click()
    On Error GoTo PROC_ERR
    
    Dim Client As cClients
    Dim oSpy As frmSpy
    Set oSpy = New frmSpy
    Set Client = gOClients(CStr(miConnection))

    oSpy.Display miConnection, "Spying on: " & Client.ClientIP & " connected to: " & Client.Site, Me

PROC_EXIT:
    Unload oSpy
    Set Client = Nothing
    Set oSpy = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmDetails" & vbCrLf & "Procedure: " & "cmdSpy_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub
