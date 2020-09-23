VERSION 5.00
Begin VB.Form frmBounce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bounce Info"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMaxConnections 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000011&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtListenPort 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Connections"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "IP"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Port"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Listen Port"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbDone As Boolean

Public Sub Display(ByVal pstrSiteName As String, ByRef pobjOwner As Form)
    
    On Error GoTo PROC_ERR
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces\" & pstrSiteName
        txtName = pstrSiteName
        .ValueKey = "IP"
        txtIP = .Value
        .ValueKey = "Port"
        txtPort = .Value
        .ValueKey = "ListenPort"
        txtListenPort = .Value
        .ValueKey = "MaxUsers"
        txtMaxConnections.Text = .Value
    End With

    Me.Show vbModeless, pobjOwner
    
    Do Until mbDone = True
        DoEvents
    Loop

    mbDone = False

    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounce" & vbCrLf & "Procedure: " & "Display" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub cmdCancel_Click()
    
    On Error GoTo PROC_ERR
    
    mbDone = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounce" & vbCrLf & "Procedure: " & "cmdCancel_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo PROC_ERR
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces\" & txtName.Text
        .ValueType = REG_SZ
        .ValueKey = "IP"
        If Len(txtIP.Text) > 6 Then
            .Value = txtIP.Text
        End If
        .ValueKey = "Port"
        If txtPort.Text <> "" And IsNumeric(txtPort.Text) Then
            .Value = txtPort.Text
        End If
        .ValueKey = "ListenPort"
        If txtListenPort.Text <> "" And IsNumeric(txtListenPort.Text) Then
            .Value = txtListenPort.Text
        End If
        .ValueKey = "MaxUsers"
        If txtMaxConnections.Text <> "" And IsNumeric(txtMaxConnections.Text) Then
            .Value = txtMaxConnections.Text
        End If
    End With
    
PROC_EXIT:
    mbDone = True
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounce" & vbCrLf & "Procedure: " & "cmdCancel_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub Form_Load()
    'frmMain.mTT.AddTool txtIP, "This is the IP or HOST that users connecting on the" & vbCrLf & "ListenPort will be bounced/redirected to", , , True, "IP", ttiInfo, ttsmDefault, 5, False, False, ttaLeft
    'frmMain.mTT.AddTool txtPort, "This is the Port for the IP or HOST that users will be connecting to", , , True, "Port", ttiInfo, ttsmDefault, 5, False, False, ttaLeft
    'frmMain.mTT.AddTool txtListenPort, "This is the Port that the bounce should Listen for connections on.", , , True, "Listen Port", ttiInfo, ttsmDefault, 5, False, False, ttaLeft
    'frmMain.mTT.AddTool txtMaxConnections, "Total connections that this bounce can have at 1 time.", , , True, "Max Connections", ttiInfo, ttsmDefault, 5, False, False, ttaLeft
End Sub

Private Sub txtIP_GotFocus()
    SelectAll txtIP
End Sub

Private Sub txtListenPort_GotFocus()
    SelectAll txtListenPort
End Sub

Private Sub txtMaxConnections_GotFocus()
    SelectAll txtMaxConnections
End Sub

Private Sub txtPort_GotFocus()
    SelectAll txtPort
End Sub

