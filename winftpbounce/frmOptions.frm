VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Restore Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   660
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   838
      TabCaption(0)   =   "Preferences"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(2)=   "Label18"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "txtPassword"
      Tab(0).Control(5)=   "txtUsername"
      Tab(0).Control(6)=   "txtTelnetPort"
      Tab(0).Control(7)=   "txtUserID"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Sites"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemoveBounce"
      Tab(1).Control(1)=   "cmdRenameBounce"
      Tab(1).Control(2)=   "cmdShowBounce"
      Tab(1).Control(3)=   "cmdAddBounce"
      Tab(1).Control(4)=   "lvSites"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Allowed IPs"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvIPs"
      Tab(2).Control(1)=   "cmdCheck"
      Tab(2).Control(2)=   "cmdAdd"
      Tab(2).Control(3)=   "cmdRemoveIP"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Restriction Messages"
      TabPicture(3)   =   "frmOptions.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "txtSite"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtUL"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtDL"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chkSite"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "chkDL"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "chkUL"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin MSComctlLib.ListView lvSites 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Site"
            Object.Tag             =   "Site"
            Text            =   "Site"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "lport"
            Object.Tag             =   "lport"
            Text            =   "Listen Port"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "rport"
            Object.Tag             =   "rport"
            Text            =   "Remote Port"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.CommandButton cmdAddBounce 
         Caption         =   "Add Bounce"
         Height          =   495
         Left            =   -71280
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdShowBounce 
         Caption         =   "Edit Bounce"
         Height          =   495
         Left            =   -71280
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdRenameBounce 
         Caption         =   "Rename Bounce"
         Height          =   495
         Left            =   -71280
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemoveBounce 
         Caption         =   "Remove Bounce"
         Height          =   495
         Left            =   -71280
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   -73440
         TabIndex        =   21
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtTelnetPort 
         Height          =   285
         Left            =   -73440
         TabIndex        =   17
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   -73440
         TabIndex        =   16
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   -73440
         TabIndex        =   15
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkUL 
         Caption         =   "Upload Restricted"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1300
         Width           =   1935
      End
      Begin VB.CheckBox chkDL 
         Caption         =   "Download Restricted"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   700
         Width           =   1935
      End
      Begin VB.CheckBox chkSite 
         Caption         =   "Site Command Restricted"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1900
         Width           =   2535
      End
      Begin VB.TextBox txtDL 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtUL 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtSite 
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvIPs 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "IP"
            Object.Tag             =   "IP"
            Text            =   "IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Comment"
            Object.Tag             =   "Comment"
            Text            =   "Comment"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "CheckIP"
         Height          =   495
         Left            =   -71160
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   -71160
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdRemoveIP 
         Caption         =   "Remove"
         Height          =   495
         Left            =   -71160
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ident UserID"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Telnet Port"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Telnet Username"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Telnet Password"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmOptions.frx":0070
      Top             =   70
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options "
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
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDone As Boolean

Public Sub Display(ByVal Parent As Form)
    On Error GoTo PROC_ERR

    RefreshSettings
    
    Me.Show vbModeless, Parent
    
    Do Until mbDone = True
        DoEvents
    Loop
    
    mbDone = False
    
    Unload Me
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "Display" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo PROC_ERR
    Dim strIPAdd As String
    
    strIPAdd = InputBox("Enter an IP Mask.", "Allowed IP")
    
    If strIPAdd <> vbNullString Then
        If InStr(1, strIPAdd, "@") <> 0 Then
            AddIP strIPAdd
        Else
            MsgBox "Format is ident@ip, ex: user@127.0.0.1", vbCritical, "Error"
        End If
    End If

PROC_EXIT:
    FillIps
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdAdd_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdAddBounce_Click()
    On Error GoTo PROC_ERR

    Dim objBounce As frmBounce
    Dim strSiteName As String
    Dim lngCount As Long
    Dim astrBounces() As String
    Set objBounce = New frmBounce
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces"
        .EnumerateSections astrBounces(), lngCount
        Do Until Len(strSiteName) > 0
            strSiteName = InputBox("Enter the name for the new site.", "Add Site", "Site " & CStr(lngCount + 1))
            If InArray(astrBounces(), strSiteName) Then
                MsgBox "Name '" & strSiteName & "' already used. Please select another name."
                strSiteName = ""
            End If
        Loop
        .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
        .CreateKey
        .ValueType = REG_SZ
        .ValueKey = "IP"
        .Value = ""
        .ValueKey = "Port"
        .Value = ""
        .ValueKey = "ListenPort"
        .Value = ""
        .ValueKey = "MaxUsers"
        .Value = "10"
        objBounce.Display strSiteName, Me
        
    End With
        
PROC_EXIT:
    Unload objBounce
    Set objBounce = Nothing
    FillSites
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdAddBounce_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo PROC_ERR
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce"
        .ValueType = REG_SZ
        .ValueKey = "UserID"
        .Value = txtUserID.Text
        gstrIdent = txtUserID.Text
        .ValueKey = "AllowDLMessage"
        .Value = txtDL.Text
        gstrDlMsg = txtDL.Text
        .ValueKey = "AllowULMessage"
        .Value = txtUL.Text
        gstrULMsg = txtUL.Text
        .ValueKey = "AllowSiteMessage"
        .Value = txtSite.Text
        gstrSiteMsg = txtSite.Text
        .ValueKey = "TelnetPort"
        gvarTelnetPort = txtTelnetPort.Text
        If .Value <> txtTelnetPort.Text Then
            .Value = txtTelnetPort.Text
            frmMain.WriteToLog "Server", "Telnet Server Stopped (options, port changed)"
            DoEvents
            frmMain.SetupTelnet
            DoEvents
        Else
            .Value = txtTelnetPort.Text
        End If
        .ValueKey = "TelnetUsername"
        .Value = txtUsername.Text
        gstrLogin = txtUsername.Text
        .ValueKey = "TelnetPassword"
        .Value = txtPassword.Text
        gstrPassword = txtPassword.Text
        .ValueKey = "AllowDL"
        .Value = IIf(chkDL.Value = vbChecked, "no", "yes")
        gbDlToggle = IIf(chkDL.Value = vbChecked, True, False)
        .ValueKey = "AllowSite"
        .Value = IIf(chkSite.Value = vbChecked, "no", "yes")
        gbSiteToggle = IIf(chkSite.Value = vbChecked, True, False)
        .ValueKey = "AllowUL"
        .Value = IIf(chkUL.Value = vbChecked, "no", "yes")
        gbULToggle = IIf(chkUL.Value = vbChecked, True, False)
    End With
    
    If MsgBox("Do you want to restart the bounces now? (if you do, all current connections will be lost. If not, the changes will take effect the next time you start wBounce.)", vbYesNo, "Restart") = vbYes Then
        frmMain.SetupBounces
    End If
    
    gRefreshSettings
    
PROC_EXIT:
    mbDone = True
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdSave_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdCheck_Click()
    On Error GoTo PROC_ERR
    
    MsgBox CheckIP(InputBox("Enter IP", "Check IP"))
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdCheck_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdClose_Click()
    On Error GoTo PROC_ERR
    mbDone = True
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdClose_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub


Public Sub FillIps()
    ' fills the list of ip's on the settings screen
    
    On Error GoTo PROC_ERR
    
    Dim strIPAdd As String
    Dim astrKeys() As String
    Dim lngKeyCount As Long
    Dim intKey As Integer
    Dim oItem As ListItem
    
    lvIPs.ListItems.Clear
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\AllowedIPs"
        .EnumerateValues astrKeys(), lngKeyCount
        For intKey = 1 To lngKeyCount
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\GH\FTPBounce\AllowedIPs"
            .ValueKey = astrKeys(intKey)
            strIPAdd = .Value
            If strIPAdd <> "" Then
                Set oItem = lvIPs.ListItems.Add(, strIPAdd, strIPAdd)
            End If
        Next intKey
    End With
    
PROC_EXIT:
    Set oItem = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "FillIps" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub FillSites()
    On Error GoTo PROC_ERR

    Dim lngCount As Long
    Dim astrBounces() As String
    Dim intLoop As Integer
    Dim oItem As ListItem
    
    With goReg
        lvSites.ListItems.Clear
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces"
        .EnumerateSections astrBounces(), lngCount
        For intLoop = 1 To lngCount
            Set oItem = lvSites.ListItems.Add(, , astrBounces(intLoop))
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & astrBounces(intLoop)
            .ValueKey = "ListenPort"
            oItem.SubItems(1) = .Value
            .ValueKey = "Port"
            oItem.SubItems(2) = .Value
        Next intLoop
    End With

PROC_EXIT:
    Set oItem = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "FillSites" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub RefreshSettings()
    ' fills all the controls on the settings screen, and all the global/module variables needed.
    
    On Error GoTo PROC_ERR
        
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce"
        .ValueKey = "UserID"
        txtUserID.Text = .Value
        gstrIdent = .Value
        .ValueKey = "AllowDL"
        chkDL.Value = IIf(.Value = "yes", vbUnchecked, vbChecked)
        gbDlToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowUL"
        chkUL.Value = IIf(.Value = "yes", vbUnchecked, vbChecked)
        gbULToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowSite"
        chkSite.Value = IIf(.Value = "yes", vbUnchecked, vbChecked)
        gbSiteToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowDLMessage"
        txtDL.Text = .Value
        gstrDlMsg = .Value
        .ValueKey = "AllowULMessage"
        txtUL.Text = .Value
        gstrULMsg = .Value
        .ValueKey = "AllowSiteMessage"
        txtSite.Text = .Value
        gstrSiteMsg = .Value
        .ValueKey = "TelnetPort"
        txtTelnetPort.Text = .Value
        gvarTelnetPort = .Value
        .ValueKey = "TelnetUsername"
        txtUsername.Text = .Value
        gstrLogin = .Value
        .ValueKey = "TelnetPassword"
        txtPassword.Text = .Value
        gstrPassword = .Value
    End With
    
    FillIps
    FillSites
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "RefreshSettings" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdDefaults_Click()
    On Error GoTo PROC_ERR

    If MsgBox("Are you sure you want to overwrite all of your changes wit hthe default settings?", vbYesNo, "Restore") = vbYes Then
        DefaultSettings
    End If
    
PROC_EXIT:
    RefreshSettings
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdDefaults_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdRemoveBounce_Click()
    On Error GoTo PROC_ERR

    If Not lvSites.SelectedItem Is Nothing Then
        With goReg
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .DeleteKey
        End With
    End If

PROC_EXIT:
    FillSites
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdRemoveBounce_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdRemoveIP_Click()
    On Error GoTo PROC_ERR
    
    If Not lvIPs.SelectedItem Is Nothing Then
        RemoveIP lvIPs.SelectedItem.Text
    End If
    
PROC_EXIT:
    FillIps
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "cmdRemoveIP_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdRenameBounce_Click()
    On Error GoTo PROC_ERR

    Dim objBounce As frmBounce
    Dim strSiteName As String
    Dim lngCount As Long
    Dim astrBounces() As String
    Dim sIP As String
    Dim sPort As String
    Dim sListenPort As String
    Dim sMaxUsers As String
    
    If Not lvSites.SelectedItem Is Nothing Then
        With goReg
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\GH\FTPBounce\Bounces"
            .EnumerateSections astrBounces(), lngCount
            Do Until Len(strSiteName) > 0
                strSiteName = InputBox("Enter the new name for the site.", "Add Site", "Site " & CStr(lngCount + 1))
                If InArray(astrBounces(), strSiteName) Then
                    MsgBox "Name '" & strSiteName & "' already used. Please select another name."
                    strSiteName = ""
                End If
            Loop
            
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
            .CreateKey
            .ValueType = REG_SZ
            
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .ValueKey = "IP"
            sIP = .Value
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
            .ValueKey = "IP"
            .Value = sIP
            
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .ValueKey = "Port"
            sPort = .Value
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
            .ValueKey = "Port"
            .Value = sPort
            
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .ValueKey = "ListenPort"
            sListenPort = .Value
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
            .ValueKey = "ListenPort"
            .Value = sListenPort
                        
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .ValueKey = "MaxUsers"
            sMaxUsers = .Value
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & strSiteName
            .ValueKey = "MaxUsers"
            .Value = sMaxUsers
            
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & lvSites.SelectedItem.Text
            .DeleteKey
        End With
    End If

PROC_EXIT:
    FillSites
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdRenameBounce_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdShowBounce_Click()
    On Error GoTo PROC_ERR

    Dim objBounce As frmBounce
    Set objBounce = New frmBounce
    
    If Not lvSites.SelectedItem Is Nothing Then
        objBounce.Display lvSites.SelectedItem.Text, Me
    End If
    
PROC_EXIT:
    Unload objBounce
    Set objBounce = Nothing
    FillSites
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmOptions" & vbCrLf & "Procedure: " & "cmdShowBounce_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

