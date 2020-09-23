VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Win FTP Bouncer"
   ClientHeight    =   5205
   ClientLeft      =   600
   ClientTop       =   1365
   ClientWidth     =   8295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8295
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6840
      Top             =   5520
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   6120
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2284
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":281E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkOld 
      Caption         =   "History"
      Height          =   255
      Left            =   6500
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chkConnected 
      Caption         =   "Current"
      Height          =   255
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2770
      Width           =   735
   End
   Begin MSComctlLib.ListView lvLog 
      Height          =   2135
      Left            =   0
      TabIndex        =   2
      Top             =   610
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3757
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "TimeStamp"
         Object.Tag             =   "TimeStamp"
         Text            =   "TimeStamp"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Origin"
         Object.Tag             =   "Origin"
         Text            =   "Origin"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Description"
         Object.Tag             =   "Description"
         Text            =   "Description"
         Object.Width           =   10108
      EndProperty
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   2160
      Left            =   0
      TabIndex        =   4
      Top             =   3045
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Connected"
         Object.Tag             =   "Connected"
         Text            =   "Connected"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "IP"
         Object.Tag             =   "IP"
         Text            =   "IP"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "To"
         Object.Tag             =   "To"
         Text            =   "Bouncing To"
         Object.Width           =   4586
      EndProperty
   End
   Begin MSComctlLib.ListView lvHistory 
      Height          =   2160
      Left            =   0
      TabIndex        =   13
      Top             =   3045
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Connected"
         Object.Tag             =   "Connected"
         Text            =   "Connected"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "IP"
         Object.Tag             =   "IP"
         Text            =   "IP"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "To"
         Object.Tag             =   "To"
         Text            =   "Bouncing To"
         Object.Width           =   4586
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsTelnet 
      Left            =   6840
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   7800
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock wsListen 
      Index           =   0
      Left            =   7320
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock wsClient 
      Index           =   0
      Left            =   8280
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsIdentListen 
      Left            =   7320
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock wsIdentClient 
      Index           =   0
      Left            =   6840
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsIdentReply 
      Index           =   0
      Left            =   7800
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":3352
      Top             =   70
      Width           =   480
   End
   Begin VB.Label lblTotalDownloaded 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total Bytes  Downloaded"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblTotalDownloads 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Total Downloads"
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblTotalUploads 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Uploads"
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblMiddle 
      Appearance      =   0  'Flat
      BackColor       =   &H0051B6F2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Connections (right click on a connection for options)                /                     Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server Information "
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
      Left            =   10
      TabIndex        =   1
      Top             =   300
      Width           =   8260
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
      Height          =   620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Menu mnuBouncer 
      Caption         =   "&Bouncer"
      Begin VB.Menu mnuSideBar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Main|FONT:Tahoma|FCOLOR:&H00FFFFFF&|BCOLOR:&H0051B6F2&|FSIZE:12|GRADIENT}"
      End
      Begin VB.Menu mnuBouncerRestart 
         Caption         =   "{IMG:5}&Restart Bounces"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBounceInformation 
         Caption         =   "{IMG:6}&Bounce Information"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBouncerOptions 
         Caption         =   "{IMG:2}&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBouncerExit 
         Caption         =   "{IMG:1}E&xit ALT+F4"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSidebar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Help|FONT:Tahoma|FCOLOR:&H00FFFFFF&|BCOLOR:&H0051B6F2&|FSIZE:12|GRADIENT}"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "{IMG:3}&Help F2"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "{IMG:4}&About"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuConnections 
      Caption         =   "Connections"
      Visible         =   0   'False
      Begin VB.Menu mnuConnectionsSidebar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Users|FONT:Tahoma|FCOLOR:&H00FFFFFF&|BCOLOR:&H0051B6F2&|FSIZE:12|GRADIENT}"
      End
      Begin VB.Menu mnuConnectionSpy 
         Caption         =   "{IMG:7}Spy / View Details"
      End
      Begin VB.Menu mnuConnectionDisconnect 
         Caption         =   "{IMG:8}Disconnect"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkConnected_Click()
    On Error Resume Next
    lvUsers.Visible = CBool(chkConnected.Value)
    lvHistory.Visible = Not CBool(chkConnected.Value)
    chkOld.Value = IIf(chkConnected.Value = vbChecked, vbUnchecked, vbChecked)
    lvLog.SetFocus
End Sub

Private Sub chkOld_Click()
    On Error Resume Next
    lvUsers.Visible = Not CBool(chkOld.Value)
    lvHistory.Visible = CBool(chkOld.Value)
    chkConnected.Value = IIf(chkOld.Value = vbChecked, vbUnchecked, vbChecked)
    lvLog.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo PROC_ERR
    
    Dim bFirstTime As Boolean
    Set goReg = New cRegistry
    
    Set gOClients = New Collection
    Set gOOldClients = New Collection
    Set goIdents = New CVector
    
    gbShutdown = False
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .ValueType = REG_SZ
        .SectionKey = "Software\GH\FTPBounce"
        If Not .KeyExists Then
            bFirstTime = True
            DefaultSettings
        End If
    End With
    
    gRefreshSettings
    
    chkConnected.Value = vbChecked
    chkOld.Value = vbUnchecked
    chkConnected_Click

    If bFirstTime = True Then
        frmFirst.Show
        OnTop frmFirst, True
        frmFirst.SetFocus
    End If
    
    WriteToLog "Ident", "Ident Daemon Initialized (will start on connection)"

    SetupTelnet
    
    SetupBounces
    
    SetMenus Me.hwnd, IL
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "Form_Load" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not ShutDown()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo PROC_ERR
        
    ReleaseMenus Me.hwnd
    
    wsTelnet.Close
    
    Set goReg = Nothing

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "Form_Unload" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub lvUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo PROC_ERR
    
    If Button = vbRightButton Then
        If Not lvUsers.SelectedItem Is Nothing Then
            PopupMenu mnuConnections
        End If
    End If

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "lvUsers_MouseUp" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub mnuBounceInformation_Click()
    frmBounces.Display Me
End Sub

Private Sub mnuBouncerExit_Click()
    On Error GoTo PROC_ERR
    
    Unload Me

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "mnuBouncerExit_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub mnuBouncerOptions_Click()
    On Error GoTo PROC_ERR
        
    frmOptions.Display Me

PROC_EXIT:
    Unload frmOptions
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "mnuBouncerOptions_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub mnuBouncerRestart_Click()
    On Error GoTo PROC_ERR

    WriteToLog "Bounce", "Stoping Bounces"

    SetupBounces
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "mnuBouncerRestart_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub mnuConnectionDisconnect_Click()
    On Error Resume Next
    
    Dim ofrm As Form
    
    If Not lvUsers.SelectedItem Is Nothing Then
        WriteToLog "Server", "Kicked:" & wsClient(lvUsers.SelectedItem.Tag).RemoteHostIP
        wsServer(lvUsers.SelectedItem.Tag).Close
        wsClient(lvUsers.SelectedItem.Tag).Close
        CreateOld lvUsers.SelectedItem.Tag
        Unload wsClient(lvUsers.SelectedItem.Tag)
        Unload wsServer(lvUsers.SelectedItem.Tag)
    
        For Each ofrm In Forms
            If (ofrm.Name = "frmSpy" Or ofrm.Name = "frmDetails") And ofrm.Tag = CStr(lvUsers.SelectedItem.Tag) Then
                ofrm.mbDone = True
            End If
        Next ofrm
    
        Set ofrm = Nothing
    End If
    
End Sub

Private Sub mnuConnectionSpy_Click()
    On Error GoTo PROC_ERR
    
    Dim oDetails As frmDetails
    If Not lvUsers.SelectedItem Is Nothing Then
        Set oDetails = New frmDetails
        oDetails.Display lvUsers.SelectedItem.Tag, Me
    End If

PROC_EXIT:
    Unload oDetails
    Set oDetails = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "mnuConnectionSpy_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub WriteToLog(ByVal sOrigin As String, ByVal sMessage As String)
    On Error GoTo PROC_ERR
    
    Dim oItem As ListItem
    Set oItem = lvLog.ListItems.Add(, , Now)
    oItem.SubItems(1) = sOrigin
    oItem.SubItems(2) = sMessage
    oItem.EnsureVisible
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "WriteToLog" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub SetupTelnet()
    ' sets up the telnet listener, and all its variables.
    
    On Error GoTo PROC_ERR
    
    If wsTelnet.State <> sckClosed Then
        wsTelnet.Close
    End If
    
    wsTelnet.LocalPort = gvarTelnetPort
    
    
    wsTelnet.Listen
    gstrUserID = ""
    gstrUserPassword = ""
    gstrUserCommand = ""
    gbAcceptedId = False
    gbSuccessLoging = False
    
    WriteToLog "Telnet", "Telnet server started"

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "SetupTelnet" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    WriteToLog "Server", "Error starting telnet server."
    Resume PROC_EXIT
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpHelp_Click()
    MsgBox "No help file yet, care to write one?", vbCritical, "Oops"
End Sub

Private Sub Timer1_Timer()
    
    On Error GoTo PROC_ERR
    
    UpdateDisplay
    
    If gbShutdown = True Then
        If lvUsers.ListItems.Count <= 0 Then
            Me.Caption = "Shutting Down......"
            Unload Me
        End If
    End If

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "Timer1_Timer" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT

End Sub

Public Sub UpdateDisplay()
    
    On Error GoTo PROC_ERR
    
    ' setup main stats
    lblTotalDownloads.Caption = gintDl
    lblTotalUploads.Caption = gintUL
    lblTotalDownloaded.Caption = gdblDLB
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "UpdateDisplay" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Function ShutDown() As Boolean
    
    On Error GoTo PROC_ERR
    
    Dim intI As Integer
    If lvUsers.ListItems.Count > 0 Then
        intI = MsgBox("There are open connections." & vbCrLf & vbCrLf & "Would you like to wait until they finish before closing?", vbYesNo, "Shut Down")
        Select Case intI
            Case vbYes
                gbShutdown = True
                ShutDown = False
            Case vbNo
                Unload Me
                ShutDown = True
        End Select
    Else
        If Not gbShutdown Then
            MsgBox "No open connections, shutting down immediately", vbInformation, "Shutdown"
        End If
        Unload Me
        ShutDown = True
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "ShutDown" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Function

Private Sub wsTelnet_Close()
    'When user wants to close the telnet connection
    
    On Error GoTo PROC_ERR
    
    wsTelnet.Close 'Close the telnet port
    wsTelnet.LocalPort = gvarTelnetPort
    wsTelnet.Listen 'Listen for the new user
    
    'Initialisation of the telnet server variables
    gstrUserID = ""
    gstrUserPassword = ""
    gstrUserCommand = ""
    gbAcceptedId = False
    gbSuccessLoging = False
    WriteToLog "Telnet", "Connection closed"

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsTelnet_Close" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsTelnet_ConnectionRequest(ByVal requestID As Long)
    'User wants to connect to the server
    
    On Error GoTo PROC_ERR
    
    If wsTelnet.State <> sckClosed Then
        wsTelnet.Close
    End If
    wsTelnet.Accept requestID
    'Send him the accepted message and ask him to logon to the server
    wsTelnet.SendData "Login: "
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsTelnet_ConnectionRequest" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT

End Sub

Private Sub wsTelnet_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo PROC_ERR

    Dim strDataIn As String
    Dim strMyName As String
    Dim astrDirList() As String
    Dim intDirPointer As Integer
    Dim bSBool As Boolean
    Dim intI As Integer
    Dim oItem As ListItem
    
    'User sending some information
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce"
        .ValueType = REG_SZ
        
        wsTelnet.GetData strDataIn 'Receive the input from the client
        If gbSuccessLoging Then 'Check whether user had already logged in
            If Len(gstrUserCommand) > 0 Then
                If Asc(strDataIn) = 127 Then
                    gstrUserCommand = Left$(gstrUserCommand, Len(gstrUserCommand) - 1)
                End If
            End If
            If Right(strDataIn, 2) = vbCrLf Then
                gstrUserCommand = gstrUserCommand & Left(strDataIn, Len(strDataIn) - 2)
                If Trim(gstrUserCommand) Like "help" Or Trim(gstrUserCommand) = "?" Then
                    'User requested help
                    wsTelnet.SendData vbCrLf & "List of Commands supported by this server" & vbCrLf
                    wsTelnet.SendData "who                  -- to list the currently connected users." & vbCrLf
                    wsTelnet.SendData "stats                -- General statistics about the bouncers current session." & vbCrLf
                    wsTelnet.SendData "kick #               -- Kicks the user specified by # (from who)." & vbCrLf
                    wsTelnet.SendData "kickban #            -- Kickbans the user specified by # (from who)." & vbCrLf
                    wsTelnet.SendData "checkip IP           -- Checks an IP against the Allowed IP's list." & vbCrLf
                    wsTelnet.SendData "addip IP             -- Adds an IP to the Allowed IP's list." & vbCrLf
                    wsTelnet.SendData "removeip IP          -- Removes an IP from the Allowed IP's list." & vbCrLf
                    wsTelnet.SendData "listips              -- Lists all IP's on the Allowed IP's list." & vbCrLf
                    wsTelnet.SendData "setuser USERNAME     -- Changes the telnet username." & vbCrLf
                    wsTelnet.SendData "setpass PASSWORD     -- Changes the telnet password." & vbCrLf
                    wsTelnet.SendData "allowdl true/false   -- Allows or Disallows downloading." & vbCrLf
                    wsTelnet.SendData "allowdl              -- Displays download setting." & vbCrLf
                    wsTelnet.SendData "allowul true/false   -- Allows or Disallows uploading." & vbCrLf
                    wsTelnet.SendData "allowul              -- Displays upload setting." & vbCrLf
                    wsTelnet.SendData "allowsite true/false -- Allows or Disallows the Site command." & vbCrLf
                    wsTelnet.SendData "allowsite            -- Displays the Site command setting." & vbCrLf
                    wsTelnet.SendData "sitemsg MESSAGE      -- Changes the message displayed when disallowing Site." & vbCrLf
                    wsTelnet.SendData "dlmsg MESSAGE        -- Changes the message displayed when disallowing downloading." & vbCrLf
                    wsTelnet.SendData "ulmsg MESSAGE        -- Changes the message displayed when disallowing uploading." & vbCrLf
                    wsTelnet.SendData "sitemsg              -- Displays the message displayed when disallowing Site." & vbCrLf
                    wsTelnet.SendData "dlmsg                -- Displays the message displayed when disallowing downloading." & vbCrLf
                    wsTelnet.SendData "ulmsg                -- Displays the message displayed when disallowing uploading." & vbCrLf
                    wsTelnet.SendData "exit                 -- to quit logout from the server." & vbCrLf
                    wsTelnet.SendData "time                 -- to get the time of the server." & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "exit" Then
                    'User wants to terminate the session
                    wsTelnet_Close
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "stats" Then
                    'Send the user the server's current stats
                    wsTelnet.SendData vbCrLf & "-- Stats -- " & vbCrLf
                    wsTelnet.SendData "Uploads:          " & CStr(gintUL) & vbCrLf
                    wsTelnet.SendData "Downloads:        " & CStr(gintDl) & vbCrLf
                    wsTelnet.SendData "Downloaded Bytes: " & CStr(gdblDLB) & vbCrLf
                    wsTelnet.SendData "Connecitons:      " & CStr(gdblConnections) & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) = "dlmsg" Then
                    .ValueKey = "AllowDLMessage"
                    wsTelnet.SendData vbCrLf & "dlmsg: " & .Value & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "dlmsg *" Then
                    'Set the AllowDlMessage
                    wsTelnet.SendData vbCrLf & "Changing setting for Download message: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowDLMessage"
                    .Value = Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) = "ulmsg" Then
                    .ValueKey = "AllowULMessage"
                    wsTelnet.SendData vbCrLf & "ulmsg: " & .Value & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "ulmsg *" Then
                    'Set the AllowULMessage
                    wsTelnet.SendData vbCrLf & "Changing setting for Upload message: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowULMessage"
                    .Value = Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) = "sitemsg" Then
                    .ValueKey = "AllowSiteMessage"
                    wsTelnet.SendData vbCrLf & "sitemsg: " & .Value & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "sitemsg *" Then
                    'Set the AllowSiteMessage
                    wsTelnet.SendData vbCrLf & "Changing setting for Site message: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowSiteMessage"
                    .Value = Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowsite true" Or Trim(gstrUserCommand) Like "allowsite false" Then
                    'Set the AllowSite value
                    wsTelnet.SendData vbCrLf & "Changing setting for Site Restricted to: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowSite"
                    .Value = IIf(Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) = "true", "yes", "no")
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowul true" Or Trim(gstrUserCommand) Like "allowul false" Then
                    'Set the AllowUL value
                    wsTelnet.SendData vbCrLf & "Changing setting for Upload Restricted to: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowUL"
                    .Value = IIf(Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) = "true", "yes", "no")
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowdl true" Or Trim(gstrUserCommand) Like "allowdl false" Then
                    'Set the AllowDL value
                    wsTelnet.SendData vbCrLf & "Changing setting for Download Restricted to: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    .ValueKey = "AllowDL"
                    .Value = IIf(Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) = "true", "yes", "no")
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowsite" Then
                    'Sho the AllowSite value
                    .ValueKey = "AllowSite"
                    wsTelnet.SendData vbCrLf & "Setting for Site Restricted: " & IIf(.Value = "yes", "true", "false") & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowdl" Then
                    'Show the AllowDL value
                    .ValueKey = "AllowDL"
                    wsTelnet.SendData vbCrLf & "Setting for Download Restricted: " & IIf(.Value = "yes", "true", "false") & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "allowul" Then
                    'Show the AllowUL value
                    .ValueKey = "AllowUL"
                    wsTelnet.SendData vbCrLf & "Setting for Upload Restricted: " & IIf(.Value = "yes", "true", "false") & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "setpass *" Then
                    'Set the servers telnet password
                    wsTelnet.SendData vbCrLf & "Changing Password" & vbCrLf
                    .ValueKey = "TelnetPassword"
                    .Value = Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "setuser *" Then
                    'Set the servers telnet username
                    wsTelnet.SendData vbCrLf & "Changing Telnet Username" & vbCrLf
                    .ValueKey = "TelnetUsername"
                    .Value = Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "time" Then
                    'Send the user the server's current time
                    wsTelnet.SendData vbCrLf & Time & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "listips" Then
                    'Send the user the server's current list of allowed ips
                    wsTelnet.SendData vbCrLf & "-- Allowed IPs --" & vbCrLf
                    Dim strSIP As String
                    
                    Dim astrSKeys() As String
                    Dim lngIKeyCount As Long
                    Dim intIKey As Integer
                    GetIps astrSKeys, lngIKeyCount
                    For intIKey = 1 To lngIKeyCount
                        If astrSKeys(intIKey) <> "" Then
                            wsTelnet.SendData "IP" & intIKey & " : " & astrSKeys(intIKey) & vbCrLf
                        End If
                    Next intIKey
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "removeip *" Then
                    'Remove an allowed ip
                    wsTelnet.SendData vbCrLf & "Removing IP: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    RemoveIP Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "addip *" Then
                    'Add an allowed ip
                    wsTelnet.SendData vbCrLf & "Adding IP: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    AddIP Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    gRefreshSettings
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "checkip *" Then
                    'Check an allowed ip
                    wsTelnet.SendData vbCrLf & "Checking IP: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    bSBool = CheckIP(Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))))
                    wsTelnet.SendData "IP: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & " returns " & IIf(bSBool = True, "True", "False") & vbCrLf
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "kick #" Or Trim(gstrUserCommand) Like "kick ##" Or Trim(gstrUserCommand) Like "kick ###" Then
                    'Kick a user/disconnect
                    wsTelnet.SendData vbCrLf & "Kicking: " & Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " "))) & vbCrLf
                    SelectClient Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    mnuConnectionDisconnect_Click
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "kickban #" Or Trim(gstrUserCommand) Like "kickban ##" Or Trim(gstrUserCommand) Like "kickban ###" Then
                    'kick and remove that users IP from the allowed list
                    wsTelnet.SendData vbCrLf & "Kicking and Removing: " & Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")) & vbCrLf
                    SelectClient Trim(Mid(gstrUserCommand, InStr(1, gstrUserCommand, " ")))
                    'mnuConnectionDisconnectRemove_Click
                    Prompt
                    gstrUserCommand = ""
                ElseIf Trim(gstrUserCommand) Like "who" Then
                    If lvUsers.ListItems.Count > 0 Then
                        wsTelnet.SendData vbCrLf & Chr(9) & "ID" & Chr(9) & "User" & Chr(9) & Chr(9) & "IP" & Chr(9) & Chr(9) & "Downloads" & Chr(9) & Chr(9) & "Bytes Downloaded" & vbCrLf
                        wsTelnet.SendData "------------------------------------------------------------" & vbCrLf
                        For Each oItem In lvUsers.ListItems
                            Set gOClient = gOClients(CStr(oItem.Tag))
                            wsTelnet.SendData Chr(9) & oItem.Tag & Chr(9) & gOClient.User & Chr(9) & Chr(9) & gOClient.ClientIP & Chr(9) & Chr(9) & gOClient.Downloads & Chr(9) & Chr(9) & gOClient.Bytes & vbCrLf
                        Next oItem
                    Else
                        wsTelnet.SendData vbCrLf & "No users connected" & vbCrLf
                    End If
                    gstrUserCommand = ""
                    Prompt
                ElseIf Trim(gstrUserCommand) Like "whodisc" Then
                    If lvUsers.ListItems.Count > 0 Then
                        wsTelnet.SendData vbCrLf & Chr(9) & "ID" & Chr(9) & "User" & Chr(9) & Chr(9) & "IP" & Chr(9) & Chr(9) & "Downloads" & Chr(9) & Chr(9) & "Bytes Downloaded" & vbCrLf
                        wsTelnet.SendData "------------------------------------------------------------" & vbCrLf
                        For Each oItem In lvUsers.ListItems
                            Set gOClient = gOOldClients(CStr(oItem.Tag))
                            wsTelnet.SendData Chr(9) & oItem.Tag & Chr(9) & gOClient.User & Chr(9) & Chr(9) & gOClient.ClientIP & Chr(9) & Chr(9) & gOClient.Downloads & Chr(9) & Chr(9) & gOClient.Bytes & vbCrLf
                        Next oItem
                    Else
                        wsTelnet.SendData vbCrLf & "No users connected" & vbCrLf
                    End If
                    gstrUserCommand = ""
                    Prompt
                Else
                    'Its an invalied command or command not suported by the telnet server
                    wsTelnet.SendData vbCrLf & "Invalied Command" & vbCrLf & "For list of commands use the help command" & vbCrLf
                    gstrUserCommand = ""
                    Prompt
                End If
                
            Else
                wsTelnet.SendData strDataIn
                If Asc(strDataIn) <> 127 Then
                    gstrUserCommand = gstrUserCommand & strDataIn
                End If
            End If
        Else
            ' Authentication of the user
            
            If gbAcceptedId And Right(strDataIn, 2) = vbCrLf Then
                wsTelnet.SendData vbCrLf & "Verifying your login information...." & vbCrLf
                If gstrUserPassword = vbNullString And Len(strDataIn) >= 2 Then
                    gstrUserPassword = Left(strDataIn, Len(strDataIn) - 2)
                End If
                If gstrUserID = gstrLogin And gstrUserPassword = gstrPassword Then
                    'Sleep (2000)
                    wsTelnet.SendData "Welcome " & gstrLogin & vbCrLf
                    wsTelnet.SendData "help -- to get the list of commands supported by the server." & vbCrLf
                    Prompt
                    gbSuccessLoging = True
                    WriteToLog "Telnet", "User '" & gstrLogin & "' logged in"
                Else
                    gbAcceptedId = False
                    gstrUserID = ""
                    gstrUserPassword = ""
                    wsTelnet.SendData "Password Incorrect" & vbCrLf
                    wsTelnet.SendData "Login: "
                    WriteToLog "Telnet", "User failed authentication."
                    Exit Sub
                End If
            ElseIf Right(strDataIn, 2) = vbCrLf Then
                If gstrUserID = vbNullString And Len(strDataIn) >= 2 Then
                    gstrUserID = Left(strDataIn, Len(strDataIn) - 2)
                End If
                wsTelnet.SendData vbCrLf & "Enter Password:"
                gbAcceptedId = True
                Exit Sub
            ElseIf Not gbAcceptedId Then
                'wsTelnet.SendData strDataIn
            End If
            If gbAcceptedId Then
                If Right(strDataIn, 2) <> vbCrLf Then
                    gstrUserPassword = gstrUserPassword & strDataIn
                End If
            Else
                If IsAlphaNumeric(strDataIn) Then
                    gstrUserID = gstrUserID & strDataIn
                End If
            End If
        End If
    End With
    
PROC_EXIT:
    Set oItem = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsTelnet_DataArrival" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub wsTelnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo PROC_ERR

    WriteToLog "Telnet", "Error occured in telnet session, disconnecting user."
    wsTelnet_Close
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsTelnet_Error" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub Prompt()
    
    On Error GoTo PROC_ERR
    
    wsTelnet.SendData "Prompt>"
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "Prompt" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub SelectClient(ByVal pvarData As Variant)
    
    On Error GoTo PROC_ERR
    
    Dim intI As Integer
    Dim oItem As ListItem
    
    If lvUsers.ListItems.Count > 0 Then
        For Each oItem In lvUsers.ListItems
            If oItem.Tag = pvarData Then
                oItem.Selected = True
                Exit For
            End If
        Next oItem
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "SelectClient" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsIdentClient_Close(Index As Integer)
    If goIdents(Index) = "" Then
        goIdents(Index) = "*"
    End If
    Unload wsIdentClient(Index)
End Sub

Private Sub wsIdentClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Do While wsIdentClient(Index).State = sckInProgress
        ' wait for it to finish
    Loop
    wsIdentClient(Index).GetData strData
    goIdents(Index) = strData
    wsIdentClient(Index).Close
End Sub

Private Sub wsIdentClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If goIdents(Index) = "" Then
        goIdents(Index) = "*"
    End If
    Unload wsIdentClient(Index)
End Sub

Private Sub wsIdentListen_ConnectionRequest(ByVal requestID As Long)
    
    On Error GoTo PROC_ERR
    
    Load wsIdentReply(wsIdentReply.Count)
    wsIdentReply(wsIdentReply.Count - 1).Accept requestID
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsIdentListen_ConnectionRequest" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsIdentReply_Close(Index As Integer)
    
    On Error GoTo PROC_ERR
    
    Unload wsIdentReply(Index)
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsIdentReply_Close" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsIdentReply_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    On Error GoTo PROC_ERR
    
    Dim strData As String
    Dim strDelim As String
    Dim lngCPort As Long
    Dim lngSPort As Long
    Dim lngI As Long
    Dim strX As String
    Dim sendstr As String
    Dim ws As Winsock
    sendstr = ""
    lngCPort = 0
    lngSPort = 0
    strX = ""
    ' get the data from the server
    WriteToLog "Ident", "Incoming request from: " & wsIdentReply(Index).RemoteHostIP
    wsIdentReply(Index).GetData strData
    strData = Trim(strData)
    ' loop until we get all the data
    Do Until wsIdentReply(Index).State <> sckInProgress
        DoEvents
    Loop
    ' check and see if request is valid
    If InStr(1, strData, ",") = 0 Then
        WriteToLog "Ident", "Invalid Request from " & wsIdentReply(Index).RemoteHostIP
    Else
        lngI = 1
        strX = Mid(strData, lngI, 1)
        ' go through and parse for client/server ports..
        While strX <> "," And lngI <= Len(strData)
            strX = Mid(strData, lngI, 1)
            If strX <> " " Then
                lngSPort = lngSPort & strX
            End If
            lngI = lngI + 1
        Wend
        lngCPort = Mid(strData, lngI + 1, Len(strData) - (lngI + 1))
        '  MsgBox "whole string: " & strData & "len: " & Len(strData)
        '  MsgBox "server port: " & sPort & vbCrLf & "client port: " & cPort
        ' remove all stuff we don't need...
        strData = Trim(strData)
        ' remove linefeed & delim
        strDelim = Right(strData, 2)
        strData = Mid(strData, 1, Len(strData) - 2)
        ' build the string we need to send and write to socket
        sendstr = strData & " : USERID : " & "WIN32" & " : " & gstrIdent & strDelim
        wsIdentReply(Index).SendData sendstr
        WriteToLog "Ident", "Received query [" & strData & "] from " & wsIdentReply(Index).RemoteHostIP
        ' code used in testing pass-through ident, cant figure out how
'        For Each ws In wsServer
'            WriteToLog ws.RemotePort & " = Remote"
'            WriteToLog ws.LocalPort & " = Local"
'        Next ws
'
'        For Each ws In wsClient
'            WriteToLog ws.RemotePort & " = Remote Client"
'            WriteToLog ws.LocalPort & " = Local Client"
'        Next ws
        DoEvents
    End If
    wsIdentReply(Index).Close
    Unload wsIdentReply(Index)
    
    wsIdentListen.Close
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsIdentReply_DataArrival" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsIdentReply_SendComplete(Index As Integer)
    
    On Error GoTo PROC_ERR
    
    WriteToLog "Ident", "Replied to " & wsIdentReply(Index).RemoteHost & "with username: " & gstrIdent & " system type: " & "WIN32"
    wsIdentReply(Index).Close
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsIdentReply_SendComplete" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub RemoveClient(ByVal plngIndex As Long)
    
    On Error GoTo PROC_ERR
    
    Dim oItem As ListItem
    Dim lIndex As Integer
    lIndex = -1
    For Each oItem In lvUsers.ListItems
        If oItem.Tag = plngIndex Then
            lIndex = oItem.Index
        End If
    Next oItem
        
    If lIndex > -1 Then
        lvUsers.ListItems.Remove lIndex
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "RemoveClient" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Public Sub CreateOld(ByVal plngIndex As Long)
    
    On Error GoTo PROC_ERR
    
    Dim oItem As ListItem
    Dim Client As cClients
    Set Client = gOClients(CStr(plngIndex))
    gOOldClients.Add Client, CStr(plngIndex)
    Set oItem = lvHistory.ListItems.Add(, , Now())
    oItem.SubItems(1) = Client.ClientIP
    oItem.SubItems(2) = Client.Site
    gOClients.Remove CStr(plngIndex)
    RemoveClient plngIndex
    
    If IsFormLoaded("frmBounces") Then
        frmBounces.BuildSiteStats
    End If
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "CreateOld" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsClient_Close(pintIndex As Integer)
    
    'If the remote client disconnects, disconnect our server
    'connection and unload them both
    
    On Error GoTo PROC_ERR
    
    wsServer(pintIndex).Close
    wsClient(pintIndex).Close
    
    If gaintMaxUsers(wsClient(pintIndex).Tag) <> 0 Then
        gaintMaxUsers(wsClient(pintIndex).Tag) = gaintMaxUsers(wsClient(pintIndex).Tag) - 1
    End If
    
    Unload wsClient(pintIndex)
    Unload wsServer(pintIndex)
    
    CreateOld pintIndex
    'UpdateSiteStats
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsClient_Close" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsClient_DataArrival(pintIndex As Integer, ByVal plngBytesTotal As Long)
    
    On Error GoTo PROC_ERR
    
    Dim strDataIn As String
    Dim Client As cClients
    
    wsClient(pintIndex).GetData strDataIn, vbString
    Set Client = gOClients(CStr(pintIndex))
    If InStr(1, strDataIn, "RETR ") <> 0 Then
        Client.DownloadAdd
        gintDl = gintDl + 1
    End If
    If InStr(1, strDataIn, "STOR ") <> 0 Then
        Client.UploadAdd
        gintUL = gintUL + 1
    End If
    
    If Left(strDataIn, 3) = "CWD" <> 0 Then
        ' CWD
        Client.CWD = Trim(Right(strDataIn, Len(strDataIn) - 3))
    End If
    
    If InStr(1, strDataIn, "USER ") <> 0 Then
        ' User
        Client.User = Left(strDataIn, Len(strDataIn) - 2)
        Client.User = Right(Client.User, Len(Client.User) - 5)
    End If
    
    If Len(Client.ServerCommands) >= 20000 Then
        Client.ServerCommands = ""
    End If
    
    Client.ServerCommands = Client.ServerCommands & strDataIn
    
    If Len(Client.Log) >= 20000 Then
        Client.Log = ""
    End If
    
    Client.Log = Client.Log & strDataIn
    
    If InStr(1, strDataIn, "RETR ") <> 0 And gbDlToggle = False Then
        wsClient(pintIndex).SendData "226 " & gstrDlMsg & vbCrLf
    ElseIf InStr(1, strDataIn, "STOR ") <> 0 And gbULToggle = False Then
        wsClient(pintIndex).SendData "226 " & gstrULMsg & vbCrLf
    ElseIf InStr(1, strDataIn, "SITE ") <> 0 And gbSiteToggle = False Then
        wsClient(pintIndex).SendData "226 " & gstrSiteMsg & vbCrLf
    Else
        serverData strDataIn, pintIndex
    End If
    gOClients.Remove CStr(pintIndex)
    DoEvents
    gOClients.Add Client, CStr(pintIndex)
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Select Case Err.Number
        Case 5
            Resume PROC_EXIT
        Case 457
            gOClients.Remove CStr(pintIndex)
            gOClients.Add Client, CStr(pintIndex)
            Resume Next
        Case Else
            MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsClient_DataArrival" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
            Resume PROC_EXIT
    End Select
    
    '    Select Case Err.Number
    '        Case 5
    '            Resume exit_routine
    '        Case 457
    '            mOClients.Remove CStr(pintIndex)
    '            mOClients.Add Client, CStr(pintIndex)
    '            Resume Next
    '        Case Else
    '            MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsClient_DataArrival" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    '            Resume PROC_EXIT
    '    End Select
    
End Sub

Private Sub wsClient_Error(pintIndex As Integer, ByVal pintNumber As Integer, pstrDescription As String, ByVal plngScode As Long, ByVal pstrSource As String, ByVal pstrHelpFile As String, ByVal plngHelpContext As Long, pbCancelDisplay As Boolean)
    On Error Resume Next
    wsClient(pintIndex).Close
    wsServer(pintIndex).Close
    
    Unload wsClient(pintIndex)
    Unload wsServer(pintIndex)
    
    CreateOld pintIndex
    
    pbCancelDisplay = True
    
End Sub

Private Sub serverData(pstrChrSend As String, pintIndex As Integer)
    
    On Error GoTo PROC_ERR
    
    If wsServer(pintIndex).State = 7 Then
        wsServer(pintIndex).SendData pstrChrSend
    ElseIf wsServer(pintIndex).State = 9 Or wsServer(pintIndex).State = 0 Then
        wsServer(pintIndex).Close
        wsServer(pintIndex).RemoteHost = ""
        wsServer(pintIndex).RemotePort = 0
        wsClient(pintIndex).Close
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "serverData" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsServer_Close(pintIndex As Integer)
    On Error Resume Next
    'When the remote server disconnects, close the port,
    wsServer(pintIndex).Close
    Unload wsServer(pintIndex)
    Unload wsClient(pintIndex)
    
    CreateOld pintIndex
    
End Sub

Private Sub wsServer_DataArrival(pintIndex As Integer, ByVal plngBytesTotal As Long)
    'When data arrives, if we are connected, forward it on
    
    On Error GoTo PROC_ERR
    
    Dim strDataIn As String
    Dim Client As cClients
    Dim sCWD As String
    wsServer(pintIndex).GetData strDataIn, vbString
    
    Set Client = gOClients(CStr(pintIndex))
    
    If Len(Client.ClientCommands) >= 20000 Then
        Client.ClientCommands = ""
    End If
    
    Client.ClientCommands = Client.ClientCommands & strDataIn
    
    If InStr(1, strDataIn, "550 ") <> 0 Then
        ' CWD
        Client.CWD = ""
    End If
    
    If Left(strDataIn, 4) = "257 " Then
        ' CWD
        If Len(strDataIn) > 5 Then
            sCWD = Right(strDataIn, Len(strDataIn) - 5)
            sCWD = Left(sCWD, InStr(1, sCWD, Chr(34)) - 1)
            Client.CWD = sCWD
        End If
    End If
    
    If Len(Client.Log) >= 20000 Then
        Client.Log = ""
    End If
    
    Client.Log = Client.Log & strDataIn
    
    UpdateSpyWindows Client.Index, strDataIn
    
    If InStr(1, strDataIn, " bytes).") <> 0 And InStr(1, strDataIn, "[Ratio") = 0 Then
        Dim strTemp As String
        strTemp = CleanString(strDataIn)
        strTemp = Left(strTemp, Len(strTemp) - 8)
        strTemp = Mid(strTemp, InStrRev(strTemp, "(") + 1)
        Client.BytesAdd CDbl(strTemp)
        gdblDLB = gdblDLB + CDbl(strTemp)
    End If
    If wsClient(pintIndex).State = 7 Then
        wsClient(pintIndex).SendData strDataIn
    End If
    gOClients.Remove CStr(pintIndex)
    gOClients.Add Client, CStr(pintIndex)
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Select Case Err.Number
        Case 5
            Resume PROC_EXIT
        Case Else
            MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsServer_DataArrival" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
            Resume PROC_EXIT
    End Select
    
    '
    '    Select Case Err.Number
    '        Case 5
    '            Resume exit_routine
    '        Case Else
    '            MsgBox Err.Number & " : " & Err.Description
    '            Resume exit_routine
    '    End Select
    
End Sub

Private Sub wsServer_Error(pintIndex As Integer, ByVal pintNumber As Integer, pstrDescription As String, ByVal plngScode As Long, ByVal pstrSource As String, ByVal pstrHelpFile As String, ByVal plngHelpContext As Long, pbCancelDisplay As Boolean)
    
    On Error GoTo PROC_ERR
    
    wsClient(pintIndex).Close
    wsServer(pintIndex).Close
    
    Unload wsClient(pintIndex)
    Unload wsServer(pintIndex)
    
    CreateOld pintIndex
    
    pbCancelDisplay = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsServer_Error" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub wsListen_ConnectionRequest(Index As Integer, ByVal plngRequestID As Long)
    
    On Error GoTo PROC_ERR
    
    ConnectionRequest plngRequestID, Index
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "wsListen_ConnectionRequest" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Sub ConnectionRequest(ByVal plngRequestID As Long, ByVal pintListen As Integer)
    'When a connection request is heard by the dedicated
    'listener, spawn a new client/server pair to handle the
    'request.
    
    On Error GoTo PROC_ERR
    
    Dim oItem As ListItem
    Dim Client As cClients
    Dim intIndex As Integer
    
    intIndex = GetIndex()
    Do Until IsLoaded(wsServer(intIndex)) = False
        intIndex = GetIndex()
    Loop
    Set Client = New cClients
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces\" & wsListen(pintListen).Tag
        .ValueKey = "MaxUsers"
        If gaintMaxUsers(pintListen) >= .Value Then
            wsListen(pintListen).SendData "226 MAX USERS REACHED, PLEASE TRY AGAIN LATER"
            DoEvents
            wsListen(pintListen).Close
            DoEvents
            wsListen(pintListen).Listen
            GoTo PROC_EXIT
        ElseIf gaintMaxUsers(pintListen) < .Value Then
            gaintMaxUsers(pintListen) = gaintMaxUsers(pintListen) + 1
        End If
    End With
    
    Load wsServer(intIndex)
    Load wsClient(intIndex)
    wsClient(intIndex).Accept plngRequestID
    wsClient(intIndex).Tag = pintListen
    wsServer(intIndex).Tag = pintListen
    Client.Index = intIndex
    Client.ClientIP = wsClient(intIndex).RemoteHostIP
    Client.ClientPort = wsListen(pintListen).RemotePort
    Client.TimeStamp = Now
    
    ' start the ident server
    If Not wsIdentListen.State = sckClosed Then
        wsIdentListen.Close
    End If
    wsIdentListen.Listen
    
    '----
    ' Get Ident from user here....
    Load wsIdentClient(intIndex)
    wsIdentClient(intIndex).RemotePort = 113
    wsIdentClient(intIndex).RemoteHost = wsClient(intIndex).RemoteHostIP
    
    '----
    'If CheckIP(wsClient(intIndex).RemoteHostIP) = False Then
    'Else
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces\" & wsListen(pintListen).Tag
        .ValueKey = "Port"
        wsServer(intIndex).RemotePort = .Value
        Client.ServerPort = .Value
        .ValueKey = "IP"
        wsServer(intIndex).RemoteHost = .Value
        Client.ServerIP = .Value
        .ValueKey = "ListenPort"
        Client.ClientPort = .Value
        Client.Site = wsListen(pintListen).Tag
    End With
    '----
    ' more ident code
    wsIdentClient(intIndex).Connect
    Do Until wsIdentClient(intIndex).State = sckConnected
        ' wait for the connect
        DoEvents
        ' should have a timeout here ....
    Loop
    wsIdentClient(intIndex).SendData CStr(wsListen(pintListen).LocalPort & " , " & wsServer(intIndex).RemotePort) & vbCrLf
    If intIndex > goIdents.Last Then
        goIdents.Last = intIndex
    End If
    
    Do Until goIdents(intIndex) <> ""
        ' do nothing
        ' should have some kind of timeout here...
        DoEvents
    Loop
    ' now we have our ident
    Client.Ident = CleanString(Trim(Right(goIdents(intIndex), Len(goIdents(intIndex)) - InStrRev(goIdents(intIndex), ":"))))
    WriteToLog "Ident", "Ident for client " & Client.ClientIP & " is " & Client.Ident
    
    '----
    ' now we check to see if they are a valid user
    If Not CheckIP(Client.Ident & "@" & wsClient(intIndex).RemoteHostIP) Then
        GoTo INVALID_USER
    End If
        
    wsServer(intIndex).LocalPort = 0
    wsServer(intIndex).Connect
    Client.LocalPort = wsServer(intIndex).LocalPort
    gOClients.Add Client, CStr(intIndex)
    
    Set oItem = lvUsers.ListItems.Add(, , Now())
    oItem.SubItems(1) = Client.ClientIP
    oItem.SubItems(2) = Client.Site
    oItem.Tag = intIndex
    
    gdblConnections = gdblConnections + 1
    'End If
    
    If IsFormLoaded("frmBounces") Then
        frmBounces.BuildSiteStats
    End If
    
PROC_EXIT:
    Set oItem = Nothing
    Exit Sub
    
INVALID_USER:
    gaintMaxUsers(pintListen) = gaintMaxUsers(pintListen) - 1
    wsClient(intIndex).Close
    GoTo PROC_EXIT
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "ConnectionRequest" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Private Function CleanString(ByVal pstrCleanMe As String) As String
    On Error Resume Next
    If Len(pstrCleanMe) > 4 Then
        If Left(pstrCleanMe, 1) = vbCr Then
            pstrCleanMe = Right(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Left(pstrCleanMe, 1) = vbLf Then
            pstrCleanMe = Right(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Left(pstrCleanMe, 1) = vbCr Then
            pstrCleanMe = Right(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Left(pstrCleanMe, 1) = vbLf Then
            pstrCleanMe = Right(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Right(pstrCleanMe, 1) = vbCr Then
            pstrCleanMe = Left(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Right(pstrCleanMe, 1) = vbLf Then
            pstrCleanMe = Left(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Right(pstrCleanMe, 1) = vbCr Then
            pstrCleanMe = Left(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
        If Right(pstrCleanMe, 1) = vbLf Then
            pstrCleanMe = Left(pstrCleanMe, Len(pstrCleanMe) - 1)
        End If
    End If
    CleanString = pstrCleanMe
End Function

Private Function GetIndex() As Integer
    Dim intIndex As Integer
    Randomize
    intIndex = 0
retry:
    Do Until InStr(1, IndexList, "," & CStr(intIndex) & ",") = 0
        On Error GoTo retry:
        intIndex = Rnd * Rnd / Rnd + Rnd
    Loop
    IndexList = IndexList & CStr(intIndex) & ","
    GetIndex = intIndex
End Function

Public Sub SetupBounces()
    ' creates all the winsock controls for the bounces, and sets up the max users code.
    
    On Error GoTo PROC_ERR
    
    Dim strSiteName As String
    Dim lngCount As Long
    Dim astrBounces() As String
    Dim intLoop As Integer
    Dim objWs As Winsock
    
    WriteToLog "Bounce", "Starting Bounces"
    
    ' close all current listeners
    For Each objWs In wsListen
        If objWs.Index <> 0 Then
            objWs.Close
            Unload objWs
        End If
    Next objWs
    
    ' close all current servers
    For Each objWs In wsServer
        If objWs.Index <> 0 Then
            objWs.Close
            Unload objWs
        End If
    Next objWs
    
    ' close all current clients
    For Each objWs In wsClient
        If objWs.Index <> 0 Then
            objWs.Close
            Unload objWs
        End If
    Next objWs
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces"
        .EnumerateSections astrBounces(), lngCount
        ReDim gaintMaxUsers(1 To lngCount) As Integer
        ReDim gaintMaxUsersCount(1 To lngCount) As Integer
        For intLoop = LBound(astrBounces) To UBound(astrBounces)
            Load wsListen(intLoop)
            .SectionKey = "Software\GH\FTPBounce\Bounces\" & astrBounces(intLoop)
            .ValueKey = "ListenPort"
            wsListen(intLoop).LocalPort = .Value
            wsListen(intLoop).Close
            wsListen(intLoop).Listen
            wsListen(intLoop).Tag = astrBounces(intLoop)
            'lstSiteName.AddItem astrBounces(intLoop)
            WriteToLog "Bounce", "Started Bounce: " & astrBounces(intLoop)
            .ValueKey = "MaxUsers"
            gaintMaxUsersCount(intLoop) = .Value
        Next intLoop
    End With
    
    'BuildSiteStats
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "SetupBounces" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Public Sub UpdateSpyWindows(ByVal iConnection As Long, ByVal strData As String)
    On Error GoTo PROC_ERR
    
    Dim ofrm As Form
    
    For Each ofrm In Forms
        If ofrm.Name = "frmDetails" And CStr(ofrm.Tag) = CStr(iConnection) Then
            ofrm.RefreshInformation
        End If
        If ofrm.Name = "frmSpy" And CStr(ofrm.Tag) = CStr(iConnection) Then
            ofrm.AddToLog strData
        End If
    Next ofrm

PROC_EXIT:
    Set ofrm = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmMain" & vbCrLf & "Procedure: " & "UpdateSpyWindows" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub
