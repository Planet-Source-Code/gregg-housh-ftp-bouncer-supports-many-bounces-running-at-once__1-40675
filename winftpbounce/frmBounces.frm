VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBounces 
   Caption         =   "Bounces"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin MSComctlLib.ListView lvBounces 
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
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
         Key             =   "Site"
         Object.Tag             =   "Site"
         Text            =   "Site"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Connections"
         Object.Tag             =   "Connections"
         Text            =   "Connections"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   120
      Picture         =   "frmBounces.frx":0000
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Bounce Connections "
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
      Left            =   15
      TabIndex        =   0
      Top             =   300
      Width           =   4305
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmBounces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDone As Boolean

Public Sub Display(ByVal Parent As Form)
    On Error GoTo PROC_ERR
    
    BuildSiteStats
    
    Me.Show vbModeless, Parent
    
    Do Until mbDone = True
        DoEvents
    Loop

    mbDone = False

    Unload Me

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounces" & vbCrLf & "Procedure: " & "Display" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Private Sub cmdClose_Click()
    On Error GoTo PROC_ERR
    
    mbDone = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounces" & vbCrLf & "Procedure: " & "cmdClose_Click" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub BuildSiteStats()
    ' Fills the list of connections for each site on the Server Info screen
    
    On Error GoTo PROC_ERR
    
    Dim intLoop As Integer
    Dim lngCount As Long
    Dim astrBounces() As String
    Dim oItem As ListItem
    
    ' site connections counts
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\Bounces"
        .EnumerateSections astrBounces(), lngCount
        lvBounces.ListItems.Clear
        For intLoop = LBound(gaintMaxUsers) To UBound(gaintMaxUsers)
            Set oItem = lvBounces.ListItems.Add(, , astrBounces(intLoop))
            oItem.Tag = intLoop
            oItem.SubItems(1) = gaintMaxUsers(intLoop) & " of " & gaintMaxUsersCount(intLoop)
        Next intLoop
    End With
    
PROC_EXIT:
    Set oItem = Nothing
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "frmBounces" & vbCrLf & "Procedure: " & "BuildSiteStats" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@unix.net with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

