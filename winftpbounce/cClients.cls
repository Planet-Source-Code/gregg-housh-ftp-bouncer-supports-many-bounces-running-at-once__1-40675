VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mlIndex As Long
Private mlDownloads As Double
Private mlUploads As Double
Private mlBytes As Double
Private msClientIP As String
Private msServerIP As String
Private mlClientPort As Long
Private mlServerPort As Long
Private mlLocalPort As Long
Private msStatus As String
Private msCWD As String
Private msUser As String
Private msClientCommands As String
Private msServerCommands As String
Private msLog As String
Private msSite As String

Private m_datTimeStamp As Date
Private m_strIdent As String

Public Sub BytesAdd(ByVal lBytes As Double)
    
    On Error GoTo PROC_ERR
    
    mlBytes = mlBytes + lBytes
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Err.Raise Err.Number
    
End Sub

Public Sub DownloadAdd()
    
    On Error GoTo PROC_ERR
    
    mlDownloads = mlDownloads + 1
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Err.Raise Err.Number
    
End Sub

Public Sub UploadAdd()
    
    On Error GoTo PROC_ERR
    
    mlUploads = mlUploads + 1
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Err.Raise Err.Number
    
End Sub

Public Property Get Bytes() As Double
    
    On Error GoTo PROC_ERR
    
    Bytes = mlBytes
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Bytes(ByVal NewBytes As Double)
    
    On Error GoTo PROC_ERR
    
    mlBytes = NewBytes
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ClientCommands() As String
    
    On Error GoTo PROC_ERR
    
    ClientCommands = msClientCommands
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ClientCommands(ByVal NewClientCommands As String)
    
    On Error GoTo PROC_ERR
    
    msClientCommands = NewClientCommands
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ClientIP() As String
    
    On Error GoTo PROC_ERR
    
    ClientIP = msClientIP
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ClientIP(ByVal NewClientIP As String)
    
    On Error GoTo PROC_ERR
    
    msClientIP = NewClientIP
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ClientPort() As Long
    
    On Error GoTo PROC_ERR
    
    ClientPort = mlClientPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ClientPort(ByVal NewClientPort As Long)
    
    On Error GoTo PROC_ERR
    
    mlClientPort = NewClientPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get CWD() As String
    
    On Error GoTo PROC_ERR
    
    CWD = msCWD
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let CWD(ByVal NewCWD As String)
    
    On Error GoTo PROC_ERR
    
    msCWD = NewCWD
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Downloads() As Double
    
    On Error GoTo PROC_ERR
    
    Downloads = mlDownloads
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Downloads(ByVal NewDownloads As Double)
    
    On Error GoTo PROC_ERR
    
    mlDownloads = NewDownloads
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Ident() As String
    
    On Error GoTo PROC_ERR
    
    Ident = m_strIdent
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Ident(ByVal NewValue As String)
    
    On Error GoTo PROC_ERR
    
    m_strIdent = NewValue
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Index() As Long
    
    On Error GoTo PROC_ERR
    
    Index = mlIndex
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Index(ByVal NewIndex As Long)
    
    On Error GoTo PROC_ERR
    
    mlIndex = NewIndex
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get LocalPort() As Long
    
    On Error GoTo PROC_ERR
    
    LocalPort = mlLocalPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let LocalPort(ByVal NewLocalPort As Long)
    
    On Error GoTo PROC_ERR
    
    mlLocalPort = NewLocalPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Log() As String
    
    On Error GoTo PROC_ERR
    
    Log = msLog
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Log(ByVal NewLog As String)
    
    On Error GoTo PROC_ERR
    
    msLog = NewLog
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ServerCommands() As String
    
    On Error GoTo PROC_ERR
    
    ServerCommands = msServerCommands
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ServerCommands(ByVal NewServerCommands As String)
    
    On Error GoTo PROC_ERR
    
    msServerCommands = NewServerCommands
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ServerIP() As String
    
    On Error GoTo PROC_ERR
    
    ServerIP = msServerIP
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ServerIP(ByVal NewServerIP As String)
    
    On Error GoTo PROC_ERR
    
    msServerIP = NewServerIP
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get ServerPort() As Long
    
    On Error GoTo PROC_ERR
    
    ServerPort = mlServerPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let ServerPort(ByVal NewServerPort As Long)
    
    On Error GoTo PROC_ERR
    
    mlServerPort = NewServerPort
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Site() As String
    
    On Error GoTo PROC_ERR
    
    Site = msSite
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Site(ByVal NewSite As String)
    
    On Error GoTo PROC_ERR
    
    msSite = NewSite
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Status() As String
    
    On Error GoTo PROC_ERR
    
    Status = msStatus
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Status(ByVal NewStatus As String)
    
    On Error GoTo PROC_ERR
    
    msStatus = NewStatus
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get TimeStamp() As Date
    
    On Error GoTo PROC_ERR
    
    TimeStamp = m_datTimeStamp
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let TimeStamp(ByVal NewValue As Date)
    
    On Error GoTo PROC_ERR
    
    m_datTimeStamp = NewValue
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get Uploads() As Double
    
    On Error GoTo PROC_ERR
    
    Uploads = mlUploads
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let Uploads(ByVal NewUploads As Double)
    
    On Error GoTo PROC_ERR
    
    mlUploads = NewUploads
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Get User() As String
    
    On Error GoTo PROC_ERR
    
    User = msUser
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

Public Property Let User(ByVal NewUser As String)
    
    On Error GoTo PROC_ERR
    
    msUser = NewUser
    
PROC_EXIT:
    Exit Property
    
PROC_ERR:
    Err.Raise Err.Number
    
End Property

