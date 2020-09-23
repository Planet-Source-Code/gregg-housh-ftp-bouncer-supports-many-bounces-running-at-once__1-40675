Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function SetWindowPos _
  Lib "user32" _
  (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) _
  As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

    
Private Const mcHWND_NOTOPMOST = -2
Private Const mcHWND_TOPMOST = -1
Private Const mcSWP_NOSIZE = 1
Private Const mcSWP_NOMOVE = 2

'Global gobjScript As ScriptControl
Public IndexList As String
Public goReg As cRegistry

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public gstrUserID          As String       ' Store the UserID
Public gstrUserPassword    As String       ' Store the Password
Public gbAcceptedId        As Boolean      ' Toggle betwen the accepting of UserID or Password
Public gbSuccessLoging     As Boolean      ' User logged in Successfully
Public gstrUserCommand     As String       ' Command issued by the user

Public gintUL              As Integer      ' Stores the total # of uploads since the bounce was started
Public gintDl              As Integer      ' Stores the total # of downloads since the bounce was started
Public gdblDLB             As Double       ' Stores the total downloaded bytes since the bounce was started
Public gdblConnections     As Double       ' Stores the total # of connections since the bounce was started
Public gOClients           As Collection   ' Collection of currently connected clients
Public gOOldClients        As Collection   ' History of connections, all the mOClients that have disconnected.
Public gOClient            As cClients     ' Generic Client object, used for temp storage of a Client to display data
Public gbShutdown          As Boolean      ' Toggles whether we should shutdown when the last connected client disconnects
Public gaintMaxUsers()     As Integer      ' Stores the # of connected users.
Public gaintMaxUsersCount() As Integer     ' Stores the Max Users setting from bounces.
Public goIdents            As CVector      ' stores the idents returned by the wsIdentClient_DataArrival

' these hold settings
Public gbSiteToggle        As Boolean      ' Stores whether the SITE command is allowed
Public gbDlToggle          As Boolean      ' Stores whether Downloading is allowed
Public gbULToggle          As Boolean      ' Stores whether Uploading is allowed
Public gstrSiteMsg         As String       ' Stores the string to send when the SITE command is used, and it it restricted
Public gstrDlMsg           As String       ' Stores the string to send when the RETR (download) command is used, and it it restricted
Public gstrULMsg           As String       ' Stores the string to send when the STOR (upload) command is used, and it it restricted
Public gstrLogin           As String       ' userid to allow in (for login)
Public gstrPassword        As String       ' userid's password (for login)
Public gvarTelnetPort      As Variant      ' Stores the Port to Listen on for telnet connections
Public gstrIdent           As String       ' Stores the default ident

Public Function InArray(ByRef pastrArray() As String, ByVal pstrValue As String) As Boolean
    
    On Error GoTo PROC_ERR
    
    Dim lngLoop As Long
    Dim bToggle As Boolean
    
    For lngLoop = LBound(pastrArray) To UBound(pastrArray)
        If pastrArray(lngLoop) = pstrValue Then
            bToggle = True
            Exit For
        End If
    Next lngLoop
    
PROC_EXIT:
    InArray = bToggle
    Exit Function
    
PROC_ERR:
    Select Case Err.Number
        Case 9
            Resume PROC_EXIT
        Case Else
            MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modMain" & vbCrLf & "Procedure: " & "InArray2" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
            Resume PROC_EXIT
    End Select
End Function

Public Sub OnTop(pFrm As Form, ByVal pbIsOnTop As Boolean)
    Dim lngState As Long
    
    On Error Resume Next
    
    If pbIsOnTop Then
        lngState = mcHWND_TOPMOST
    Else
        lngState = mcHWND_NOTOPMOST
    End If
    
    SetWindowPos pFrm.hwnd, lngState, 0&, 0&, 0&, 0&, mcSWP_NOSIZE Or mcSWP_NOMOVE
    
End Sub

Public Sub SelectAll(ByRef pobjText As TextBox)
    
    On Error GoTo PROC_ERR
    
    With pobjText
        If Len(.Text) > 0 Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modMain" & vbCrLf & "Procedure: " & "SelectAll" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Public Function IsAlphaNumeric(ByVal sData As String, Optional bAllowSpaces As Boolean = False) As Boolean
    On Error GoTo PROC_ERR

    Dim bToggle As Boolean
    Dim iLoop As Integer
    Dim sMatch As String
    
    If bAllowSpaces Then
        sMatch = "[A-Z,a-z,0-9, ]"
    Else
        sMatch = "[A-Z,a-z,0-9]"
    End If
        
    bToggle = True
    
    If Len(sData) > 0 Then
        For iLoop = 1 To Len(sData)
            If Not Mid(sData, iLoop, 1) Like sMatch Then
                bToggle = False
                Exit For
            End If
        Next iLoop
    Else
        bToggle = False
    End If

PROC_EXIT:
    IsAlphaNumeric = bToggle
    Exit Function
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modMain" & vbCrLf & "Procedure: " & "IsAlphaNumeric" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Function

Public Function IsLoaded(TheControl As Control) As Boolean

    On Error Resume Next
    IsLoaded = (TheControl.Name = TheControl.Name)
End Function

Public Function IsFormLoaded(ByVal strFormName As String) As Boolean

    Dim i As Integer

    For i = 0 To Forms.Count - 1

        If (Forms(i).Name = strFormName) Then
            IsFormLoaded = True
            Exit For
        End If

    Next

End Function

Public Function OpenURL(ByVal sURL As String, ByVal hwnd As Long)
    On Error Resume Next
    Dim lurl As Long
    lurl = ShellExecute(hwnd, vbNullString, sURL, vbNullString, "", SW_SHOWNORMAL)
End Function


