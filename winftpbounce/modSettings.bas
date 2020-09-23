Attribute VB_Name = "modSettings"
Option Explicit

Public Sub DefaultSettings()
    ' creates the all the keys and default settings for wBounce in the registry.
    On Error GoTo PROC_ERR
    
    ' first time that wBounce has been loaded
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .ValueType = REG_SZ
        .SectionKey = "Software\GH\FTPBounce"
        If Not .KeyExists Then
            .SectionKey = "Software\GH"
            If Not .KeyExists Then
                .CreateKey
            End If
            .SectionKey = "Software\GH\FTPBounce"
            .CreateKey
        Else
            .DeleteKey
            DoEvents
            .CreateKey
        End If
        .SectionKey = "Software\GH\FTPBounce\AllowedIPs"
        If Not .KeyExists Then
            .CreateKey
            .ValueKey = "IP1"
            .Value = "*@*.*.*.*"
        End If
        .SectionKey = "Software\GH\FTPBounce\Bounces"
        If Not .KeyExists Then
            .CreateKey
            ' create 2 default sites
            .SectionKey = "Software\GH\FTPBounce\Bounces\Site 1"
            .CreateKey
            .ValueKey = "IP"
            .Value = "127.0.0.1"
            .ValueKey = "Port"
            .Value = "21"
            .ValueKey = "ListenPort"
            .Value = "80"
            .ValueKey = "MaxUsers"
            .Value = "10"
            .SectionKey = "Software\GH\FTPBounce\Bounces\Site 2"
            .CreateKey
            .ValueKey = "IP"
            .Value = "127.0.0.1"
            .ValueKey = "Port"
            .Value = "21"
            .ValueKey = "ListenPort"
            .Value = "23"
            .ValueKey = "MaxUsers"
            .Value = "10"
        End If
        
        .SectionKey = "Software\GH\FTPBounce"
        .ValueType = REG_SZ
        .ValueKey = "UserID"
        .Value = "wBounce"
        .ValueKey = "AllowDLMessage"
        .Value = "DOWNLOAD RESTRICTED"
        .ValueKey = "AllowULMessage"
        .Value = "UPLOAD RESTRICTED"
        .ValueKey = "AllowSiteMessage"
        .Value = "SITE COMMAND RESTRICTED"
        .ValueKey = "TelnetPort"
        .Value = "1112"
        .ValueKey = "TelnetUsername"
        .Value = "default"
        .ValueKey = "TelnetPassword"
        .Value = "default"
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "DefaultSettings" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Function GetIps(ByRef pastrIPArray() As String, ByRef plngCount As Long)
    
    On Error GoTo PROC_ERR
    
    Dim intKey As Integer
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\AllowedIps"
        .EnumerateValues pastrIPArray(), plngCount
        For intKey = 1 To plngCount
            .ValueKey = pastrIPArray(intKey)
            pastrIPArray(intKey) = .Value
        Next intKey
    End With
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "GetIps" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Function

Public Function CheckIP(ByVal pstrIPCheck As String) As Boolean
    
    On Error GoTo PROC_ERR
    
    Dim bToggle As Boolean
    Dim astrKeys() As String
    Dim lngKeyCount As Long
    Dim intKey As Integer
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\AllowedIPs"
        .EnumerateValues astrKeys(), lngKeyCount
        For intKey = 1 To lngKeyCount
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\GH\FTPBounce\AllowedIPs"
            .ValueKey = astrKeys(intKey)
            If (pstrIPCheck Like .Value) Then
                bToggle = True
                Exit For
            End If
        Next intKey
    End With
    
    CheckIP = bToggle
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "CheckIP" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Function

Public Sub AddIP(ByVal pstrIPAdd As String)
    ' adds an ip to the list of allowed ips in the registry
    
    On Error GoTo PROC_ERR
    
    Dim astrIPArray() As String
    Dim lngCount As Long
    Dim strIPFind As String
    Dim strChar As String
    Dim strLastBlank As String
    Dim lngI As Long
    
    GetIps astrIPArray, lngCount
    
    If lngCount > 0 Then
        astrIPArray = QuickSort(astrIPArray)
    End If
    
    If Not InArray(astrIPArray(), pstrIPAdd) Then
    
        strLastBlank = ""
        With goReg
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\GH\FTPBounce\AllowedIps"
            If Not .KeyExists Then
                .CreateKey
            End If
            .ValueType = REG_SZ
        End With
        For lngI = 1 To lngCount
            With goReg
                .ValueKey = astrIPArray(lngI)
                strIPFind = .Value
            End With
            If strIPFind = "" Then
                strLastBlank = astrIPArray(lngI)
                Exit For
            End If
        Next lngI
        
        If strLastBlank <> "" Then
            goReg.ValueKey = strLastBlank
            goReg.Value = pstrIPAdd
        Else
            strChar = CStr(lngCount + 1)
            goReg.ValueKey = "IP" & strChar
            goReg.Value = pstrIPAdd
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "AddIP" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

Public Sub RemoveIP(ByVal pstrIPRemove As String)
    
    On Error GoTo PROC_ERR
    
    Dim astrIPArray() As String
    Dim lngCount As Long
    Dim intI As Integer
    
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce\AllowedIps"
        .EnumerateValues astrIPArray, lngCount
        
        For intI = 1 To lngCount
            .ValueKey = astrIPArray(intI)
            If .Value = pstrIPRemove Then
                .Value = ""
                Exit For
            End If
        Next intI
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "RemoveIP" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
    
End Sub

Public Sub gRefreshSettings()
    
    On Error GoTo PROC_ERR
        
    With goReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\GH\FTPBounce"
        .ValueKey = "UserID"
        gstrIdent = .Value
        .ValueKey = "AllowDL"
        gbDlToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowUL"
        gbULToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowSite"
        gbSiteToggle = IIf(.Value = "yes", True, False)
        .ValueKey = "AllowDLMessage"
        gstrDlMsg = .Value
        .ValueKey = "AllowULMessage"
        gstrULMsg = .Value
        .ValueKey = "AllowSiteMessage"
        gstrSiteMsg = .Value
        .ValueKey = "TelnetPort"
        gvarTelnetPort = .Value
        .ValueKey = "TelnetUsername"
        gstrLogin = .Value
        .ValueKey = "TelnetPassword"
        gstrPassword = .Value
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "An Error Has Occured." & vbCrLf & "Module: " & "modSettings" & vbCrLf & "Procedure: " & "gRefreshSettings" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & vbCrLf & "If this error continues, and/or" & vbCrLf & "causes the application to end unexpectedly." & vbCrLf & "Email gregg@sc.am with" & vbCrLf & "this information, and a detailed" & vbCrLf & "description of the events leading to it.", vbCritical, "Error"
    Resume PROC_EXIT
End Sub

