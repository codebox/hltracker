Attribute VB_Name = "modMain"
Option Explicit
Option Base 1

Public Const APP_NAME = "Codebox Hotline Tracker"
Public Const SHORT_APP_NAME = "CodeboxTracker"
Public Const ERR_BASE = vbObjectError + 1234
Public Const TCP_LISTEN = 5498
Public Const UDP_LISTEN = 5499
Public Const MAX_TOOLTIP As Integer = 23 'Len(APP_NAME)
Private Const TXTBOX_MAXLEN = 30000
Private Const TXTBOX_OKLEN = 29000
Public Const DISABLED_COL = &H80000013

Public Const SEP_CHAR = "|"

Public Const RELEASE_BUILD = True

Private Const MAX_SERVER_STRING_LEN = 255
Private Const MAX_USER_COUNT = 65535
Private Const MAX_PORT = 65535
Private Const DEF_IP = "127.0.0.2"
Private Const DEF_PORT = 5500
Private Const BLOCK_MSG_PORT = DEF_PORT
Private Const BLOCK_MSG_USERS = 31337

Public Const FILTER_TYPE_NAME = "1"
Public Const FILTER_TYPE_DESC = "2"
Public Const FILTER_TYPE_PORT = "3"
Public Const FILTER_TYPE_IP = "4"
Public Const FILTER_TYPE_USERS = "5"

Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * MAX_TOOLTIP
End Type
' IDEA - improve HTML reports for larger listings
Public Const NIM_ADD = &H0
Public Const NIM_MOD = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Private Const LIST_REQ_MAGIC_NUM = "HTRK"
Private Const VERSION_REQ_MAGIC_NUM = "HTVR"
'Private Const MAX_BATCH_SIZE = 65535
Private Const MAX_BATCH_SIZE = 30000

'Private mcolServers As Collection
Public mobjServers As clsServers
Public mobjLog As clsLog
Public mlListingCounter As Long
Public mobjSettings As clsSettings
Public mdtStartedAt As Date
Public mobjFilters As clsFilterRules
Public mobjBlockedIPs As clsIPRanges

Public Enum StatusMsgType
    ServerAdded = 1
    ServerExpired = 2
    ListingRequest = 3
    MirrorInfo = 4
    ServerFiltered = 5
    BadPassword = 6
End Enum

Private mlNextID As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function GetUID() As String
    Static lCounter As Long
    lCounter = lCounter + 1
    GetUID = "ID" & CStr(lCounter)
End Function

Public Function BuildFakeServer(sServer As String) As clsServer
    Dim arrServer() As String
    Dim sServerName As String
    Dim sServerDesc As String
    Dim objIPAddress As clsIP
    Dim sIP As String
    Dim bSettingIP As Boolean
    Dim lUserCount As Long
    Dim objServer As clsServer
    Dim lPort As Long

    On Error GoTo errhandler
    arrServer = Split(sServer, Chr(1), 5)
    sServerName = arrServer(0)
    sServerDesc = arrServer(4)
    
    If Len(sServerName) > MAX_SERVER_STRING_LEN Then sServerName = Left$(sServerName, MAX_SERVER_STRING_LEN)
    If Len(sServerDesc) > MAX_SERVER_STRING_LEN Then sServerDesc = Left$(sServerDesc, MAX_SERVER_STRING_LEN)
    bSettingIP = True
    Set objIPAddress = New clsIP
    objIPAddress.BuildFromString (arrServer(1))
    sIP = objIPAddress.Describe
    
AfterIP:
    Set objIPAddress = Nothing
    bSettingIP = False
    
    lUserCount = arrServer(3)
    If lUserCount > MAX_USER_COUNT Then lUserCount = MAX_USER_COUNT
    
    lPort = arrServer(2)
    If lPort > MAX_PORT Or lPort < 1 Then lPort = DEF_PORT
    
    Set objServer = BuildNewServerObject
    objServer.Description = sServerDesc
    objServer.ServerType = Fake
    objServer.IPString = sIP
    objServer.LastCheckedIn = Now
    objServer.ServerName = sServerName
    objServer.Port = lPort
    objServer.UserCount = lUserCount
    
    Set BuildFakeServer = objServer
    Set objServer = Nothing
    
    Exit Function
errhandler:
    If bSettingIP Then
        sIP = DEF_IP
        Resume AfterIP
    Else
        Set BuildFakeServer = Nothing
    End If
    
End Function

Public Sub logit(sText As String, Optional nID = 0)
    If nID = 3 Then
     'Print #1, sText
     'Debug.Print sText
    End If
End Sub

Public Function CloneCollection(colThis As Collection) As Collection
    Dim objMirror As clsMirroredTracker
    Dim colNew As Collection
    Dim objClone As clsMirroredTracker
    
    Set colNew = New Collection
    
    For Each objMirror In colThis
        Set objClone = objMirror.Clone
        colNew.Add objClone, objClone.Address
    Next objMirror
    
    Set objClone = Nothing
    Set CloneCollection = colNew
    Set colNew = Nothing
End Function

Public Function CloneCollection2(colThis As clsFilterRules) As clsFilterRules
    Dim objFilter As IFilterRule
    Dim objNew As clsFilterRules
    Dim objClone As IFilterRule
    
    On Error GoTo errhandler
    
    Set objNew = New clsFilterRules
    
    For Each objFilter In colThis
        Set objClone = objFilter.Clone
        'Debug.Print objClone.Description
        objNew.Add objClone
        Set objClone = Nothing
    Next objFilter
    
    Set objClone = Nothing
    Set CloneCollection2 = objNew
    Set objNew = Nothing
    Exit Function
errhandler:
    
End Function

Public Function CloneCollection3(objServers As clsServers) As clsServers
    Dim objServer As clsServer
    Dim objNewServers As clsServers
    Dim objClone As clsServer
    
    Set objNewServers = New clsServers
    
    For Each objServer In objServers
        Set objClone = objServer.Clone
        objNewServers.AddOrUpdate objClone
    Next objServer
    
    Set objClone = Nothing
    Set CloneCollection3 = objNewServers
    Set objNewServers = Nothing
End Function

Public Function CloneCollection4(objIPRanges As clsIPRanges) As clsIPRanges
    Dim objIPRange As clsIPRange
    Dim objNewIPRanges As clsIPRanges
    Dim objClone As clsIPRange
    
    Set objNewIPRanges = New clsIPRanges
    
    For Each objIPRange In objIPRanges
        Set objClone = objIPRange.Clone
        objNewIPRanges.Add objClone
    Next objIPRange
    
    Set objClone = Nothing
    Set CloneCollection4 = objNewIPRanges
    Set objNewIPRanges = Nothing
End Function

' ########################################
' TODO List
'
' * Test against bad TCP/UDP requests (overflows, ctrl chars etc)
' * Encrypt password in registry
' ########################################

Public Sub ErrorReport(lErrNumber As Long, sErrSourceMod As String, sErrSourceProc As String, sErrDescription As String, Optional lErrLine As Long = 0)
    Dim sErrText As String
    
    sErrText = "Error [" & CStr(lErrNumber) & "] " & vbCrLf & sErrDescription & vbCrLf
    sErrText = sErrText & "happened in " & sErrSourceMod & "." & sErrSourceProc & IIf(Right$(sErrSourceProc, 2) <> "()", "()", "")
    If lErrLine > 0 Then
        sErrText = sErrText & vbCrLf & "on line " & CStr(lErrLine)
    End If
    If Not RELEASE_BUILD Then MsgBox sErrText, vbExclamation, "Error"
    
End Sub

Public Function GetUpTime() As String
    Dim lDaysUp As Long
    Dim lHoursUp As Long
    Dim lMinsUp As Long
    Dim lSecsUp As Long
    Dim lTotalSecsUp As Long
    
'    lDaysUp = DateDiff("d", mdtStartedAt, Now)
'    lHoursUp = DateDiff("h", mdtStartedAt, Now) Mod 24
'    lMinsUp = DateDiff("n", mdtStartedAt, Now) Mod 60
'    lSecsUp = DateDiff("s", mdtStartedAt, Now) Mod 60
    
    lTotalSecsUp = DateDiff("s", mdtStartedAt, Now)
    lSecsUp = lTotalSecsUp Mod 60
    lMinsUp = lTotalSecsUp \ 60
    If lMinsUp > 60 Then
        lMinsUp = lMinsUp Mod 60
        lHoursUp = lTotalSecsUp \ (60 * 60)
        If lHoursUp > 24 Then
            lHoursUp = lHoursUp Mod 24
            lDaysUp = lTotalSecsUp \ (60& * 60 * 24)
        End If
    End If
    
    GetUpTime = CStr(lDaysUp) & "d " & CStr(lHoursUp) & "h " & CStr(lMinsUp) & "m " & CStr(lSecsUp) & "s"

End Function
Public Sub UpdateServerCounter()
    frmMain.lblServerCount.Caption = mobjServers.Count
End Sub

Public Sub UpdateUserCounter()
    frmMain.lblUserCount.Caption = mobjServers.UserCount
End Sub

Public Function BuildNewServerObject() As clsServer
    Dim objServer As clsServer
    
    Set objServer = New clsServer
    Set BuildNewServerObject = objServer
    Set objServer = Nothing
    
End Function

Public Sub DoAlert()
    frmMain.FlashForm
    Beep
End Sub

Public Function IsIPValid(nIP1, nIP2, nIP3, nIP4) As Boolean
    Dim nRule As Integer
    IsIPValid = True
    If nIP1 = 0 Or nIP1 = 127 Or nIP1 = 10 Or nIP1 >= 224 Then
        IsIPValid = False
        nRule = 1
    ElseIf (nIP1 = 192) And (nIP2 = 168) Then
        IsIPValid = False
        nRule = 2
    ElseIf (nIP1 = 172) And (nIP2 >= 16) And (nIP2 < 32) Then
        IsIPValid = False
        nRule = 3
    End If
End Function

Public Sub AddServer(objServer As clsServer)
    If mobjServers Is Nothing Then Set mobjServers = New clsServers
    mobjServers.AddOrUpdate objServer
    mobjLog.NewServerInfo objServer
End Sub

Public Function BuildServerInfo(objBytes As clsByteArray, sServerIP As String) As clsServer
    Dim objServer As clsServer
    Dim nNameSize As Integer
    Dim nDescSize As Integer
    Dim arrIP() As String
    Dim sPassword As String
    Dim nPasswordSize As Integer
    Dim nNextBytePointer As Long
    Dim bCheckingForPass As Boolean
    
    On Error GoTo errhandler
    Set objServer = BuildNewServerObject
    
    arrIP = Split(sServerIP, ".")
    objServer.SetIP CByte(arrIP(0)), CByte(arrIP(1)), CByte(arrIP(2)), CByte(arrIP(3))
    
    If objBytes.GetInteger(1) <> 1 Then Err.Raise ERR_BASE, , "Bad initial bytes"
    objServer.Port = objBytes.GetInteger(3)
    objServer.UserCount = objBytes.GetInteger(5)
    
    nNameSize = objBytes.GetByte(13)
    objServer.ServerName = objBytes.GetString(14, CLng(nNameSize))
    nNextBytePointer = 14 + nNameSize
    
    nDescSize = objBytes.GetByte(nNextBytePointer)
    objServer.Description = objBytes.GetString(nNextBytePointer + 1, CLng(nDescSize))
    nNextBytePointer = nNextBytePointer + 1 + nDescSize
    
    bCheckingForPass = True
    nPasswordSize = objBytes.GetByte(nNextBytePointer)
    sPassword = objBytes.GetString(nNextBytePointer + 1, CLng(nPasswordSize))
    bCheckingForPass = False
    
AfterPass:
    If VerifyPassword(sPassword) Or Not mobjSettings.RequirePassword Then
        Set BuildServerInfo = objServer
    Else
        StatusMsg "Bad password from server '" & objServer.ServerName & "' at " & objServer.IPString, BadPassword
        Set BuildServerInfo = Nothing
    End If
    
    Set objServer = Nothing
    
    Exit Function
errhandler:
    If bCheckingForPass Then Resume AfterPass
    'StatusMsg "ERROR (BuildServerInfo):" & Err.Description
    Set BuildServerInfo = Nothing
End Function

Private Function VerifyPassword(sPass As String) As Boolean
    VerifyPassword = (mobjSettings.Password = sPass)
End Function

Public Function VerifyVersionRequest(objBytes As clsByteArray) As Boolean
    On Error GoTo errhandler
    
    VerifyVersionRequest = (objBytes.GetString(1, 4) = VERSION_REQ_MAGIC_NUM)
    
    Exit Function
errhandler:
    VerifyVersionRequest = False

End Function

Public Function VerifyListingRequest(objBytes As clsByteArray) As Boolean
    Dim sMagic As String
    Dim nVersion As Integer
    Dim nLoginSize As Integer
    Dim sLogin As String
    Dim nPassSize As Integer
    Dim sPass As String
    
    On Error GoTo errhandler
    
    sMagic = objBytes.GetString(1, 4)
    If sMagic <> LIST_REQ_MAGIC_NUM Then
        Err.Raise ERR_BASE, , "Bad magic number in listing request: " & sMagic
    End If
    nVersion = objBytes.GetInteger(5)
    
    If nVersion > 1 Then
        nLoginSize = objBytes.GetByte(7)
        sLogin = objBytes.GetString(8, CLng(nLoginSize))
        nPassSize = objBytes.GetByte(nLoginSize + 8 + 1)
        sPass = objBytes.GetString(nLoginSize + 8 + 2, CLng(nPassSize))
    End If
    
    VerifyListingRequest = True
    
    Exit Function
errhandler:
    'StatusMsg "ERROR (VerifyListingRequest):" & Err.Description
    VerifyListingRequest = False
End Function

Public Sub StatusMsg(sText As String, enmMsgType As StatusMsgType)
    Dim nAsciiCode As Integer
    Dim bShowMsg As Boolean
    Dim sMsg As String
    
'    For nAsciiCode = 0 To 31
'        sText = Replace(sText, Asc(nAsciiCode), ".")
'    Next nAsciiCode
'    For nAsciiCode = 127 To 255
'        sText = Replace(sText, Asc(nAsciiCode), ".")
'    Next nAsciiCode
    
    bShowMsg = ((enmMsgType = BadPassword) And mobjSettings.ShowLogBadPass) Or _
                ((enmMsgType = ListingRequest) And mobjSettings.ShowLogListReq) Or _
                ((enmMsgType = MirrorInfo) And mobjSettings.ShowLogMirrors) Or _
                ((enmMsgType = ServerAdded) And mobjSettings.ShowLogServerAdd) Or _
                ((enmMsgType = ServerExpired) And mobjSettings.ShowLogServerExpire) Or _
                ((enmMsgType = ServerFiltered) And mobjSettings.ShowLogServerFiltered)
                
    If bShowMsg Then
        If Len(frmMain.txtData.Text) > TXTBOX_MAXLEN Then
            frmMain.txtData.Text = Right$(frmMain.txtData.Text, TXTBOX_OKLEN)
        End If
        
        frmMain.txtData.Text = frmMain.txtData.Text & IIf(Len(frmMain.txtData.Text) > 0, vbCrLf, "")
        
        If mobjSettings.ShowLogTimestamp Then
            sMsg = Format$(Now, "dd/mm/yyyy hh:mm:ss") & " "
        End If
        
        sMsg = sMsg & sText
        frmMain.txtData.Text = frmMain.txtData.Text & sMsg
        frmMain.txtData.SelStart = Len(frmMain.txtData.Text)
        mobjLog.NewDisplayLogEntry sMsg
    End If
    
End Sub

Public Sub SendListing(sckWinsock As Winsock, bIsBlocked As Boolean)
    Dim objBytes As clsByteArray
    
    If Not bIsBlocked Then
        Set objBytes = BuildList
    Else
        Set objBytes = BuildBlockMsg
    End If
    
    'mobjListingManager.AddNew sckWinsock, objBytes
    sckWinsock.SendData objBytes.GetBytes
    'sckWinsock.Close
    'Unload sckWinsock
    Set objBytes = Nothing
End Sub

Public Sub SendVersion(sckWinsock As Winsock)
    Dim objBytes As clsByteArray
    
    Set objBytes = New clsByteArray
    objBytes.AppendString APP_NAME & " v" & App.Major & "." & App.Minor & "." & App.Revision & " by rob"
    sckWinsock.SendData objBytes.GetBytes
    'sckWinsock.Close
    'Unload sckWinsock
    Set objBytes = Nothing
End Sub

Private Function BuildBlockMsg() As clsByteArray
    Dim objServer As clsServer
    Dim objServers As clsServers
    
    Set objServers = New clsServers
    Set objServer = New clsServer
    objServer.Description = mobjSettings.BlockedMsg
    objServer.Port = BLOCK_MSG_PORT
    objServer.UserCount = BLOCK_MSG_USERS
    objServer.SetIP 127, 0, 0, 1
    objServer.ServerType = Fake
    
    objServers.AddOrUpdate objServer
    Set objServer = Nothing
    Set BuildBlockMsg = BuildList(objServers)
    Set objServers = Nothing
End Function

Private Function BuildList(Optional objServers As clsServers = Nothing) As clsByteArray
    Dim objBytes As clsByteArray
    Dim objServerList As clsByteArray
    Dim objServer As clsServer
    Dim objThisServersBytes As clsByteArray
    Dim X As Integer
    Set objBytes = New clsByteArray
    
    If objServers Is Nothing Then Set objServers = mobjServers
    
    objBytes.AppendString LIST_REQ_MAGIC_NUM
    objBytes.AppendNumber 1, 2
    
    Set objServerList = New clsByteArray
    
    For Each objServer In objServers
        If Not objServer.Filtered Then
            Set objThisServersBytes = objServer.BuildListingEntry
            X = X + 1
            If (objServerList.Size + objThisServersBytes.Size) > MAX_BATCH_SIZE Then
                objBytes.AppendNumber 1, 2
                objBytes.AppendNumber objServerList.Size + 4, 2 'Remaining size of listing
                objBytes.AppendNumber objServers.Count, 2
                objBytes.AppendNumber objServers.Count, 2 'twice
                If objServerList.Size > 0 Then
                    objBytes.AppendByteObject objServerList
                End If
                objServerList.Clear
            End If
            objServerList.AppendByteObject objThisServersBytes
        End If
    Next objServer
    
    If objServerList.Size > 0 Then
        objBytes.AppendNumber 1, 2
        objBytes.AppendNumber objServerList.Size + 4, 2 'Remaining size of listing
        objBytes.AppendNumber objServers.Count, 2
        objBytes.AppendNumber objServers.Count, 2 'twice
        objBytes.AppendByteObject objServerList
    End If
    
    Set BuildList = objBytes
    Set objServerList = Nothing
    Set objBytes = Nothing
End Function
