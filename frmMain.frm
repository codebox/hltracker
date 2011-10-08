VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000013&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Codebox Tracker"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckTracker 
      Index           =   0
      Left            =   5625
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2940
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   "blue"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02A4
            Key             =   "hand"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0406
            Key             =   "red"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrAlert 
      Left            =   4980
      Top             =   300
   End
   Begin VB.Timer tmrCheck 
      Left            =   4200
      Top             =   240
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2115
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1020
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2235
      Begin VB.CommandButton cmdHide 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         ToolTipText     =   "Send to System Tray"
         Top             =   640
         Width           =   315
      End
      Begin VB.CommandButton cmd2 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "Settings..."
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton cmdShowData 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   1800
         MaskColor       =   &H8000000F&
         TabIndex        =   8
         ToolTipText     =   "Generate Report"
         Top             =   150
         Width           =   315
      End
      Begin VB.Label lblUserCount 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1140
         TabIndex        =   7
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Users:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Listings:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblListingCount 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1140
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Servers:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
      Begin VB.Label lblServerCount 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1140
         TabIndex        =   1
         Top             =   420
         Width           =   555
      End
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   480
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Index           =   0
      Left            =   360
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private mlTCPConnMax As Long
Private mlUDPConnMax As Long
Private mnLastSettingsFrame As Integer

Private Const CHECK_INTERVAL = 1
Private Const MILLISEC_PER_MINUTE = 60000
Private Const MIN_WIDTH = 2445
Private Const MIN_HEIGHT = 1350
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONUP = &H202
Private Const REPORT_TITLE = "Tracker Index"
Private Const MAX_ICON_FLASH = 5
Private Const MAX_SOCKET_AGE = 5

Private Const MOD_NAME = "frmMain"

Private mlHorizDiff As Long
Private mlVertDiff As Long
Private mbInSystemTray As Boolean
Private mnfIconData As NOTIFYICONDATA
Private mlNextTrackerSockIndex As Long
Private mcolMirrorTrackers As Collection

' IDEA
' configure alerts (audio/wav/flash etc)
' watch log sizes and trim


Private Sub AddFakeServers()
    Dim dctServers As Dictionary
    Dim objTest  As clsServer
    Dim arrServers() As String
    Dim vntServer As Variant
    
    arrServers = Split(mobjSettings.FakeServers, Chr(0))
    For Each vntServer In arrServers
        Set objTest = BuildFakeServer(CStr(vntServer))
        If Not objTest Is Nothing Then
            AddServer objTest
        End If
    Next vntServer
    
    Set objTest = Nothing
End Sub

Private Sub SendMeToTray()
    If Not mbInSystemTray Then
        Shell_NotifyIcon NIM_ADD, mnfIconData
        Me.Hide
        'mnuShow.Caption = "Show"
        mbInSystemTray = True
    End If
End Sub

Private Sub RestoreMeToDesktop()
    If mbInSystemTray Then
        Shell_NotifyIcon NIM_DELETE, mnfIconData
        Me.Show
        Me.WindowState = vbNormal
        txtData.SelStart = Len(txtData.Text)
        mbInSystemTray = False
    End If
    
End Sub

Private Sub cmd2_Click()
    On Error GoTo errhandler
     frmSettings.InitialFrame = mnLastSettingsFrame
     frmSettings.Show vbModal
     
     Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmd2_Click", Err.Description, Erl
End Sub

Private Sub cmdHide_Click()
    On Error GoTo errhandler

    SendMeToTray
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdHide_Click", Err.Description, Erl
End Sub

Private Sub cmdShowData_Click()
    Dim sHTML As String
    Dim objServer As clsServer
    Dim sData As String
    Dim objFileSystem As FileSystemObject
    Dim objTextStream As TextStream
    Dim sFilePath As String
    
    On Error GoTo errhandler
    
    sHTML = "<HTML><HEAD><TITLE>" & REPORT_TITLE & "</TITLE></HEAD><BODY><FONT FACE='arial'><CENTER><B><U>" & REPORT_TITLE & "</U></B>"
    sHTML = sHTML & "<BR><BR></CENTER>Report generated at <B>" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "</B><BR><BR>"
    
    If mobjServers.Count = 0 Then
        sData = "- No servers currently active -<BR><BR>"
    Else
        sData = "<TABLE border=1 align=center cellpadding=2><TR bgcolor=C0C0C0><TD><B>Name</B></TD><TD><B>Description</B></TD><TD><B>IP Address</B></TD><TD><B>Port</B></TD><TD><B>User Count</B></TD><TD><B>Last Check-In Time</B></TD></TR>"
        For Each objServer In mobjServers
            sData = sData & "<TR><TD>" & objServer.ServerName & "</TD><TD>" & objServer.Description & "</TD><TD>" & objServer.IPString & "</TD><TD>" & objServer.Port & "</TD><TD>" & objServer.UserCount & "</TD><TD>" & objServer.LastCheckedIn & "</TD></TR>"
            'sData = sData & "<TR><TD><a href='hotline://" & objServer.IPString & ":" & objServer.Port & "'>" & objServer.Name & "</a></TD><TD>" & objServer.Description & "</TD><TD>" & objServer.IPString & "</TD><TD>" & objServer.Port & "</TD><TD>" & objServer.UserCount & "</TD><TD>" & objServer.LastCheckedIn & "</TD></TR>"
        Next objServer
        sData = sData & "</TABLE>"
    End If
    sHTML = sHTML & sData
    sHTML = sHTML & "<BR><BR>"
    sHTML = sHTML & "<TABLE cellpadding=5>"
    sHTML = sHTML & MakeStatsRow("Listings", CStr(mlListingCounter))
    sHTML = sHTML & MakeStatsRow("Servers", mobjServers.Count)
    sHTML = sHTML & MakeStatsRow("Users", mobjServers.UserCount)
    sHTML = sHTML & MakeStatsRow("Up Time", GetUpTime)
    sHTML = sHTML & "</TABLE>"
    sHTML = sHTML & "</FONT></HTML>"
    
    Set objFileSystem = New FileSystemObject
    sFilePath = App.Path & "\report.html"
    Set objTextStream = objFileSystem.CreateTextFile(sFilePath, True)
    objTextStream.Write sHTML
    objTextStream.Close
    Set objTextStream = Nothing
    Set objFileSystem = Nothing
    
    LaunchDoc Me.hwnd, sFilePath
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdShowData_Click", Err.Description, Erl
End Sub

Private Function MakeStatsRow(sName As String, sNumber As String) As String
    MakeStatsRow = "<TR><TD BGCOLOR=C0C0C0><B>" & sName & ":</B></TD><TD BGCOLOR=E0E0E0>" & sNumber & "</TD></TR>"
End Function

Private Sub Form_Load()
    On Error GoTo errhandler

    ScaleMode = 3
    mdtStartedAt = Now
    Set mobjServers = New clsServers
    mobjServers.SendStatusMsgs = True
    Set mobjSettings = New clsSettings
    Set mobjLog = New clsLog
    Set mcolMirrorTrackers = New Collection
    
    Set mobjFilters = mobjSettings.Filters
    Set mobjBlockedIPs = New clsIPRanges
    mobjBlockedIPs.BuildFromRegString mobjSettings.BlockedIPs
    
    'Open App.Path & "\debug.log" For Append As #1

    Height = mobjSettings.FormHeight
    Width = mobjSettings.FormWidth
    txtData.ForeColor = mobjSettings.TextColour
    txtData.BackColor = mobjSettings.BackColour
    
    tmrCheck.Interval = CHECK_INTERVAL * MILLISEC_PER_MINUTE
    
    sckTCP(0).Protocol = sckTCPProtocol
    sckTCP(0).LocalPort = TCP_LISTEN
    sckTCP(0).Listen

    sckUDP.Protocol = sckUDPProtocol
    sckUDP.Bind UDP_LISTEN
        
    mlTCPConnMax = 0
    mlUDPConnMax = 0
    
    AddFakeServers
    UpdateUserCounter
    
    mlHorizDiff = 7 'ScaleWidth - txtData.Width
    mlVertDiff = 69 'ScaleHeight - txtData.Height
    
    With mnfIconData
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = APP_NAME
        .cbSize = Len(mnfIconData)
    End With
    
    SetWindowPos Me.hwnd, IIf(mobjSettings.Float, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ReCalculateCounts
    mnLastSettingsFrame = 1
    
    If mobjSettings.DoMirroring Then
        DoMirror
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "Form_Load", Err.Description, Erl
    
End Sub
' IDEA - include allowed/denied in hit log
Public Sub ReCalculateCounts()
    UpdateUserCounter
    UpdateServerCounter
End Sub

Public Sub CallbackSetInitFrame(nFrame As Integer)
    mnLastSettingsFrame = nFrame
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue  As Long
    ' Callback fn for mousemove over icon in systray
    On Error Resume Next
    If mbInSystemTray Then
        If X = WM_LBUTTONUP Then
            RestoreMeToDesktop
        Else
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errhandler

    If WindowState = vbNormal Then
        mobjSettings.FormHeight = Height
        mobjSettings.FormWidth = Width
    End If
    
    Set mobjLog = Nothing
    Set mcolMirrorTrackers = Nothing
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "Form_QueryUnload", Err.Description, Erl
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = vbNormal Then
        If Width < MIN_WIDTH Then Width = MIN_WIDTH
        If Height < MIN_HEIGHT Then Height = MIN_HEIGHT
        txtData.Width = ScaleWidth - mlHorizDiff
        txtData.Height = ScaleHeight - mlVertDiff
        frmMain.txtData.SelStart = Len(frmMain.txtData.Text)
    End If
End Sub

Public Sub CallBackGotServerCount(sTracker As String, lCount As Long)
    StatusMsg "... retrieving " & CStr(lCount) & IIf(lCount = 1, " server ", " servers ") & "from '" & sTracker & "' ...", MirrorInfo
End Sub

Public Sub CallBackListingComplete(sTracker As String, sckControl As MSWinsockLib.Winsock)
    StatusMsg "... mirror of '" & sTracker & "' complete.", MirrorInfo
    DestroyMirrorSocket sckControl
End Sub

Public Sub CallBackServerData(colServers As Collection)
    Dim objServer As clsServer
    For Each objServer In colServers
        AddServer objServer
    Next objServer
End Sub

Private Sub sckTCP_Close(Index As Integer)
    On Error GoTo errhandler

    If Index > 0 Then
        'Debug.Print "CLOSING sckTCP(" & CStr(Index) & ")"
        sckTCP(Index).Close
        Unload sckTCP(Index)
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTCP_Close", Err.Description, Erl
End Sub

Private Sub sckTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error GoTo errhandler

    If Index = 0 Then
        If Not mobjBlockedIPs.AddressIsBlocked(sckTCP(Index).RemoteHostIP) Or Not mobjSettings.RefuseBlocked Then
          mlTCPConnMax = mlTCPConnMax + 1
          Load sckTCP(mlTCPConnMax)
        ' Store the time the socket was created so it can be cleaned up if it hangs around too long
          sckTCP(mlTCPConnMax).Tag = Now
          sckTCP(mlTCPConnMax).LocalPort = 0
          sckTCP(mlTCPConnMax).Accept requestID
        Else
            'refuse
        End If
    End If
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTCP_ConnectionRequest", Err.Description, Erl
End Sub

Private Sub sckTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim objBytes As clsByteArray
    Dim arrListingData() As Byte
    Dim sData As String
    Dim bValidRequest As Boolean
    Dim bIsBlocked As Boolean
    On Error GoTo errhandler
    
    Set objBytes = New clsByteArray
    'sckTCP(Index).GetData arrData, vbArray + vbByte
1   sckTCP(Index).GetData sData, vbString
    
    'objBytes.AppendBytes arrData
2   objBytes.AppendString sData
    
3   If VerifyVersionRequest(objBytes) Then
4       SendVersion sckTCP(Index)
    Else
5       bValidRequest = VerifyListingRequest(objBytes)
        
        If bValidRequest Then
            If Not mobjBlockedIPs.AddressIsBlocked(sckTCP(Index).RemoteHostIP) Then
6               StatusMsg "Listing request from: " & sckTCP(Index).RemoteHostIP, ListingRequest
                bIsBlocked = False
            Else
                StatusMsg "Blocked listing request from: " & sckTCP(Index).RemoteHostIP, ListingRequest
                bIsBlocked = True
            End If
7           SendListing sckTCP(Index), bIsBlocked
8           mlListingCounter = mlListingCounter + 1
9           UpdateListingCounter
10          mobjLog.NewHitInfo sckTCP(Index).RemoteHostIP
11          If mobjSettings.AlertList Then DoAlert
        End If
    End If
    
    Set objBytes = Nothing
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTCP_DataArrival", Err.Description, Erl
End Sub

Private Sub sckTCP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Debug.Print Now & "WINSOCK ERROR " & CStr(Index) & " " & Description
End Sub

Private Sub sckTCP_SendComplete(Index As Integer)
    ' TOD - use this to break listing into chunks
    'mobjListingManager.SendNextBatch sckTCP(Index)
   ' sckTCP(Index).Close
    'Unload sckTCP(Index)
End Sub

Private Sub sckTracker_Connect(Index As Integer)
    On Error GoTo errhandler

    sckTracker(Index).SendData "HTRK" & Chr(0) & Chr(1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTracker_Connect", Err.Description, Erl
    
End Sub

Private Sub sckTracker_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim arrdata() As Byte
    Dim objBytes As clsByteArray
    
    On Error GoTo errhandler

1   sckTracker(Index).GetData arrdata, vbByte + vbArray
    'ShowArray arrdata
2   Set objBytes = New clsByteArray
3   objBytes.AppendBytes arrdata
4   mcolMirrorTrackers("T" & CStr(Index)).ProcessListing objBytes
5   Set objBytes = Nothing
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTracker_DataArrival", Err.Description, Erl

End Sub

Private Sub sckTracker_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo errhandler
    
    'Debug.Print "sckTracker ERROR " & Index & " = " & Description
    StatusMsg "...unable to connect to " & mcolMirrorTrackers("T" & Index).Address, MirrorInfo
    DestroyMirrorSocket sckTracker(Index)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckTracker_Error", Err.Description, Erl

End Sub

Private Sub DestroyMirrorSocket(sckMirror As Winsock)
    mcolMirrorTrackers.Remove "T" & CStr(sckMirror.Index)
    sckMirror.Close
    'Unload sckMirror
End Sub

Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)
    Dim objBytes As clsByteArray
    Dim arrServerData() As Byte
    Dim sServerData As String
    Dim objServer As clsServer
    Dim sData As String
    On Error GoTo errhandler
    
1   Set objBytes = New clsByteArray
2   sckUDP.GetData sData, vbString
    
3   objBytes.AppendString sData
4   Set objServer = BuildServerInfo(objBytes, sckUDP.RemoteHostIP)
5   If Not objServer Is Nothing Then
6       objServer.LastCheckedIn = Now
7       AddServer objServer
8       ReCalculateCounts
    End If
    
9  Set objBytes = Nothing

    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "sckUDP_DataArrival", Err.Description, Erl
End Sub

Private Sub UpdateListingCounter()
    lblListingCount.Caption = mlListingCounter
End Sub

Public Sub DoMirror()
    Dim objMirror As clsMirroredTracker
    Dim objNewMirrorTracker As clsTracker
    
    CleanUpSockets sckTracker, "sckTracker", True
    
    Set mcolMirrorTrackers = New Collection
    
    For Each objMirror In mobjSettings.TrackerMirrors
        If objMirror.Enabled Then
            mlNextTrackerSockIndex = mlNextTrackerSockIndex + 1
            Load sckTracker(mlNextTrackerSockIndex)
            sckTracker(mlNextTrackerSockIndex).Tag = Now
            
            StatusMsg "Connecting to '" & objMirror.Address & "'...", MirrorInfo
            Set objNewMirrorTracker = New clsTracker
            objNewMirrorTracker.Address = objMirror.Address
            Set objNewMirrorTracker.WinSockControl = sckTracker(mlNextTrackerSockIndex)
            Set objNewMirrorTracker.MainForm = Me
            objNewMirrorTracker.Connect
            mcolMirrorTrackers.Add objNewMirrorTracker, "T" & CStr(mlNextTrackerSockIndex)
        End If
    Next objMirror
    Set objNewMirrorTracker = Nothing
    
End Sub

Private Sub tmrCheck_Timer()
    Static nMinsSinceLastMirror As Integer
    Static bDoingMirror As Boolean
    
    On Error GoTo errhandler

    If (Not bDoingMirror) And (nMinsSinceLastMirror >= mobjSettings.MirrorInterval) Then
        bDoingMirror = True
        nMinsSinceLastMirror = 0
        If mobjSettings.DoMirroring Then
            DoMirror
        End If
        bDoingMirror = False
    Else
        nMinsSinceLastMirror = nMinsSinceLastMirror + 1
    End If
    
    ExpireServers
    CleanUpSockets sckTCP, "sckTCP"
    
    Exit Sub
errhandler:
    bDoingMirror = False
    ErrorReport Err.Number, MOD_NAME, "tmrCheck_Timer", Err.Description, Erl
End Sub

Private Sub CleanUpSockets(sckArray As Variant, Optional sDesc As String = "sck?", Optional bCleanAll As Boolean = False)
    Dim sckThisSocket As Winsock
    Dim sSocketDate As String
    Dim anCloseThese() As Integer
    Dim nCloseCount As Integer
    Dim vntThisSocket As Variant
    Dim nSocketAge As Integer
    
    On Error GoTo errhandler
    
    For Each sckThisSocket In sckArray
        If sckThisSocket.Index > 0 Then
            sSocketDate = sckThisSocket.Tag
            nSocketAge = DateDiff("n", CDate(sSocketDate), Now)
            'Debug.Print sDesc & "(" & CStr(sckThisSocket.Index) & ") now " & nSocketAge & " minutes old"
            If (nSocketAge >= MAX_SOCKET_AGE Or sckThisSocket.State = sckError) Or bCleanAll Then
                nCloseCount = nCloseCount + 1
                ReDim Preserve anCloseThese(1 To nCloseCount)
                anCloseThese(nCloseCount) = sckThisSocket.Index
            End If
        End If
    Next sckThisSocket
    
    If nCloseCount > 0 Then
        For Each vntThisSocket In anCloseThese()
            'Debug.Print "Cleaning up socket " & vntThisSocket
            sckArray(CInt(vntThisSocket)).Close
            Unload sckArray(CInt(vntThisSocket))
        Next vntThisSocket
    End If
    'Debug.Print sDesc & " count is now " & sckArray.Count
    
    Exit Sub
    
errhandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ExpireServers()
    Dim objServer As clsServer
    Dim nMinutesSinceLastCheckedIn As Double
    Dim colServersToRetire As Collection
    
    Set colServersToRetire = New Collection
    
    For Each objServer In mobjServers
        nMinutesSinceLastCheckedIn = DateDiff("n", objServer.LastCheckedIn, Now)
        If nMinutesSinceLastCheckedIn >= mobjSettings.ExpireInterval And Not (objServer.ServerType = Fake) Then
            colServersToRetire.Add objServer
        End If
    Next objServer
    
    For Each objServer In colServersToRetire
        mobjServers.Remove objServer
        Debug.Print "Expiring server " & objServer.ServerName
        'If Not objServer.ServerType = Mirrored Then
            StatusMsg "Server at " & objServer.IPString & "[" & objServer.ServerName & "] has expired", ServerExpired
        'End If
    Next objServer
    
    ReCalculateCounts
    
End Sub

Public Sub FlashForm()
    tmrAlert.Interval = 500
End Sub

Private Sub tmrAlert_Timer()
    Static nCount As Integer
    Static bIconIsRed As Boolean
    On Error GoTo errhandler
    
    bIconIsRed = Not bIconIsRed
    SwapIcon bIconIsRed
    
    nCount = nCount + 1
    If nCount = MAX_ICON_FLASH Then
        nCount = 0
        If bIconIsRed Then
            SwapIcon False
        End If
        tmrAlert.Interval = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "tmrAlert_Timer", Err.Description, Erl
End Sub

Private Sub SwapIcon(bShowRed As Boolean)
    Me.Icon = ImageList1.ListImages(IIf(bShowRed, "red", "blue")).Picture
    mnfIconData.hIcon = Me.Icon
    Shell_NotifyIcon NIM_MOD, mnfIconData
End Sub

