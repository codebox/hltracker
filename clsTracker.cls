VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Address As String
Public WinSockControl As Winsock
Public UniqueID As String
Public MainForm As frmMain

Private mbNeedMagic As Boolean
Private mbNeedInfo As Boolean

Private mobjBufferedData As clsByteArray
Private mobjCurrentHeader As clsServerInfoHeader

Private Const TRACKER_PORT = 5498

Public Event Error(sInfo As String)
Public Event ServerData(colServers As Collection)
Public Event GotServerCount(sTracker As String, lCount As Long)
Public Event ListingComplete(sckControl As Winsock)

Public Sub Connect()
    Set mobjCurrentHeader = New clsServerInfoHeader
    mbNeedMagic = True
    mbNeedInfo = True
    mobjBufferedData.Clear
    If WinSockControl.State <> sckClosed Then WinSockControl.Close
    WinSockControl.RemoteHost = Address
    WinSockControl.RemotePort = TRACKER_PORT
    WinSockControl.Connect
End Sub

Private Sub ProcessInfoRecord(objBytes As clsByteArray, ByRef lDataPointer As Long)
    Static bSentCount As Boolean
    mobjCurrentHeader.DataSize = CLng(objBytes.GetInteger(lDataPointer + 2))
    mobjCurrentHeader.ServerCount = CLng(objBytes.GetInteger(lDataPointer + 4))
    
    If Not bSentCount Then
      ' Make sure it only gets sent once per mirror...
        MainForm.CallBackGotServerCount Address, mobjCurrentHeader.ServerCount
        bSentCount = True
    End If
    lDataPointer = lDataPointer + 8

End Sub

Public Sub ProcessListing(objBytes As clsByteArray)
    Dim sData As String
    Dim lStartOfDataForThisServer As Long
    Dim IP(1 To 4) As Byte
    Dim lPort As Long
    Dim lUsers As Long
    Dim nLenOfName As Integer
    Dim nLenOfDesc As Integer
    Dim sName As String
    Dim sDescription As String
    Dim objNewServer As clsServer
    Dim lDataSize As Long
    Dim lDataPointer As Long
    Dim sMagic As String
    Dim nVersion As Integer
    Dim nMessageType As Integer
    Dim nDataSize As Integer
    Dim nNumberOfServers As Integer
    Dim bReadingBytes As Boolean
    Dim colServers As Collection
    Dim objTemp As clsByteArray
    
    On Error GoTo ErrHandler
    
    Set objTemp = New clsByteArray
    
  ' Start with any data left over from last time...
    If mobjBufferedData.Size > 0 Then
        objTemp.AppendByteObject mobjBufferedData
        mobjBufferedData.Clear
    End If
    
    objTemp.AppendByteObject objBytes
    Set colServers = New Collection
    
    lDataSize = objTemp.Size
    
    lDataPointer = 1
    
    bReadingBytes = True
    
    If mbNeedMagic Then
        sMagic = objTemp.GetString(lDataPointer, 4)
        nVersion = objTemp.GetInteger(lDataPointer + 4)
        mbNeedMagic = False
        lDataPointer = lDataPointer + 6
    End If
    
    If mbNeedInfo Then
        ProcessInfoRecord objTemp, lDataPointer
        mbNeedInfo = False
    End If
    
    Do While mobjCurrentHeader.ServersSoFar < mobjCurrentHeader.ServerCount
        If objTemp.GetByte(lDataPointer) = 0 Then
            ProcessInfoRecord objTemp, lDataPointer
        Else
            Set objNewServer = New clsServer
1           objNewServer.ServerType = Mirrored
2           objNewServer.SetIP objTemp.GetByte(lDataPointer), objTemp.GetByte(lDataPointer + 1), objTemp.GetByte(lDataPointer + 2), objTemp.GetByte(lDataPointer + 3)
3           objNewServer.Port = objTemp.GetByte(lDataPointer + 4) * 256& + objTemp.GetByte(lDataPointer + 5)
4           objNewServer.UserCount = objTemp.GetInteger(lDataPointer + 6)
            
5           nLenOfName = objTemp.GetByte(lDataPointer + 10)
6           objNewServer.ServerName = objTemp.GetString(lDataPointer + 11, CLng(nLenOfName))
            
7           nLenOfDesc = objTemp.GetByte(lDataPointer + 11 + nLenOfName)
8           objNewServer.Description = objTemp.GetString(lDataPointer + 11 + nLenOfName + 1, CLng(nLenOfDesc))
9           lDataPointer = lDataPointer + 11 + nLenOfName + 1 + nLenOfDesc
11          objNewServer.LastCheckedIn = Now
            colServers.Add objNewServer
13          mobjCurrentHeader.ServersSoFar = mobjCurrentHeader.ServersSoFar + 1
        End If
    Loop
14  bReadingBytes = False
    
15  MainForm.CallBackServerData colServers
16  MainForm.CallBackListingComplete Address, WinSockControl
    
    Exit Sub
OutOfLoop:
    bReadingBytes = False
17  If objTemp.Size >= lDataPointer Then
18      mobjBufferedData.AppendByteObject objTemp.GetByteObject(lDataPointer, objTemp.Size - lDataPointer + 1)
    End If
19  mobjCurrentHeader.CurrentDataOffset = lDataPointer + mobjCurrentHeader.CurrentDataOffset - 1
    
20  mobjCurrentHeader.HeaderSeqID = mobjCurrentHeader.HeaderSeqID + 1
    
21  MainForm.CallBackServerData colServers
    Set colServers = Nothing
    Set objNewServer = Nothing
    
    Exit Sub
ErrHandler:
    If bReadingBytes Then
        'Debug.Print "ERR " & Err.Description & " line " & Erl
        
        Resume OutOfLoop
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub Class_Initialize()
    Set mobjBufferedData = New clsByteArray
End Sub

Private Sub Class_Terminate()
    Set mobjBufferedData = Nothing
End Sub


