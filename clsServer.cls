VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IPAddress(1 To 4) As Byte
Private mlPort As Long
Private mlUserCount As Integer
Private msServerName As String
Private msDescription As String
Public LastCheckedIn As Date
Private msIPString As String
Private mtServerType As ServerType
Public DoFilterCalc As Boolean
Private mbFiltered As Boolean
Private msFilteredBy As String

Public Enum ServerType
    Real = 1
    Fake = 2
    Mirrored = 3
End Enum

Public Function Clone() As clsServer
    Dim objClone As clsServer
    Set objClone = New clsServer
    objClone.Description = Description
    objClone.LastCheckedIn = LastCheckedIn
    objClone.SetIP IPAddress(1), IPAddress(2), IPAddress(3), IPAddress(4)
    objClone.Port = Port
    objClone.ServerName = ServerName
    objClone.UserCount = UserCount
    Set Clone = objClone
    Set objClone = Nothing
End Function

Public Property Get FilteredBy() As String
    FilteredBy = msFilteredBy
End Property

' Filtered
Public Property Get Filtered() As Boolean
    Dim objFilter As IFilterRule
    Dim bFoundMatch As Boolean
    
    If Me.ServerType = Fake Then
        Filtered = False
    Else
        If DoFilterCalc Then
            For Each objFilter In mobjFilters
                If objFilter.Enabled Then
                    bFoundMatch = objFilter.CheckForMatch(Me)
                    If bFoundMatch Then
                        'Debug.Print "Server " & Me.ServerName & " blocked by " & objFilter.Description
                        Exit For
                    End If
                End If
            Next objFilter
            
            If bFoundMatch Then
                mbFiltered = objFilter.Block
                msFilteredBy = IIf(mbFiltered, objFilter.Name, "")
            Else
                mbFiltered = Not mobjSettings.FilterDefault
                msFilteredBy = IIf(mbFiltered, "Default Rule", "")
            End If
            DoFilterCalc = False
        End If
        Filtered = mbFiltered
    End If
End Property

' ServerType
Public Property Get ServerType() As ServerType
    ServerType = mtServerType
End Property
Public Property Let ServerType(tServerType As ServerType)
    mtServerType = tServerType
End Property

' Port
Public Property Get Port() As Long
    Port = mlPort
End Property
Public Property Let Port(lPort As Long)
    If lPort < 0 Or lPort > 65535 Then Err.Raise ERR_BASE, , "Invalid port number " & CStr(lPort)
    mlPort = lPort
End Property

' UserCount
Public Property Get UserCount() As Long
    UserCount = mlUserCount
End Property
Public Property Let UserCount(lUserCount As Long)
    If lUserCount < 0 Then lUserCount = 0
    If lUserCount > 65535 Then lUserCount = 65535
    mlUserCount = lUserCount
End Property

' ServerName
Public Property Get ServerName() As String
    ServerName = msServerName
End Property
Public Property Let ServerName(sServerName As String)
    If Len(sServerName) = 0 Then Err.Raise ERR_BASE, , "Invalid Server Name"
    If Len(sServerName) > 255 Then sServerName = Left$(sServerName, 255)
    msServerName = sServerName
End Property

' Description
Public Property Get Description() As String
    Description = msDescription
End Property
Public Property Let Description(sDescription As String)
    If Len(sDescription) > 255 Then sDescription = Left$(sDescription, 255)
    msDescription = sDescription
End Property

' IPString
Public Property Get IPString() As String
    IPString = msIPString
End Property
Public Property Let IPString(sIPString As String)
    msIPString = sIPString
End Property

' UniqueID
Public Property Get UniqueID() As String
    UniqueID = IPString & ":" & CStr(Port)
End Property

Public Function SetIP(nIP1 As Byte, nIP2 As Byte, nIP3 As Byte, nIP4 As Byte)
    IPAddress(1) = nIP1
    IPAddress(2) = nIP2
    IPAddress(3) = nIP3
    IPAddress(4) = nIP4
    IPString = CStr(nIP1) & "." & CStr(nIP2) & "." & CStr(nIP3) & "." & CStr(nIP4)
End Function

Public Function BuildListingEntry() As clsByteArray
    Dim objBytes As clsByteArray
    
    Set objBytes = New clsByteArray
    
    objBytes.AppendBytes IPAddress
    objBytes.AppendNumber CLng(Port), 2
    objBytes.AppendNumber CLng(UserCount), 2
    objBytes.AppendNumber 0, 2
    objBytes.AppendNumber Len(ServerName), 1
    objBytes.AppendString ServerName
    objBytes.AppendNumber Len(Description), 1
    objBytes.AppendString Description
    
    Set BuildListingEntry = objBytes
    
    Set objBytes = Nothing
End Function

Private Sub Class_Initialize()
    DoFilterCalc = True
End Sub
