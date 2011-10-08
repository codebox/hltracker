VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolServers As Collection
Private mlID As Long
Private mlServerCount As Long
Private mlUserCount As Long

Public SendStatusMsgs As Boolean

' ########################################################
' PROPERTIES

Public Property Get UserCount() As Long
    UserCount = mlUserCount
End Property

Public Property Get Count() As Long
    Count = mlServerCount
End Property

' ########################################################
' METHODS

Public Sub NotifyFiltersChanged()
    Dim objServer As clsServer
    Dim lCount As Long
    
    For Each objServer In mcolServers
        objServer.DoFilterCalc = True
    Next objServer
    
    mlUserCount = 0
    mlServerCount = 0
    For Each objServer In mcolServers
        'Debug.Assert InStr(1, objServer.Description, "x") = 0
        If Not objServer.Filtered Then
            mlUserCount = mlUserCount + objServer.UserCount
            mlServerCount = mlServerCount + 1
        End If
    Next objServer
    
    frmMain.ReCalculateCounts
    
    Set objServer = Nothing

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolServers.[_NewEnum]
End Function

Public Function Item(vntKey As Variant) As clsServer
    Set Item = mcolServers.Item(vntKey)
End Function

Public Function HasKey(vntKey As Variant) As Boolean
    Dim objDummy As clsServer
    
    On Error Resume Next
    
    Set objDummy = mcolServers.Item(vntKey)
    HasKey = (Err.Number = 0)
    Set objDummy = Nothing

End Function

Public Function BuildRegString() As String
    Dim objServer As clsServer
    Dim sText As String
    
    For Each objServer In mcolServers
        sText = sText & IIf(Len(sText) > 0, Chr(0), "") & objServer.ServerName & Chr(1) & objServer.IPString & Chr(1) & objServer.Port & Chr(1) & objServer.UserCount & Chr(1) & objServer.Description
    Next objServer
    BuildRegString = sText
    
End Function

Public Sub AddOrUpdate(objServer As clsServer, Optional nPosn As Integer = 0)
    Dim objExistingServer As clsServer
    Dim objFilter As IFilterRule
    Dim bBlockServer As Boolean
    Dim bRuleApplies As Boolean
    Dim bGettingExistingServer As Boolean
    Dim bIsNewServer As Boolean
    
    On Error GoTo ErrHandler
    
    bGettingExistingServer = True
    Set objExistingServer = mcolServers(objServer.UniqueID)
    bGettingExistingServer = False

    If Not objServer.Filtered Then
      ' Adjust the user count
        mlUserCount = mlUserCount + objServer.UserCount
        mlServerCount = mlServerCount + 1
    End If
    If Not objExistingServer.Filtered Then
      ' Adjust the user count
        mlServerCount = mlServerCount - 1
        mlUserCount = mlUserCount - objExistingServer.UserCount
        'objServer.UserCount = objServer.UserCount - objExistingServer.UserCount
    End If
    
  ' Replace the existing one with the new one...
    mcolServers.Remove objServer.UniqueID
    If nPosn > 0 Then
        mcolServers.Add objServer, objServer.UniqueID, nPosn
    Else
        mcolServers.Add objServer, objServer.UniqueID
    End If

AfterGet:
    
    frmMain.ReCalculateCounts
    'logit objServer.IPString & "|" & objServer.Port & "|" & objServer.UserCount & "|" & objServer.ServerName & "|" & objServer.Description, 3
    If objServer.Filtered Then
        If bIsNewServer Then
            StatusMsg "Server '" & objServer.ServerName & "' was filtered " & _
                IIf(Len(objServer.FilteredBy) > 0, "[" & objServer.FilteredBy & "]", "out"), ServerFiltered
        End If
    Else
        If bIsNewServer Then
            If SendStatusMsgs And Not (objServer.ServerType = Fake) Then
                StatusMsg "Added Server: '" & objServer.ServerName & "' at " & objServer.IPString, ServerAdded
            End If
            If mobjSettings.AlertNew And (objServer.ServerType <> Mirrored) Then DoAlert
        End If
    End If
    
    Exit Sub
ErrHandler:
    If bGettingExistingServer Then
      ' It's a new server...
        bIsNewServer = True
        If nPosn > 0 And mcolServers.Count >= nPosn Then
            mcolServers.Add objServer, objServer.UniqueID, nPosn
        Else
            mcolServers.Add objServer, objServer.UniqueID
        End If
        If Not objServer.Filtered Then
            mlServerCount = mlServerCount + 1
            mlUserCount = mlUserCount + objServer.UserCount
        End If
        Resume AfterGet
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Public Sub Remove(objServer As clsServer)
    Dim objExistingServer As clsServer

    On Error GoTo ErrHandler
    
    Set objExistingServer = mcolServers(objServer.UniqueID)
    mcolServers.Remove objServer.UniqueID
    If Not objExistingServer.Filtered Then
        If mobjSettings.AlertExpire Then DoAlert
        mlServerCount = mlServerCount - 1
        mlUserCount = mlUserCount - objExistingServer.UserCount
    End If
    frmMain.ReCalculateCounts
    
    Exit Sub
ErrHandler:
    'StatusMsg "ERROR (Remove):" & Err.Description
End Sub

Private Sub Class_Initialize()
    Set mcolServers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolServers = Nothing
End Sub
