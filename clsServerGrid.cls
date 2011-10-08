VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mlvwGrid As ListView
Attribute mlvwGrid.VB_VarHelpID = -1
Private mobjServers As clsServers

Public Event SelectionChanged(objServer As clsServer)

' #####################################
' PROPERTIES

' ListViewControl <->
Public Property Set ListViewControl(lvwGrid As ListView)
    Set mlvwGrid = lvwGrid
End Property
Public Property Get ListViewControl() As ListView
    Set ListViewControl = mlvwGrid
End Property

' SelectedServer ->
Public Property Get SelectedServer() As clsServer
    Dim objServer As clsServer
    Dim lviSelected As ListItem
    
    On Error GoTo ErrHandler
    
    Set lviSelected = mlvwGrid.SelectedItem
    If lviSelected Is Nothing Then
        Set SelectedServer = Nothing
    Else
        Set objServer = mobjServers.Item(lviSelected.Key)
        Set SelectedServer = objServer
        Set objServer = Nothing
    End If
    
    Exit Function
ErrHandler:
    Set SelectedServer = Nothing

End Property

' #####################################
' METHODS

Public Function GetRegString() As String
    GetRegString = mobjServers.BuildRegString
End Function

Public Sub AddNewServer(objServer As clsServer, Optional nPosn As Integer = -1)
    If mobjServers.HasKey(objServer.UniqueID) Then
        Err.Raise ERR_BASE, , "A server at this IP address and port is already in the grid"
    Else
        mobjServers.AddOrUpdate objServer
        AddServerToGrid objServer, nPosn
    End If
End Sub

Public Function DeleteServer(objServer As clsServer) As Integer

    If mobjServers.HasKey(objServer.UniqueID) Then
        mobjServers.Remove objServer
        DeleteServer = DeleteServerFromGrid(objServer)
    Else
        Err.Raise ERR_BASE, , "This server is not present in the grid"
    End If

End Function

Public Sub UpdateServer(objServer As clsServer)
    Dim nServerPosn As Integer
    nServerPosn = DeleteServer(SelectedServer)
    AddNewServer objServer, nServerPosn
End Sub

Public Sub MergeServers(objServers As clsServers)
    Dim objServer As clsServer
    Dim nPosn As Integer
    
    For Each objServer In objServers
        If objServer.ServerType = Fake Then
            objServers.Remove objServer
        End If
    Next objServer
    
    nPosn = 1
    For Each objServer In mobjServers
        'If objServers.HasKey(objServer.UniqueID) Then
            'If MsgBox("A real server with an IP of " & objServer.IPString & " and port number of " & _
                       objServer.Port & " already exists - do you want to over-write it with the fake one?", _
                       vbQuestion + vbYesNo, "Careful...") = vbYes Then
        '        objServers.AddOrUpdate objServer
        '    End If
        'Else
            objServers.AddOrUpdate objServer, nPosn
            nPosn = nPosn + 1
        'End If
    Next objServer
    
End Sub

Public Sub RequestSelectionEvent()
    RaiseSelectionEvent
End Sub
' #####################################

Private Sub RaiseSelectionEvent()
    RaiseEvent SelectionChanged(SelectedServer)
End Sub

Private Sub RefreshGrid()
    Dim objServer As clsServer
    
    mlvwGrid.ListItems.Clear
    For Each objServer In mobjServers
        If objServer.ServerType = Fake Then
            AddServerToGrid objServer
        End If
    Next objServer
    
End Sub

Private Function DeleteServerFromGrid(objServer As clsServer) As Integer
    DeleteServerFromGrid = mlvwGrid.ListItems.Item(objServer.UniqueID).Index
    mlvwGrid.ListItems.Remove objServer.UniqueID
    RaiseSelectionEvent
End Function

Private Sub AddServerToGrid(objServer As clsServer, Optional nPosn As Integer = -1)
    Dim lviServer As ListItem
    Dim lviSubServer As ListSubItem
    
    If Not objServer Is Nothing Then
        If nPosn >= 0 Then
            Set lviServer = mlvwGrid.ListItems.Add(nPosn, objServer.UniqueID)
        Else
            Set lviServer = mlvwGrid.ListItems.Add(, objServer.UniqueID)
        End If
        lviServer.Text = objServer.ServerName
        Set lviSubServer = lviServer.ListSubItems.Add
        lviSubServer.Text = objServer.Description
        lviServer.Selected = True
        Set lviServer = Nothing
        RaiseSelectionEvent
    End If
    
End Sub

Private Sub Class_Initialize()
    Set mobjServers = New clsServers
End Sub

Private Sub Class_Terminate()
    Set mobjServers = Nothing
End Sub

Private Sub mlvwGrid_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RaiseSelectionEvent
End Sub
