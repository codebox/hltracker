VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_DATA_SIZE = 3000

Private mcolWaitingSockets As Collection

Public Sub AddNew(sckWinsock As Winsock, objData As clsByteArray)
    Dim objNewSocket As clsWaitingSocket
    
    Set objNewSocket = New clsWaitingSocket
    Set objNewSocket.Socket = sckWinsock
    Set objNewSocket.Data = objData
    mcolWaitingSockets.Add objNewSocket, "S" & CStr(objNewSocket.Socket.Index)
    SendNextBatch sckWinsock
    Set objNewSocket = Nothing
End Sub

Public Sub SendNextBatch(sckWinsock As Winsock)
    Dim objNewSocket As clsWaitingSocket
    Dim objRemainingData As clsByteArray
    Dim bCheckingForObject As Boolean
    Dim objDataToSend As clsByteArray
    
    On Error GoTo errhandler
    
    bCheckingForObject = True
    Set objNewSocket = mcolWaitingSockets.Item("S" & sckWinsock.Index)
    bCheckingForObject = False
    
    If objNewSocket.Data.Size > MAX_DATA_SIZE Then
        Set objDataToSend = objNewSocket.Data.GetByteObject(1, MAX_DATA_SIZE)
        Set objNewSocket.Data = objNewSocket.Data.GetByteObject(MAX_DATA_SIZE + 1, objNewSocket.Data.Size - MAX_DATA_SIZE)
        logit "Sending another batch of listing"
    Else
        Set objDataToSend = objNewSocket.Data
        logit "Sending last batch of listing"
        mcolWaitingSockets.Remove "S" & CStr(sckWinsock.Index)
    End If
    'logit "Sending: " & objDataToSend.Describe
    
    objNewSocket.Socket.SendData objDataToSend
    Set objDataToSend = Nothing
    
    Exit Sub
errhandler:
    If bCheckingForObject Then
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub Class_Initialize()
    Set mcolWaitingSockets = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolWaitingSockets = Nothing
End Sub
