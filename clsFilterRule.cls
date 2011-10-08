VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Field As String
Public Text As String
Public Permit As Boolean
Public Enabled As Boolean


Public Function CheckForMatch(objServer As clsServer) As Boolean
    Dim bMatch As Boolean
    
    If objServer.ServerType <> Fake Then
        Select Case UCase(Field)
            Case "DESCRIPTION"
                bMatch = (InStr(1, objServer.Description, Text) > 0)
            Case "NAME"
                bMatch = (InStr(1, objServer.ServerName, Text) > 0)
            Case "IP"
                bMatch = (InStr(1, objServer.IPString, Text) > 0)
            Case "USER COUNT"
                bMatch = (InStr(1, CStr(objServer.UserCount), Text) > 0)
            Case "PORT"
                bMatch = (InStr(1, CStr(objServer.Port), Text) > 0)
        End Select
    End If
    CheckForMatch = bMatch
    
End Function

Public Property Get UniqueID() As String
    UniqueID = Field & Text & IIf(Permit, "1", "0")
End Property

Public Function Describe() As String
    Describe = IIf(Permit, "List", "Block") & " the server if the " & Field & " contains '" & Text & "'"
End Function

Public Sub BuildFromString(sData As String)
    Dim asData() As String
    
    asData = Split(sData, SEP_CHAR)
    If UBound(asData) <> 3 Or LBound(asData) <> 0 Then
    Else
        Permit = (asData(0) = "1")
        Field = asData(1)
        Text = asData(2)
        Enabled = (asData(3) = "1")
    End If
End Sub

Public Function WriteToString() As String
    WriteToString = IIf(Permit, "1", "0") & SEP_CHAR & Field & SEP_CHAR & Text & SEP_CHAR & IIf(Enabled, "1", "0")
End Function

Public Function Clone() As clsFilterRule
    Dim objNewFilter As clsFilterRule
    
    Set objNewFilter = New clsFilterRule
    objNewFilter.BuildFromString WriteToString
    Set Clone = objNewFilter
    Set objNewFilter = Nothing
End Function
