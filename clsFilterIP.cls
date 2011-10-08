VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjMinIP As clsIP
Private mobjMaxIP As clsIP
Private mbBlock As Boolean
Private mbValid As Boolean
Private mbEnabled As Boolean
Private msName As String
Private Const MY_TYPE = FILTER_TYPE_IP

Implements IFilterRule

' #################################################
' PROPERTIES

' Enabled <->
Private Property Let IFilterRule_Enabled(bEnabled As Boolean)
    mbEnabled = bEnabled
End Property
Private Property Get IFilterRule_Enabled() As Boolean
    IFilterRule_Enabled = mbEnabled
End Property

' Name ->
Private Property Get IFilterRule_Name() As String
    IFilterRule_Name = msName
End Property

' Description ->
Private Property Get IFilterRule_Description() As String
    IFilterRule_Description = IIf(IFilterRule_Block, "Block", "List") & " servers if the IP Address is "
    If IPAddress.Compare(mobjMinIP, IPAddress.typCompare.Equal, mobjMaxIP) Then
        IFilterRule_Description = IFilterRule_Description & mobjMinIP.Describe
    Else
        IFilterRule_Description = IFilterRule_Description & "between " & mobjMinIP.Describe & " and " & mobjMaxIP.Describe
    End If
End Property

' Block ->
Private Property Get IFilterRule_Block() As Boolean
    IFilterRule_Block = mbBlock
End Property

' UnqiueID ->
Private Property Get IFilterRule_UniqueID() As String
    IFilterRule_UniqueID = IFilterRule_WriteToString
End Property

' #################################################
' METHODS

Private Function IFilterRule_CheckForMatch(objServer As clsServer) As Boolean
    Dim bMatch As Boolean
    Dim sText As String
    Dim objThisIP As clsIP
    
    If Not IFilterRule_Valid Then
        ' No match
    Else
        Set objThisIP = New clsIP
        objThisIP.BuildFromString objServer.IPString
        If IPAddress.Compare(mobjMinIP, IPAddress.typCompare.Equal, mobjMaxIP) Then
            bMatch = (IPAddress.Compare(objThisIP, IPAddress.typCompare.Equal, mobjMinIP))
        Else
            bMatch = (IPAddress.Compare(objThisIP, IPAddress.typCompare.GreaterThanOrEqual, mobjMinIP)) And _
                     (IPAddress.Compare(objThisIP, IPAddress.typCompare.LessThanOrEqual, mobjMaxIP))
        End If
    End If
    IFilterRule_CheckForMatch = bMatch

End Function

Private Function IFilterRule_Clone() As IFilterRule
    Dim objClone As clsFilterIP
    Dim objIClone As IFilterRule
    
    Set objClone = New clsFilterIP
    Set objIClone = objClone
    objIClone.BuildFromString IFilterRule_WriteToString
    
    Set IFilterRule_Clone = objIClone
    Set objClone = Nothing
    Set objIClone = Nothing
End Function

' #########################

Private Sub IFilterRule_BuildFromString(sData As String)
    Dim asParts() As String
    
    On Error GoTo ErrHandler
    
    asParts = Split(sData, SEP_CHAR)
    Build asParts(1), asParts(2), (asParts(3) = "1"), (asParts(4) = "1"), asParts(5)
    
    Exit Sub
ErrHandler:
    mbValid = False
    
End Sub

' 1 -> MinIP
' 2 -> MaxIP
' 3 -> Block
' 4 -> Enabled
' 5 -> Name

Private Property Get IFilterRule_Valid() As Boolean
    IFilterRule_Valid = mbValid
End Property

Private Function IFilterRule_WriteToString() As String
    IFilterRule_WriteToString = MY_TYPE & SEP_CHAR & mobjMinIP.Describe & SEP_CHAR & mobjMaxIP.Describe & SEP_CHAR & _
                                IIf(mbBlock, "1", "0") & SEP_CHAR & IIf(mbEnabled, "1", "0") & SEP_CHAR & msName
End Function

Public Sub Build(sMinIP As String, sMaxIP As String, bBlock As Boolean, bEnabled As Boolean, Optional sName = "")
    Set mobjMinIP = New clsIP
    mobjMinIP.BuildFromString sMinIP
    Set mobjMaxIP = New clsIP
    mobjMaxIP.BuildFromString sMaxIP
    msName = sName
    mbBlock = bBlock
    mbEnabled = bEnabled
    mbValid = True
    
End Sub



