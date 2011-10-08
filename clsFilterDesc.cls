VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msName As String
Private mbExactMatch As Boolean
Private mbMatchCase As Boolean
Private msText As String
Private mbBlock As Boolean
Private mbValid As Boolean
Private mbEnabled As Boolean

Private Const MY_TYPE = FILTER_TYPE_DESC

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
    IFilterRule_Description = IIf(IFilterRule_Block, "Block", "List") & " servers if the Description " & _
                            IIf(mbExactMatch, " is ", " contains ") & msText & _
                            " (" & IIf(mbMatchCase, " match ", " ignore ") & " case)"
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
    
    If Not IFilterRule_Valid Then
        ' No match
    Else
        If mbMatchCase Then
            If mbExactMatch Then
                bMatch = (msText = objServer.Description)
            Else
                bMatch = (InStr(1, objServer.Description, msText) > 0)
            End If
        Else
            If mbExactMatch Then
                bMatch = (UCase(msText) = UCase(objServer.Description))
            Else
                bMatch = (InStr(1, UCase(objServer.Description), UCase(msText)) > 0)
            End If
        End If
    End If
    IFilterRule_CheckForMatch = bMatch

End Function

Private Function IFilterRule_Clone() As IFilterRule
    Dim objClone As clsFilterDesc
    Dim objIClone As IFilterRule
    
    Set objClone = New clsFilterDesc
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
    Build asParts(1), (asParts(4) = "1"), (asParts(2) = "1"), (asParts(3) = "1"), (asParts(5) = "1"), asParts(6)
    
    Exit Sub
ErrHandler:
    mbValid = False
    
End Sub

Private Property Get IFilterRule_Valid() As Boolean
    IFilterRule_Valid = mbValid
End Property

' 0 -> Type
' 1 -> Text
' 2 -> Block
' 3 -> Exact Match
' 4 -> Match Case
' 5 -> Enabled
' 6 -> Name

Private Function IFilterRule_WriteToString() As String
    IFilterRule_WriteToString = FILTER_TYPE_DESC & SEP_CHAR & msText & SEP_CHAR & IIf(mbBlock, "1", "0") & SEP_CHAR & _
                                IIf(mbExactMatch, "1", "0") & SEP_CHAR & mbMatchCase & SEP_CHAR & _
                                IIf(mbEnabled, "1", "0") & SEP_CHAR & msName
                                
End Function

Public Sub Build(sText As String, bMatchCase As Boolean, bBlock As Boolean, bExactMatch As Boolean, _
                    bEnabled As Boolean, Optional sName As String = "")
    msText = sText
    mbMatchCase = bMatchCase
    mbBlock = bBlock
    mbExactMatch = bExactMatch
    msName = sName
    mbEnabled = bEnabled
    mbValid = True
End Sub



