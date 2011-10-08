VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolRanges As Collection

' ########################################################
' PROPERTIES

Public Property Get Count() As Long
    Count = mcolRanges.Count
End Property

' ########################################################
' METHODS

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolRanges.[_NewEnum]
End Function

Public Function Item(vntKey As Variant) As clsIPRange
    Set Item = mcolRanges.Item(vntKey)
End Function

Public Sub Remove(vntKey As Variant)
    mcolRanges.Remove vntKey
End Sub

Public Sub Add(objNewIPRange As clsIPRange)
    On Error GoTo ErrHandler
    
    mcolRanges.Add objNewIPRange, objNewIPRange.UniqueID
    
    Exit Sub
ErrHandler:
    Remove objNewIPRange.UniqueID
    mcolRanges.Add objNewIPRange, objNewIPRange.UniqueID
End Sub

Public Function AddressIsBlocked(sIP As String) As Boolean
    Dim objRange As clsIPRange
    
    On Error Resume Next
    
    For Each objRange In mcolRanges
        If objRange.Enabled Then
            If objRange.IsInRange(sIP) Then
                AddressIsBlocked = True
                Exit For
            End If
        End If
    Next objRange
    
End Function

Public Function BuildRegString() As String
    Dim objRange As clsIPRange
    Dim sText As String
    
    For Each objRange In mcolRanges
        sText = sText & IIf(Len(sText) > 0, Chr(0), "") & objRange.WriteToString
    Next objRange
    BuildRegString = sText
    
End Function

Public Sub BuildFromRegString(sData As String)
    Dim objRange As clsIPRange
    Dim asRanges() As String
    Dim vntRange As Variant
    
    Set mcolRanges = New Collection
    asRanges = Split(sData, Chr(0))
    
    For Each vntRange In asRanges
        Set objRange = New clsIPRange
        objRange.BuildFromString CStr(vntRange)
        Add objRange
    Next vntRange
    Set objRange = Nothing
   
End Sub

Private Sub Class_Initialize()
    Set mcolRanges = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolRanges = Nothing
End Sub


