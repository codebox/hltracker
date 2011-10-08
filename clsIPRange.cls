VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbIsValid As Boolean
Private mobjFrom As clsIP
Private mobjTo As clsIP
Private msMyID As String

Public Enabled As Boolean

Public Property Get UniqueID() As String
    UniqueID = msMyID
End Property

Public Property Get Valid() As Boolean
    Valid = mbIsValid
End Property

Public Property Get FromAddress() As String
    FromAddress = mobjFrom.Describe
End Property

Public Property Get ToAddress() As String
    ToAddress = mobjTo.Describe
End Property

Public Function Describe() As String
    If FromAddress = ToAddress Then
        Describe = FromAddress
    Else
        Describe = FromAddress & "-" & ToAddress
    End If
End Function

Public Function Clone() As clsIPRange
    Dim objClone As clsIPRange
    
    Set objClone = New clsIPRange
    objClone.BuildFromString WriteToString
    Set Clone = objClone
    Set objClone = Nothing
End Function

' ###################
' 1,1.2.3.4 OR
' 1,1.2.3.4-1.2.3.5

Public Sub BuildFromString(sData As String)
    Dim nDashPosn As Integer
    Dim nCommaPosn As Integer
    
    On Error GoTo ErrHandler
    
    nCommaPosn = InStr(1, sData, ",")
    nDashPosn = InStr(1, sData, "-")
    
    If nDashPosn > 0 Then
        Build (Left$(sData, 1) = "1"), Mid$(sData, nCommaPosn + 1, nDashPosn - nCommaPosn - 1), Mid$(sData, nDashPosn + 1)
    Else
        Build (Left$(sData, 1) = "1"), Mid$(sData, nCommaPosn + 1)
    End If
    
    Exit Sub
ErrHandler:
    mbIsValid = False
End Sub

Public Function WriteToString() As String
    WriteToString = IIf(Enabled, "1", "0") & "," & Describe
End Function

Public Sub Build(bEnabled As Boolean, sFrom As String, Optional sTo As String = "")
    On Error GoTo ErrHandler
    
    mbIsValid = False
    Enabled = bEnabled
    If Len(sTo) = 0 Then sTo = sFrom
    
    Set mobjFrom = New clsIP
    mobjFrom.BuildFromString sFrom
    Set mobjTo = New clsIP
    mobjTo.BuildFromString sTo
    
    mbIsValid = True
    
    Exit Sub
ErrHandler:
    'Debug.Print "IP Range: " & Err.Description
End Sub

Public Function IsInRange(sIP As String)
    Dim objTemp As clsIP
    
    On Error GoTo ErrHandler
    
    Set objTemp = New clsIP
    objTemp.BuildFromString sIP
    IsInRange = Compare(mobjFrom, LessThanOrEqual, objTemp) And Compare(mobjTo, GreaterThanOrEqual, objTemp)
    Set objTemp = Nothing
    
    Exit Function
ErrHandler:
    IsInRange = False
    
End Function

Private Sub Class_Initialize()
    msMyID = GetUID
End Sub
