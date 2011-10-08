VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolFilters As Collection

' ########################################################
' PROPERTIES

Public Property Get Count() As Long
    Count = mcolFilters.Count
End Property

' ########################################################
' METHODS

Public Sub MoveUp(vntKey As Variant)
    Dim nCurrentPosn As Integer
    
    nCurrentPosn = GetCurrentPosn(mcolFilters.Item(vntKey))
    If nCurrentPosn > 1 Then
        MoveToHere mcolFilters.Item(vntKey), nCurrentPosn - 1, nCurrentPosn
    End If
End Sub

Public Sub MoveDown(vntKey As Variant)
    Dim nCurrentPosn As Integer

    nCurrentPosn = GetCurrentPosn(mcolFilters.Item(vntKey))
    If nCurrentPosn < mcolFilters.Count Then
        MoveToHere mcolFilters.Item(vntKey), nCurrentPosn + 1, nCurrentPosn
    End If
End Sub

Private Sub MoveToHere(objFilter As IFilterRule, nNewPosn As Integer, nCurrentPosn As Integer)
    Dim nCount As Integer
    On Error GoTo errhandler
    
    nCount = mcolFilters.Count
    If nCount > 1 Then
        If nNewPosn < 1 Then nNewPosn = 1
        If nCurrentPosn < 1 Then nCurrentPosn = 1
        If nNewPosn > nCount Then nNewPosn = nCount
        If nCurrentPosn > nCount Then nCurrentPosn = nCount
        
        If nCurrentPosn > nNewPosn Then
            mcolFilters.Remove objFilter.UniqueID
            mcolFilters.Add objFilter, objFilter.UniqueID, nNewPosn
        ElseIf nCurrentPosn < nNewPosn Then
            mcolFilters.Remove objFilter.UniqueID
            mcolFilters.Add objFilter, objFilter.UniqueID, , nNewPosn - 1
        End If
    End If
    Exit Sub
errhandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetCurrentPosn(objFilter As IFilterRule) As Integer
    Dim objThisFilter As IFilterRule
    Dim nPosn As Integer
    
    For Each objThisFilter In mcolFilters
        nPosn = nPosn + 1
        If objThisFilter.UniqueID = objFilter.UniqueID Then
            GetCurrentPosn = nPosn
            Exit For
        End If
    Next objThisFilter
    
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolFilters.[_NewEnum]
End Function

Public Function Item(vntKey As Variant) As IFilterRule
    Set Item = mcolFilters.Item(CStr(vntKey))
End Function

Public Sub Add(objFilter As IFilterRule)
    mcolFilters.Add objFilter, objFilter.UniqueID
End Sub

Public Sub Remove(vntKey As Variant)
    mcolFilters.Remove vntKey
End Sub

Private Sub Class_Initialize()
    Set mcolFilters = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolFilters = Nothing
End Sub

