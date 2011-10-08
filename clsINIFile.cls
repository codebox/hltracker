VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFilePath As String

Private Const BUFFER_SIZE = 4096

Private Const APP_OFFSET = 1000
Private Const ERR_NOFILEPATH = vbObjectError + APP_OFFSET + 1
Private Const ERR_BADFILEPATH = vbObjectError + APP_OFFSET + 2
Private Const ERR_CANTWRITE = vbObjectError + APP_OFFSET + 3

Private Const WRITE_SUCCESS = 1

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long


' ########################################################################################################################################################
' PROPERTIES

Public Property Let FilePath(sPath As String)
    If Dir(sPath) = "" Then
        Err.Raise ERR_BADFILEPATH, "clsINIFile.Let_FilePath()", "The specified file does not exist"
    Else
        msFilePath = sPath
    End If
End Property
Public Property Get FilePath() As String
    FilePath = msFilePath
End Property

' ########################################################################################################################################################
' METHODS

' ##################
' WRITE

Public Sub SetINIValue(sSection As String, sKey As String, sValue As String)
' This function writes the key/value into the specified section of the INI file.
' If the section/key does not exist then it will be created.
' If the key already exists then it will be updated.
    Dim lRetVal As Long
    Dim lErrCode As Long

    On Error GoTo errhandler
    
    If msFilePath = "" Then
        Err.Raise ERR_NOFILEPATH, "", "The 'FilePath' property has not been initialised"
    End If

    lRetVal = WritePrivateProfileString(sSection, sKey, sValue, msFilePath)
    
    If Not lRetVal = WRITE_SUCCESS Then
        lErrCode = GetLastError()
        Err.Raise ERR_CANTWRITE, , "Unable to write to the INI file. Error code was" & Str(lErrCode)
    End If
    
    Exit Sub
errhandler:
    Call ReRaiseError("SetINIValue")
    
End Sub

Public Sub SetINISection(sSection As String, dctData As Dictionary)
' This function writes the data contained in the dictionary into the specified section of the INI file.
' If the section does not exist then it will be created.
' Any existing data in the section will be removed.
    Dim lRetVal As Long
    Dim lErrCode As Long
    Dim sData As String
    Dim vntThisKey As Variant
    Dim sThisKey As String
    Dim sThisValue As String

    On Error GoTo errhandler
    
    If msFilePath = "" Then
        Err.Raise ERR_NOFILEPATH, "", "The 'FilePath' property has not been initialised"
    End If

    ' Build up the data string
    For Each vntThisKey In dctData.Keys
        sThisKey = CStr(vntThisKey)
        sThisValue = dctData(sThisKey)
        sData = sData & sThisKey & "=" & sThisValue & vbNullChar
    Next vntThisKey
    sData = sData & vbNullChar

    lRetVal = WritePrivateProfileSection(sSection, sData, msFilePath)
    
    If Not lRetVal = WRITE_SUCCESS Then
        lErrCode = GetLastError()
        Err.Raise ERR_CANTWRITE, , "Unable to write to the INI file. Error code was" & Str(lErrCode)
    End If

    Exit Sub
errhandler:
    Call ReRaiseError("SetINISection")
    
End Sub

' ##################
' READ

Public Function GetINIValue(sSection As String, sKey As String) As String
    Dim sBuffer As String * BUFFER_SIZE
    Dim lValueLength As Long
    Dim sData As String
    
    On Error GoTo errhandler
    
    If msFilePath = "" Then
        Err.Raise ERR_NOFILEPATH, "", "The 'FilePath' property has not been initialised"
    End If
    
    lValueLength = GetPrivateProfileString(sSection, sKey, "", sBuffer, BUFFER_SIZE, msFilePath)
    If lValueLength > 0 Then
        sData = Left$(sBuffer, lValueLength)
    End If

    GetINIValue = sData

    Exit Function
    
errhandler:
    Call ReRaiseError("GetINIValue")
    
End Function

Public Function GetINISection(sSection As String) As Dictionary
    Dim sBuffer As String * BUFFER_SIZE
    Dim lValueLength As Long
    Dim sData As String
    Dim dctData As Dictionary
    Dim vntItem As Variant
    Dim sItem As String
    Dim sKey As String
    Dim sValue As String
    Dim nEqualsPosn As Integer
    Dim arrResults() As String
    Dim bAddingItem As Boolean

    On Error GoTo errhandler
    
    If msFilePath = "" Then
        Err.Raise ERR_NOFILEPATH, "", "The 'FilePath' property has not been initialised"
    End If
    
    Set dctData = New Dictionary

    lValueLength = GetPrivateProfileSection(sSection, sBuffer, BUFFER_SIZE, msFilePath)
    If lValueLength > 1 Then
        sData = Left$(sBuffer, lValueLength - 1)
    End If
    arrResults = Split(sData, vbNullChar)
    
    For Each vntItem In arrResults
        bAddingItem = True
        sItem = CStr(vntItem)
        nEqualsPosn = InStr(1, sItem, "=")
        If nEqualsPosn > 1 Then
            sKey = Left$(sItem, nEqualsPosn - 1)
            sValue = Mid$(sItem, nEqualsPosn + 1)
            dctData.Add sKey, sValue
        End If
NextOne:
        bAddingItem = False
    Next vntItem
    
    Set GetINISection = dctData
    Set dctData = Nothing
    
    Exit Function
    
errhandler:
    If bAddingItem Then
        Resume NextOne
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Public Function GetINISectionList() As Collection
    Dim sBuffer As String * BUFFER_SIZE
    Dim lValueLength As Long
    Dim sData As String
    Dim arrSections() As String
    Dim vntThisSection As Variant
    Dim colSectionNames As Collection

    On Error GoTo errhandler
    
    If msFilePath = "" Then
        Err.Raise ERR_NOFILEPATH, "", "The 'FilePath' property has not been initialised"
    End If

    Set colSectionNames = New Collection

    lValueLength = GetPrivateProfileSectionNames(sBuffer, BUFFER_SIZE, msFilePath)
    If lValueLength > 1 Then
        sData = Left$(sBuffer, lValueLength - 1)
        arrSections = Split(sData, vbNullChar)
        For Each vntThisSection In arrSections
            colSectionNames.Add CStr(vntThisSection)
        Next vntThisSection
    End If

    Set GetINISectionList = colSectionNames

    Exit Function
    
errhandler:
    Call ReRaiseError("GetINISectionList")

End Function

' ########################################################################################################################################################
Private Sub ReRaiseError(sLocation)
    Dim sSource As String
    
    sSource = "clsINIFile." & sLocation & "()" & IIf(Erl > 0, " Line:" & Str(Erl), "")
    
    Err.Raise Err.Number, sSource, Err.Description
    
End Sub
