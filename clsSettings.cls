VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DEF_TEXT_COL = vbWhite
Private Const DEF_BACK_COL = 12615680

Private Const DEF_HEIGHT = 3510
Private Const DEF_WIDTH = 6240

Private Const DEF_EXPIRE_INTERVAL = 15
Private Const DEF_MIRROR_INTERVAL = 20
Private Const DEF_MAXLOGSIZE = 5
Private Const DEF_ALERTNEW = True
Private Const DEF_ALERTEXPIRE = True
Private Const DEF_ALERTLIST = True
Private Const DEF_DOHITLOG = False
Private Const DEF_DOSERVERLOG = False
Private Const DEF_DODISPLAYLOG = False
Private Const DEF_BLOCKED_MSG = "You have been banned from using this tracker"

Public Property Get RefuseBlocked() As Boolean
    RefuseBlocked = GetSetting(APP_NAME, "Settings", "RefuseBlocked", True)
End Property
Public Property Let RefuseBlocked(bRefuseBlocked As Boolean)
    SaveSetting APP_NAME, "Settings", "RefuseBlocked", bRefuseBlocked
End Property

Public Property Get BlockedMsg() As String
    BlockedMsg = GetSetting(APP_NAME, "Settings", "BlockedMsg", DEF_BLOCKED_MSG)
End Property
Public Property Let BlockedMsg(sBlockedMsg As String)
    SaveSetting APP_NAME, "Settings", "BlockedMsg", sBlockedMsg
End Property

Public Property Get DoHitLogging() As Boolean
    DoHitLogging = GetSetting(APP_NAME, "Settings", "DoHitLogging", DEF_DOHITLOG)
End Property
Public Property Let DoHitLogging(bDoHitLogging As Boolean)
    SaveSetting APP_NAME, "Settings", "DoHitLogging", bDoHitLogging
End Property

Public Property Get DoServerLogging() As Boolean
    DoServerLogging = GetSetting(APP_NAME, "Settings", "DoServerLogging", DEF_DOSERVERLOG)
End Property
Public Property Let DoServerLogging(bDoServerLogging As Boolean)
    SaveSetting APP_NAME, "Settings", "DoServerLogging", bDoServerLogging
End Property

Public Property Get DoDisplayLogging() As Boolean
    DoDisplayLogging = GetSetting(APP_NAME, "Settings", "DoDisplayLogging", DEF_DODISPLAYLOG)
End Property
Public Property Let DoDisplayLogging(bDoDisplayLogging As Boolean)
    SaveSetting APP_NAME, "Settings", "DoDisplayLogging", bDoDisplayLogging
End Property

Public Property Get ExpireInterval() As Long
    ExpireInterval = GetSetting(APP_NAME, "Settings", "ExpireInterval", DEF_EXPIRE_INTERVAL)
End Property
Public Property Let ExpireInterval(lExpireInterval As Long)
    SaveSetting APP_NAME, "Settings", "ExpireInterval", lExpireInterval
End Property

Public Property Get FormHeight() As Long
    FormHeight = GetSetting(APP_NAME, "Settings", "FormHeight", DEF_HEIGHT)
End Property
Public Property Let FormHeight(lFormHeight As Long)
    SaveSetting APP_NAME, "Settings", "FormHeight", lFormHeight
End Property

Public Property Get FormWidth() As Long
    FormWidth = GetSetting(APP_NAME, "Settings", "FormWidth", DEF_WIDTH)
End Property
Public Property Let FormWidth(lFormWidth As Long)
    SaveSetting APP_NAME, "Settings", "FormWidth", lFormWidth
End Property

Public Property Get MaxLogSize() As Long
    MaxLogSize = CLng(GetSetting(APP_NAME, "Settings", "MaxLogSize", DEF_MAXLOGSIZE))
End Property
Public Property Let MaxLogSize(lMaxLogSize As Long)
    SaveSetting APP_NAME, "Settings", "MaxLogSize", CStr(lMaxLogSize)
End Property

Public Property Get RestrictLogSize() As Boolean
    RestrictLogSize = GetSetting(APP_NAME, "Settings", "RestrictLogSize", True)
End Property
Public Property Let RestrictLogSize(bRestrictLogSize As Boolean)
    SaveSetting APP_NAME, "Settings", "RestrictLogSize", bRestrictLogSize
End Property

' #############################################
' Log Window

' ShowLogTimestamp
Public Property Get ShowLogTimestamp() As Boolean
    ShowLogTimestamp = GetSetting(APP_NAME, "Settings", "ShowLogTimestamp", True)
End Property
Public Property Let ShowLogTimestamp(bShowLogTimestamp As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogTimestamp", bShowLogTimestamp
End Property

' ShowLogServerAdd
Public Property Get ShowLogServerAdd() As Boolean
    ShowLogServerAdd = GetSetting(APP_NAME, "Settings", "ShowLogServerAdd", True)
End Property
Public Property Let ShowLogServerAdd(bShowLogServerAdd As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogServerAdd", bShowLogServerAdd
End Property

' ShowLogServerExpire
Public Property Get ShowLogServerExpire() As Boolean
    ShowLogServerExpire = GetSetting(APP_NAME, "Settings", "ShowLogServerExpire", True)
End Property
Public Property Let ShowLogServerExpire(bShowLogServerExpire As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogServerExpire", bShowLogServerExpire
End Property

' ShowLogServerFiltered
Public Property Get ShowLogServerFiltered() As Boolean
    ShowLogServerFiltered = GetSetting(APP_NAME, "Settings", "ShowLogServerFiltered", True)
End Property
Public Property Let ShowLogServerFiltered(bShowLogServerFiltered As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogServerFiltered", bShowLogServerFiltered
End Property

' ShowLogMirrors
Public Property Get ShowLogMirrors() As Boolean
    ShowLogMirrors = GetSetting(APP_NAME, "Settings", "ShowLogMirrors", True)
End Property
Public Property Let ShowLogMirrors(bShowLogMirrors As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogMirrors", bShowLogMirrors
End Property

' ShowLogBadPass
Public Property Get ShowLogBadPass() As Boolean
    ShowLogBadPass = GetSetting(APP_NAME, "Settings", "ShowLogBadPass", True)
End Property
Public Property Let ShowLogBadPass(bShowLogBadPass As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogBadPass", bShowLogBadPass
End Property

' ShowLogListReq
Public Property Get ShowLogListReq() As Boolean
    ShowLogListReq = GetSetting(APP_NAME, "Settings", "ShowLogListReq", True)
End Property
Public Property Let ShowLogListReq(bShowLogListReq As Boolean)
    SaveSetting APP_NAME, "Settings", "ShowLogListReq", bShowLogListReq
End Property

' #############################################
' MIRRORING...

' FilterDefault <->
Public Property Get FilterDefault() As Boolean
    FilterDefault = GetSetting(APP_NAME, "Settings", "FilterDefault", True)
End Property
Public Property Let FilterDefault(bFilterDefault As Boolean)
    SaveSetting APP_NAME, "Settings", "FilterDefault", bFilterDefault
End Property

Public Property Get Filters() As clsFilterRules
    Dim sFilters As String
    Dim asFilters() As String
    Dim vntFilter As Variant
    Dim objFilters As clsFilterRules
    Dim objFilter As IFilterRule
    Dim sFilterType As String
    Dim bInLoop As Boolean
    Dim bSkip As Boolean
    
    On Error GoTo ErrHandler
    
    sFilters = GetSetting(APP_NAME, "Settings", "Filters", "")
    asFilters = Split(sFilters, ",")
    
    Set objFilters = New clsFilterRules
    
    bInLoop = True
    For Each vntFilter In asFilters
        sFilterType = Left$(CStr(vntFilter), 1)
        bSkip = False
        Select Case sFilterType
            Case FILTER_TYPE_NAME
                Set objFilter = New clsFilterName
                objFilter.BuildFromString CStr(vntFilter)
            Case FILTER_TYPE_DESC
                Set objFilter = New clsFilterDesc
                objFilter.BuildFromString CStr(vntFilter)
            Case FILTER_TYPE_PORT
                Set objFilter = New clsFilterPort
                objFilter.BuildFromString CStr(vntFilter)
            Case FILTER_TYPE_USERS
                Set objFilter = New clsFilterUserCount
                objFilter.BuildFromString CStr(vntFilter)
            Case FILTER_TYPE_IP
                Set objFilter = New clsFilterIP
                objFilter.BuildFromString CStr(vntFilter)
            Case Else
                bSkip = True
        End Select
        If Not bSkip Then objFilters.Add objFilter
NextOne:
    Next vntFilter
    bInLoop = False
    
    Set objFilter = Nothing
    
    Set Filters = objFilters
    Set objFilters = Nothing
    
    Exit Property
ErrHandler:
    If bInLoop Then
        'Debug.Print "ERROR in 'Get Filters()': " & Err.Description
        Resume NextOne
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Property

Public Property Set Filters(colFilters As clsFilterRules)
    Dim sFilters As String
    Dim objFilter As IFilterRule

    For Each objFilter In colFilters
        sFilters = sFilters & IIf(Len(sFilters) > 0, ",", "") & objFilter.WriteToString
    Next objFilter
    SaveSetting APP_NAME, "Settings", "Filters", sFilters
    
End Property

Public Property Get TrackerMirrors() As Collection
    Dim colMirrors As Collection
    Dim vntMirror As Variant
    Dim asMirrors() As String
    Dim sMirrors As String
    Dim objMirror As clsMirroredTracker
    
    Set colMirrors = New Collection
    
    sMirrors = GetSetting(APP_NAME, "Settings", "TrackerMirrors", "")
    'sMirrors = "1tracker.sunwave.com,1dmp.fimble.com,1thereaper.dynip.com,1shxd.ath.cx,1supertracker.kicks-ass.org"
    
    asMirrors = Split(sMirrors, ",")

    For Each vntMirror In asMirrors
        Set objMirror = New clsMirroredTracker
        objMirror.Address = Mid$(CStr(vntMirror), 2)
        objMirror.Enabled = (Left$(vntMirror, 1) = "1")
        colMirrors.Add objMirror
    Next vntMirror
    Set objMirror = Nothing
    
    Set TrackerMirrors = colMirrors
    Set colMirrors = Nothing
End Property
Public Property Set TrackerMirrors(colTrackerMirrors As Collection)
    Dim objMirror As clsMirroredTracker
    Dim sMirrors As String

    For Each objMirror In colTrackerMirrors
        sMirrors = sMirrors & IIf(Len(sMirrors) > 0, ",", "") & IIf(objMirror.Enabled, "1", "0") & objMirror.Address
    Next objMirror

    SaveSetting APP_NAME, "Settings", "TrackerMirrors", sMirrors
End Property

Public Property Get MirrorInterval() As Long
    MirrorInterval = GetSetting(APP_NAME, "Settings", "MirrorInterval", DEF_MIRROR_INTERVAL)
End Property
Public Property Let MirrorInterval(lMirrorInterval As Long)
    SaveSetting APP_NAME, "Settings", "MirrorInterval", lMirrorInterval
End Property

Public Property Get DoMirroring() As Boolean
    DoMirroring = GetSetting(APP_NAME, "Settings", "DoMirroring", False)
End Property
Public Property Let DoMirroring(bDoMirroring As Boolean)
    SaveSetting APP_NAME, "Settings", "DoMirroring", bDoMirroring
End Property

' #############################################

Public Property Get BlockedIPs() As String
    BlockedIPs = GetSetting(APP_NAME, "Settings", "BlockedIPs", "")
End Property
Public Property Let BlockedIPs(sBlockedIPs As String)
    SaveSetting APP_NAME, "Settings", "BlockedIPs", sBlockedIPs
End Property

Public Property Get FakeServers() As String
    FakeServers = GetSetting(APP_NAME, "Settings", "FakeServers", "")
End Property
Public Property Let FakeServers(sFakeServers As String)
    SaveSetting APP_NAME, "Settings", "FakeServers", sFakeServers
End Property

Public Property Get TextColour() As OLE_COLOR
    TextColour = GetSetting(APP_NAME, "Settings", "TextColour", DEF_TEXT_COL)
End Property
Public Property Let TextColour(lColour As OLE_COLOR)
    SaveSetting APP_NAME, "Settings", "TextColour", lColour
End Property

Public Property Get BackColour() As OLE_COLOR
    BackColour = GetSetting(APP_NAME, "Settings", "BackColour", DEF_BACK_COL)
End Property
Public Property Let BackColour(lColour As OLE_COLOR)
    SaveSetting APP_NAME, "Settings", "BackColour", lColour
End Property

Public Property Get AlertNew() As Boolean
    AlertNew = GetSetting(APP_NAME, "Settings", "AlertNew", DEF_ALERTNEW)
End Property
Public Property Let AlertNew(bAlertNew As Boolean)
    SaveSetting APP_NAME, "Settings", "AlertNew", bAlertNew
End Property

Public Property Get AlertExpire() As Boolean
    AlertExpire = GetSetting(APP_NAME, "Settings", "AlertExpire", DEF_ALERTEXPIRE)
End Property
Public Property Let AlertExpire(bAlertExpire As Boolean)
    SaveSetting APP_NAME, "Settings", "AlertExpire", bAlertExpire
End Property

Public Property Get AlertList() As Boolean
    AlertList = GetSetting(APP_NAME, "Settings", "AlertList", DEF_ALERTLIST)
End Property
Public Property Let AlertList(bAlertList As Boolean)
    SaveSetting APP_NAME, "Settings", "AlertList", bAlertList
End Property

Public Property Get Float() As Boolean
    Float = GetSetting(APP_NAME, "Settings", "Float", False)
End Property
Public Property Let Float(bFloat As Boolean)
    SaveSetting APP_NAME, "Settings", "Float", bFloat
End Property

Public Property Get RequirePassword() As Boolean
    RequirePassword = GetSetting(APP_NAME, "Settings", "RequirePassword", False)
End Property
Public Property Let RequirePassword(bRequirePassword As Boolean)
    SaveSetting APP_NAME, "Settings", "RequirePassword", bRequirePassword
End Property

Public Property Get Password() As String
    Password = GetSetting(APP_NAME, "Settings", "Password", "")
End Property
Public Property Let Password(sPasword As String)
    SaveSetting APP_NAME, "Settings", "Password", sPasword
End Property


