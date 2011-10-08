VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D0E26E2B-9AF6-11D6-B8B3-400002012854}#1.0#0"; "HyperLink.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   12495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12495
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   2
      Left            =   1710
      TabIndex        =   116
      Top             =   9765
      Width           =   5640
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   2430
         TabIndex        =   127
         Top             =   270
         Width           =   1050
         Begin VB.OptionButton optMatchTypeDesc 
            Caption         =   "contains"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   69
            Top             =   90
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optMatchTypeDesc 
            Caption         =   "equals"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   70
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.TextBox txtFilterNameDesc 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   73
         Top             =   1440
         Width           =   1635
      End
      Begin VB.OptionButton optActionDesc 
         Caption         =   "Block"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   67
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optActionDesc 
         Caption         =   "List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   68
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox txtExpressionDesc 
         Height          =   285
         Left            =   3825
         MaxLength       =   128
         TabIndex        =   71
         Top             =   405
         Width           =   1635
      End
      Begin VB.CheckBox chkMatchCaseDesc 
         Caption         =   "Match Case"
         Height          =   240
         Left            =   180
         TabIndex        =   72
         Top             =   1035
         Width           =   1185
      End
      Begin VB.Label Label17 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   126
         Top             =   1485
         Width           =   1770
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "servers if the description"
         Height          =   420
         Left            =   1170
         TabIndex        =   125
         Top             =   450
         Width           =   1005
      End
   End
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   1
      Left            =   2655
      TabIndex        =   115
      Top             =   9135
      Width           =   5640
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   2430
         TabIndex        =   124
         Top             =   270
         Width           =   1050
         Begin VB.OptionButton optMatchTypeName 
            Caption         =   "equals"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   63
            Top             =   360
            Width           =   870
         End
         Begin VB.OptionButton optMatchTypeName 
            Caption         =   "contains"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   62
            Top             =   90
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.TextBox txtFilterNameName 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   66
         Top             =   1440
         Width           =   1635
      End
      Begin VB.OptionButton optActionName 
         Caption         =   "Block"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   60
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optActionName 
         Caption         =   "List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   61
         Top             =   630
         Width           =   690
      End
      Begin VB.CheckBox chkMatchCaseName 
         Caption         =   "Match Case"
         Height          =   240
         Left            =   180
         TabIndex        =   65
         Top             =   1035
         Width           =   1185
      End
      Begin VB.TextBox txtExpressionName 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   64
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   123
         Top             =   1485
         Width           =   1770
      End
      Begin VB.Label Label15 
         Caption         =   "servers if the name"
         Height          =   240
         Left            =   945
         TabIndex        =   122
         Top             =   495
         Width           =   1455
      End
   End
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   6
      Left            =   7515
      TabIndex        =   138
      Top             =   7605
      Width           =   5640
      Begin VB.OptionButton optSpecial 
         Caption         =   "List all servers"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   58
         Top             =   855
         Width           =   1545
      End
      Begin VB.OptionButton optSpecial 
         Caption         =   "Block all servers"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   57
         Top             =   585
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3465
         MaxLength       =   32
         TabIndex        =   59
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label33 
         Caption         =   "This rule will be applied if no other rule matches:"
         Height          =   195
         Left            =   225
         TabIndex        =   158
         Top             =   315
         Width           =   3435
      End
      Begin VB.Label Label24 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   139
         Top             =   1485
         Width           =   1770
      End
   End
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   5
      Left            =   1530
      TabIndex        =   133
      Top             =   10035
      Width           =   5640
      Begin VB.TextBox txtFilterNameUsers 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   80
         Top             =   1440
         Width           =   1635
      End
      Begin VB.OptionButton optActionUsers 
         Caption         =   "Block"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   74
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optActionUsers 
         Caption         =   "List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   75
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox txtRangeUsers2 
         Height          =   285
         Left            =   3825
         MaxLength       =   6
         TabIndex        =   79
         Top             =   945
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtRangeUsers1 
         Height          =   285
         Left            =   3825
         MaxLength       =   6
         TabIndex        =   78
         Top             =   405
         Width           =   1635
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   2115
         TabIndex        =   134
         Top             =   270
         Width           =   1365
         Begin VB.OptionButton optMatchTypeUsers 
            Caption         =   "is in the range"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   77
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton optMatchTypeUsers 
            Caption         =   "equals"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   76
            Top             =   90
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Label Label25 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   137
         Top             =   1485
         Width           =   1770
      End
      Begin VB.Label lblUsersTo 
         Caption         =   "to"
         Height          =   195
         Left            =   4500
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "servers if the number of users"
         Height          =   420
         Left            =   900
         TabIndex        =   135
         Top             =   405
         Width           =   1230
      End
   End
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   4
      Left            =   1215
      TabIndex        =   121
      Top             =   10125
      Width           =   5640
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   2115
         TabIndex        =   129
         Top             =   270
         Width           =   1365
         Begin VB.OptionButton optMatchTypePort 
            Caption         =   "equals"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   83
            Top             =   90
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optMatchTypePort 
            Caption         =   "is in the range"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   84
            Top             =   360
            Width           =   1410
         End
      End
      Begin VB.TextBox txtRangePort1 
         Height          =   285
         Left            =   3825
         MaxLength       =   6
         TabIndex        =   85
         Top             =   405
         Width           =   1635
      End
      Begin VB.TextBox txtRangePort2 
         Height          =   285
         Left            =   3825
         MaxLength       =   6
         TabIndex        =   86
         Top             =   945
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.OptionButton optActionPort 
         Caption         =   "List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   82
         Top             =   630
         Width           =   690
      End
      Begin VB.OptionButton optActionPort 
         Caption         =   "Block"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   81
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtFilterNamePort 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   87
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label23 
         Caption         =   "servers if the port"
         Height          =   285
         Left            =   900
         TabIndex        =   132
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label lblPortTo 
         Caption         =   "to"
         Height          =   195
         Left            =   4500
         TabIndex        =   131
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   130
         Top             =   1485
         Width           =   1770
      End
   End
   Begin VB.Frame fraFilterTypes 
      Height          =   2040
      Index           =   3
      Left            =   2070
      TabIndex        =   117
      Top             =   9990
      Width           =   5640
      Begin VB.TextBox txtFilterNameIP 
         Height          =   285
         Left            =   3825
         MaxLength       =   32
         TabIndex        =   94
         Top             =   1440
         Width           =   1635
      End
      Begin VB.OptionButton optActionIP 
         Caption         =   "Block"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   88
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optActionIP 
         Caption         =   "List"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   89
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox txtRange2 
         Height          =   285
         Left            =   3825
         MaxLength       =   15
         TabIndex        =   93
         Top             =   945
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtRange1 
         Height          =   285
         Left            =   3825
         MaxLength       =   15
         TabIndex        =   92
         Top             =   405
         Width           =   1635
      End
      Begin VB.Frame fraIPNumMatch 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   2115
         TabIndex        =   119
         Top             =   270
         Width           =   1365
         Begin VB.OptionButton optMatchTypeIP 
            Caption         =   "is in the range"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   91
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton optMatchTypeIP 
            Caption         =   "equals"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   90
            Top             =   90
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Label Label20 
         Caption         =   "The name of this rule is:"
         Height          =   195
         Left            =   1575
         TabIndex        =   128
         Top             =   1485
         Width           =   1770
      End
      Begin VB.Label lblRange 
         Caption         =   "to"
         Height          =   195
         Left            =   4500
         TabIndex        =   120
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "servers if the IP address"
         Height          =   420
         Left            =   990
         TabIndex        =   118
         Top             =   405
         Width           =   1140
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   7
      Left            =   7695
      TabIndex        =   114
      Top             =   1845
      Width           =   6090
      Begin MSComctlLib.TabStrip tbsFilterTypes 
         Height          =   2490
         Left            =   135
         TabIndex        =   8
         Top             =   2430
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4392
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Name"
               Key             =   "Name"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Description"
               Key             =   "Description"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "IP"
               Key             =   "IP"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Port"
               Key             =   "Port"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "User Count"
               Key             =   "User Count"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Default"
               Key             =   "Default"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDeleteFilter 
         Height          =   375
         Left            =   5310
         Picture         =   "frmSettings.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Delete Filter"
         Top             =   2025
         Width           =   645
      End
      Begin VB.CommandButton cmdAddFilter 
         Height          =   375
         Left            =   5310
         Picture         =   "frmSettings.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add Filter"
         Top             =   1620
         Width           =   645
      End
      Begin VB.CommandButton cmdFilterUp 
         Height          =   375
         Left            =   5310
         Picture         =   "frmSettings.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   645
      End
      Begin VB.CommandButton cmdFilterDown 
         Height          =   375
         Left            =   5310
         Picture         =   "frmSettings.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   855
         Width           =   645
      End
      Begin MSComctlLib.ListView lvwFilters 
         Height          =   2220
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "tracker"
            Text            =   "Rule Description"
            Object.Width           =   8915
         EndProperty
      End
   End
   Begin VB.Timer tmrExpire 
      Interval        =   1000
      Left            =   5490
      Top             =   6120
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   2
      Left            =   7965
      TabIndex        =   113
      Top             =   2070
      Width           =   6090
      Begin VB.Frame Frame8 
         Caption         =   "Clients"
         Height          =   3435
         Left            =   90
         TabIndex        =   154
         Top             =   1530
         Width           =   5910
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   1620
            TabIndex        =   159
            Top             =   1710
            Width           =   1140
            Begin VB.OptionButton optClientIP 
               Caption         =   "Range"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   47
               Top             =   225
               Width           =   1005
            End
            Begin VB.OptionButton optClientIP 
               Caption         =   "Single IP"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   1005
            End
         End
         Begin VB.TextBox txtBlockMessage 
            Enabled         =   0   'False
            Height          =   285
            Left            =   135
            TabIndex        =   52
            Top             =   3015
            Width           =   5685
         End
         Begin VB.OptionButton optBlockOption 
            Caption         =   "Display a message:"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   51
            Top             =   2700
            Width           =   1680
         End
         Begin VB.OptionButton optBlockOption 
            Caption         =   "Refuse the connection"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   50
            Top             =   2475
            Width           =   2310
         End
         Begin VB.TextBox txtClientIPTo 
            Height          =   285
            Left            =   4500
            TabIndex        =   49
            Top             =   1710
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox txtClientIPFrom 
            Height          =   285
            Left            =   2880
            TabIndex        =   48
            Top             =   1710
            Width           =   1320
         End
         Begin VB.CommandButton cmdDeleteClient 
            Height          =   375
            Left            =   5175
            Picture         =   "frmSettings.frx":0D08
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Delete Item"
            Top             =   855
            Width           =   645
         End
         Begin VB.CommandButton cmdAddClient 
            Height          =   375
            Left            =   5175
            Picture         =   "frmSettings.frx":104A
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Add Item"
            Top             =   450
            Width           =   645
         End
         Begin MSComctlLib.ListView lvwClients 
            Height          =   1455
            Left            =   90
            TabIndex        =   43
            Top             =   180
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "tracker"
               Text            =   "Client Address"
               Object.Width           =   8677
            EndProperty
         End
         Begin VB.Label Label32 
            Caption         =   "When a blocked client requests the list..."
            Height          =   240
            Left            =   135
            TabIndex        =   157
            Top             =   2250
            Width           =   2940
         End
         Begin VB.Label lblClientTo 
            Caption         =   "to"
            Height          =   195
            Left            =   4275
            TabIndex        =   156
            Top             =   1755
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label31 
            Caption         =   "Block requests from"
            Height          =   240
            Left            =   90
            TabIndex        =   155
            Top             =   1710
            Width           =   1545
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Servers"
         Height          =   1365
         Left            =   90
         TabIndex        =   150
         Top             =   225
         Width           =   5910
         Begin VB.TextBox txtExpire 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            IMEMode         =   3  'DISABLE
            Left            =   1695
            MaxLength       =   3
            TabIndex        =   42
            Top             =   915
            Width           =   495
         End
         Begin VB.CheckBox chkReqPass 
            Caption         =   "Require password from servers"
            Height          =   255
            Left            =   255
            TabIndex        =   40
            Top             =   270
            Width           =   3075
         End
         Begin VB.TextBox txtPass 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            IMEMode         =   3  'DISABLE
            Left            =   1695
            MaxLength       =   255
            PasswordChar    =   "*"
            TabIndex        =   41
            Top             =   570
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "minutes."
            Height          =   195
            Left            =   2295
            TabIndex        =   153
            Top             =   930
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Servers expire after"
            Height          =   195
            Left            =   135
            TabIndex        =   152
            Top             =   930
            Width           =   1515
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   195
            Left            =   615
            TabIndex        =   151
            Top             =   585
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   6
      Left            =   1800
      TabIndex        =   111
      Top             =   3915
      Width           =   6090
      Begin VB.CheckBox chkDoMirrors 
         Caption         =   "Enable Mirroring"
         Height          =   240
         Left            =   1215
         TabIndex        =   27
         Top             =   4140
         Width           =   2085
      End
      Begin VB.CommandButton cmdMirrorNow 
         Caption         =   "Mirror Now"
         Height          =   375
         Left            =   1170
         TabIndex        =   26
         Top             =   3600
         Width           =   1050
      End
      Begin VB.VScrollBar scrInterval 
         Height          =   375
         LargeChange     =   5
         Left            =   1575
         Max             =   5
         Min             =   90
         SmallChange     =   5
         TabIndex        =   149
         Top             =   3105
         Value           =   5
         Width           =   285
      End
      Begin MSComctlLib.ListView lvwMirrors 
         Height          =   2220
         Left            =   90
         TabIndex        =   21
         Top             =   225
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3916
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "tracker"
            Text            =   "Tracker Address"
            Object.Width           =   10185
         EndProperty
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1170
         TabIndex        =   22
         Top             =   2655
         Width           =   1860
      End
      Begin VB.CommandButton cmddelMirror 
         Height          =   375
         Left            =   3870
         Picture         =   "frmSettings.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Delete Mirror"
         Top             =   2565
         Width           =   645
      End
      Begin VB.CommandButton cmdAddMirror 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3150
         Picture         =   "frmSettings.frx":16CE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add Mirror"
         Top             =   2565
         Width           =   645
      End
      Begin VB.Label lblMirrorInterval 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99"
         Height          =   285
         Left            =   1170
         TabIndex        =   25
         Top             =   3150
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "minutes"
         Height          =   240
         Left            =   1980
         TabIndex        =   148
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Mirror every"
         Height          =   240
         Left            =   180
         TabIndex        =   147
         Top             =   3150
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   195
         Left            =   405
         TabIndex        =   112
         Top             =   2700
         Width           =   645
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   5
      Left            =   720
      TabIndex        =   105
      Top             =   3600
      Width           =   6090
      Begin VB.CommandButton cmdNew 
         Height          =   375
         Left            =   2880
         Picture         =   "frmSettings.frx":1A10
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "New Server"
         Top             =   4350
         Width           =   645
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1515
         MaxLength       =   255
         TabIndex        =   29
         Top             =   2070
         Width           =   4455
      End
      Begin VB.TextBox txtDescription 
         Height          =   705
         Left            =   1515
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2430
         Width           =   4455
      End
      Begin VB.TextBox txtUsers 
         Height          =   285
         Left            =   1515
         MaxLength       =   5
         TabIndex        =   31
         Top             =   3210
         Width           =   1155
      End
      Begin VB.TextBox txtIP1 
         Height          =   285
         Left            =   1485
         MaxLength       =   3
         TabIndex        =   32
         Top             =   3570
         Width           =   495
      End
      Begin VB.TextBox txtIP2 
         Height          =   285
         Left            =   2175
         MaxLength       =   3
         TabIndex        =   33
         Top             =   3570
         Width           =   495
      End
      Begin VB.TextBox txtIP3 
         Height          =   285
         Left            =   2775
         MaxLength       =   3
         TabIndex        =   34
         Top             =   3570
         Width           =   495
      End
      Begin VB.TextBox txtIP4 
         Height          =   285
         Left            =   3375
         MaxLength       =   3
         TabIndex        =   35
         Top             =   3570
         Width           =   495
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1515
         MaxLength       =   5
         TabIndex        =   36
         Top             =   3930
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpdate 
         Height          =   375
         Left            =   1515
         Picture         =   "frmSettings.frx":1D52
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Update Server"
         Top             =   4350
         Width           =   645
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   375
         Left            =   2205
         Picture         =   "frmSettings.frx":2094
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Delete Server"
         Top             =   4350
         Width           =   645
      End
      Begin MSComctlLib.ListView lvwServers 
         Height          =   1815
         Left            =   90
         TabIndex        =   28
         Top             =   180
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2485
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7779
         EndProperty
      End
      Begin VB.Label Label14 
         Caption         =   "Name:"
         Height          =   255
         Left            =   180
         TabIndex        =   110
         Top             =   2115
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Description:"
         Height          =   255
         Left            =   135
         TabIndex        =   109
         Top             =   2490
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "User Count:"
         Height          =   255
         Left            =   135
         TabIndex        =   108
         Top             =   3270
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "IP Address"
         Height          =   255
         Left            =   135
         TabIndex        =   107
         Top             =   3630
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Port Number:"
         Height          =   255
         Left            =   135
         TabIndex        =   106
         Top             =   3990
         Width           =   1095
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   8
      Left            =   7560
      TabIndex        =   104
      Top             =   1575
      Width           =   6090
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   2100
         Picture         =   "frmSettings.frx":23D6
         ScaleHeight     =   645
         ScaleWidth      =   1965
         TabIndex        =   167
         Top             =   495
         Width           =   1995
      End
      Begin prjHyperLink.HyperLink HyperLink1 
         Height          =   195
         Left            =   2250
         TabIndex        =   160
         Top             =   3510
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   344
         LinkText        =   "rob@codebox.no-ip.net"
         LinkTarget      =   "mailto:rob@codebox.no-ip.net"
         UnderLineLink   =   -1  'True
         NewLinkColour   =   16711680
         OldLinkColour   =   8388736
         ErrLinkColour   =   255
         DisLinkColour   =   8421504
         LinkStatus      =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHyperLink.HyperLink HyperLink2 
         Height          =   195
         Left            =   2235
         TabIndex        =   161
         Top             =   1350
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   344
         LinkText        =   "http://codebox.no-ip.net"
         LinkTarget      =   "http://codebox.no-ip.net"
         UnderLineLink   =   -1  'True
         NewLinkColour   =   16711680
         OldLinkColour   =   8388736
         ErrLinkColour   =   255
         DisLinkColour   =   8421504
         LinkStatus      =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblUpTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2430
         TabIndex        =   166
         Top             =   4455
         Width           =   1455
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1455
         TabIndex        =   165
         Top             =   2055
         Width           =   3435
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Codebox Tracker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2055
         TabIndex        =   164
         Top             =   1695
         Width           =   2235
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "This Hotline Tracker is 'thank-you-ware' - if you like it then email the author to say thanks"
         Height          =   435
         Left            =   1350
         TabIndex        =   163
         Top             =   2565
         Width           =   3555
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Up-Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   162
         Top             =   4140
         Width           =   1275
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   4
      Left            =   1125
      TabIndex        =   103
      Top             =   1260
      Width           =   6090
      Begin VB.Frame Frame6 
         Caption         =   "Log Window"
         Height          =   2625
         Left            =   90
         TabIndex        =   145
         Top             =   2340
         Width           =   5910
         Begin VB.CheckBox chkLogBadPassword 
            Caption         =   "Show failed logins"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   1665
            Width           =   2580
         End
         Begin VB.CheckBox chkLogFilter 
            Caption         =   "Show servers being filtered out"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1440
            Width           =   2580
         End
         Begin VB.CheckBox chkLogListing 
            Caption         =   "Show listing requests"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   1215
            Width           =   2310
         End
         Begin VB.CheckBox chkLogMirror 
            Caption         =   "Show mirror activity"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   990
            Width           =   2310
         End
         Begin VB.CheckBox chkLogServerExpire 
            Caption         =   "Show servers expiring"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   765
            Width           =   2310
         End
         Begin VB.CheckBox chkLogAddServer 
            Caption         =   "Show servers checking in"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   540
            Width           =   2310
         End
         Begin VB.CheckBox chkLogTimeStamp 
            Caption         =   "Show timestamps against entries"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   315
            Width           =   2760
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Log Files"
         Height          =   2040
         Left            =   90
         TabIndex        =   140
         Top             =   180
         Width           =   5910
         Begin VB.CheckBox chkDisplayLogging 
            Caption         =   "Save Log Window Entries"
            Height          =   255
            Left            =   180
            TabIndex        =   11
            ToolTipText     =   "Check this box to record server information into a file"
            Top             =   810
            Width           =   2205
         End
         Begin VB.CheckBox chkServerLogging 
            Caption         =   "Enable Server Logging"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            ToolTipText     =   "Check this box to record server information into a file"
            Top             =   555
            Width           =   1935
         End
         Begin VB.CheckBox chkHitLogging 
            Caption         =   "Enable Hit Logging"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            ToolTipText     =   "Check this box to record listing requests into a file"
            Top             =   315
            Width           =   1815
         End
         Begin VB.CheckBox chkTrimLogs 
            Caption         =   "Restrict log sizes"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   1215
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtLogMax 
            Height          =   285
            Left            =   1755
            TabIndex        =   13
            Top             =   1500
            Visible         =   0   'False
            Width           =   735
         End
         Begin prjHyperLink.HyperLink lnkViewHitLog 
            Height          =   195
            Left            =   2655
            TabIndex        =   141
            Top             =   345
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   344
            LinkText        =   "View Hit Log"
            LinkTarget      =   ""
            UnderLineLink   =   -1  'True
            NewLinkColour   =   16711680
            OldLinkColour   =   8388736
            ErrLinkColour   =   255
            DisLinkColour   =   8421504
            LinkStatus      =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjHyperLink.HyperLink lnkViewServerLog 
            Height          =   195
            Left            =   2655
            TabIndex        =   142
            Top             =   585
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   344
            LinkText        =   "View Server Log"
            LinkTarget      =   ""
            UnderLineLink   =   -1  'True
            NewLinkColour   =   16711680
            OldLinkColour   =   8388736
            ErrLinkColour   =   255
            DisLinkColour   =   8421504
            LinkStatus      =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjHyperLink.HyperLink lnkViewWindowLog 
            Height          =   195
            Left            =   2655
            TabIndex        =   146
            Top             =   840
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   344
            LinkText        =   "View Window Log"
            LinkTarget      =   ""
            UnderLineLink   =   -1  'True
            NewLinkColour   =   16711680
            OldLinkColour   =   8388736
            ErrLinkColour   =   255
            DisLinkColour   =   8421504
            LinkStatus      =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label27 
            Caption         =   "Keep logs below"
            Height          =   240
            Left            =   450
            TabIndex        =   144
            Top             =   1530
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label28 
            Caption         =   "Mb"
            Height          =   240
            Left            =   2610
            TabIndex        =   143
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   3
      Left            =   1395
      TabIndex        =   101
      Top             =   2295
      Width           =   6090
      Begin VB.CheckBox chkAlertNewServer 
         Alignment       =   1  'Right Justify
         Caption         =   "...a new server checks in"
         Height          =   195
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   2295
      End
      Begin VB.CheckBox chkAlertRequest 
         Alignment       =   1  'Right Justify
         Caption         =   "...someone requests the list"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   750
         Width           =   2295
      End
      Begin VB.CheckBox chkAlertExpire 
         Alignment       =   1  'Right Justify
         Caption         =   "...a server expires"
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   510
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Alert me when..."
         Height          =   195
         Left            =   135
         TabIndex        =   102
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame fraSettings 
      Height          =   5100
      Index           =   1
      Left            =   540
      TabIndex        =   98
      Top             =   2160
      Width           =   6090
      Begin VB.CommandButton cmdColour1 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   315
         Width           =   315
      End
      Begin VB.CommandButton cmdColour2 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   54
         Top             =   675
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1740
         MultiLine       =   -1  'True
         TabIndex        =   55
         Text            =   "frmSettings.frx":3B0D
         Top             =   315
         Width           =   4155
      End
      Begin VB.CheckBox chkFloat 
         Caption         =   "Always on top"
         Height          =   255
         Left            =   1755
         TabIndex        =   56
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Text Colour:"
         Height          =   195
         Left            =   180
         TabIndex        =   100
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Back Colour:"
         Height          =   195
         Left            =   180
         TabIndex        =   99
         Top             =   675
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip tbsSettings 
      Height          =   5595
      Left            =   45
      TabIndex        =   97
      Top             =   45
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Appearance"
            Key             =   "Appearance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Connections"
            Key             =   "Servers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alerts"
            Key             =   "Alerts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Logs"
            Key             =   "Logs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fake Servers"
            Key             =   "Fakes Servers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mirrors"
            Key             =   "Mirrors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filters"
            Key             =   "Filters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   6780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5445
      TabIndex        =   96
      Top             =   5715
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4500
      TabIndex        =   95
      Top             =   5715
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3660
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MIN_EXPIRE = 1
Private Const MAX_EXPIRE = 120
Private Const MOD_NAME = "frmSettings"
Private mbLoading As Boolean
Private mcolMirrors As Collection
Private mbFiltersChanged As Boolean
Private mobjTempBlockedIPs As clsIPRanges

Private WithEvents mobjServerGrid As clsServerGrid
Attribute mobjServerGrid.VB_VarHelpID = -1

Private Const FRA_VOFFSET = 315
Private Const FRA_HOFFSET = 90
Private Const FRA_HPOS_ACTION = 100
Private Const FRA_VPOS_ACTION = 200
Private Const FRA_HPOS_MATCH = 2430
Private Const FRA_VPOS_MATCH = 200
Private Const TXT_HPOS_EXPR = 3465
Private Const TXT_VPOS_EXPR = 405
Private Const CHK_HPOS_CASE = 1305
Private Const CHK_VPOS_CASE = 225
Private Const FRA_HPOS_RULE = 2430
Private Const FRA_VPOS_RULE = 1395
Private Const FRA_HPOS_NUM = 2430
Private Const FRA_VPOS_NUM = 270
Private Const FRM_HDIFF = 170
Private Const FRM_VDIFF = 1005


Public Property Let InitialFrame(nFrameNum As Integer)
    On Error GoTo errhandler
    Set tbsSettings.SelectedItem = tbsSettings.Tabs(nFrameNum)
    ShowFrame nFrameNum, fraSettings
    Exit Property
errhandler:
    ' ignore
End Property

' ################################
' Filters

Private Sub RefreshFilterList()
    Dim objFilter As IFilterRule
    Dim lsiThisFilter As ListItem
    
    lvwFilters.ListItems.Clear
    For Each objFilter In mobjFilters
        Set lsiThisFilter = lvwFilters.ListItems.Add(, objFilter.UniqueID, objFilter.Description)
        lsiThisFilter.Checked = objFilter.Enabled
    Next objFilter
    mbFiltersChanged = True
End Sub

Private Sub chkDoMirrors_Click()
    cmdMirrorNow.Enabled = (chkDoMirrors.Value = vbChecked)
End Sub

Private Sub chkTrimLogs_Click()
    On Error GoTo errhandler

    txtLogMax.Enabled = (chkTrimLogs.Value = vbChecked)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "chkTrimLogs_Click", Err.Description, Erl
    
End Sub

Private Sub cmdAddClient_Click()
    Dim objNew As clsIPRange
    Dim sIP1 As String
    Dim sIP2 As String
    
    On Error GoTo errhandler
    sIP1 = txtClientIPFrom.Text
    
    If Len(sIP1) > 0 Then
        If txtClientIPTo.Visible Then
            sIP2 = txtClientIPTo.Text
        Else
            sIP2 = sIP1
        End If
        Set objNew = New clsIPRange
        objNew.Build True, sIP1, sIP2
        If objNew.Valid Then
            mobjTempBlockedIPs.Add objNew
        Else
            Err.Raise vbObjectError, , "Invalid address/s specified"
        End If
        Set objNew = Nothing
        LoadBlockedIPs
    End If
    
    Exit Sub
errhandler:
    If Err.Number < 0 Then
        MsgBox Err.Description
    Else
        ErrorReport Err.Number, MOD_NAME, "cmdAddClient_Click", Err.Description, Erl
    End If
End Sub

Private Sub cmdDeleteClient_Click()
    Dim lviSelected As ListItem
    
    On Error GoTo errhandler
    
    Set lviSelected = lvwClients.SelectedItem
    If Not lviSelected Is Nothing Then
        mobjTempBlockedIPs.Remove lviSelected.Key
        LoadBlockedIPs
    End If
    Set lviSelected = Nothing
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdDeleteClient_Click", Err.Description, Erl
    
End Sub

Private Sub cmdFilterDown_Click()
    Dim sSelected As String
    
    On Error GoTo errhandler
    
    If Not lvwFilters.SelectedItem Is Nothing Then
        sSelected = lvwFilters.SelectedItem.Key
        If Len(sSelected) > 0 Then
            mobjFilters.MoveDown sSelected
        End If
        RefreshFilterList
        lvwFilters.ListItems(sSelected).Selected = True
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdFilterDown_Click", Err.Description, Erl
    
End Sub

Private Sub cmdFilterUp_Click()
    Dim sSelected As String
    
    On Error GoTo errhandler
    
    If Not lvwFilters.SelectedItem Is Nothing Then
        sSelected = lvwFilters.SelectedItem.Key
        If Len(sSelected) > 0 Then
            mobjFilters.MoveUp sSelected
        End If
        RefreshFilterList
        lvwFilters.ListItems(sSelected).Selected = True
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdFilterUp_Click", Err.Description, Erl
    
End Sub

Private Sub cmdMirrorNow_Click()
    On Error GoTo errhandler

    frmMain.DoMirror
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdMirrorNow_Click", Err.Description, Erl
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.CallbackSetInitFrame tbsSettings.SelectedItem.Index
End Sub

Private Sub lvwClients_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errhandler

    mobjTempBlockedIPs.Item(Item.Key).Enabled = (Item.Checked)
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "lvwClients_ItemCheck", Err.Description, Erl
    
End Sub

Private Sub lvwFilters_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim objThisFilter As IFilterRule
    
    On Error GoTo errhandler
    
    Set objThisFilter = mobjFilters.Item(Item.Key)
    objThisFilter.Enabled = Item.Checked
    Set objThisFilter = Nothing
    mbFiltersChanged = True
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "lvwFilters_ItemCheck", Err.Description, Erl
    
End Sub

Private Sub cmdDeleteFilter_Click()
    Dim lsiThisFilter As ListItem

    On Error GoTo errhandler

    Set lsiThisFilter = lvwFilters.SelectedItem
    If Not lsiThisFilter Is Nothing Then
        mobjFilters.Remove lsiThisFilter.Key
    End If
    RefreshFilterList
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdDeleteFilter_Click", Err.Description, Erl
    
End Sub

Private Sub cmdAddFilter_Click()
    Dim objFilterName As clsFilterName
    Dim objFilterDesc As clsFilterDesc
    Dim objFilterPort As clsFilterPort
    Dim objFilterUsers As clsFilterUserCount
    Dim objFilterIP As clsFilterIP

    On Error GoTo errhandler

    Select Case tbsFilterTypes.SelectedItem.Index
        Case 1
            If Len(txtExpressionName.Text) = 0 Then
                txtExpressionName.SetFocus
                Err.Raise vbObjectError, , "You need to enter a value to use in the filter"
            End If
            Set objFilterName = New clsFilterName
            objFilterName.Build txtExpressionName.Text, (chkMatchCaseName.Value = vbChecked), _
                                (optActionName(0).Value), (optMatchTypeName(1).Value), True, txtFilterNameName.Text
            mobjFilters.Add objFilterName
            Set objFilterName = Nothing
        Case 2
            If Len(txtExpressionDesc.Text) = 0 Then
                txtExpressionDesc.SetFocus
                Err.Raise vbObjectError, , "You need to enter a value to use in the filter"
            End If
            Set objFilterDesc = New clsFilterDesc
            objFilterDesc.Build txtExpressionDesc.Text, (chkMatchCaseDesc.Value = vbChecked), _
                                (optActionDesc(0).Value), (optMatchTypeDesc(1).Value), True, txtFilterNameDesc.Text
            mobjFilters.Add objFilterDesc
            Set objFilterDesc = Nothing
        Case 3
            If Len(txtRange1.Text) = 0 Then
                txtRange1.SetFocus
                Err.Raise vbObjectError, , "You need to enter a value to use in the filter"
            End If
            If Len(txtRange2.Text) = 0 And txtRange2.Visible Then
                txtRange2.SetFocus
                Err.Raise vbObjectError, , "You need to enter a second value to define an address range"
            End If
            Set objFilterIP = New clsFilterIP
            objFilterIP.Build txtRange1.Text, IIf(txtRange2.Visible, txtRange2.Text, txtRange1.Text), _
                                optActionIP(0).Value, True, txtFilterNameIP.Text
            mobjFilters.Add objFilterIP
            Set objFilterIP = Nothing
        Case 4
            If Len(txtRangePort1.Text) = 0 Then
                txtRangePort1.SetFocus
                Err.Raise vbObjectError, , "You need to enter a value to use in the filter"
            End If
            If Len(txtRangePort2.Text) = 0 And txtRangePort2.Visible Then
                txtRangePort2.SetFocus
                Err.Raise vbObjectError, , "You need to enter a second value to define a range"
            End If
            Set objFilterPort = New clsFilterPort
            objFilterPort.Build txtRangePort1.Text, IIf(txtRangePort2.Visible, txtRangePort2.Text, _
                                txtRangePort1.Text), optActionPort(0).Value, True, txtFilterNamePort.Text
            mobjFilters.Add objFilterPort
            Set objFilterPort = Nothing
        Case 5
            If Len(txtRangeUsers1.Text) = 0 Then
                txtRangeUsers1.SetFocus
                Err.Raise vbObjectError, , "You need to enter a value to use in the filter"
            End If
            If Len(txtRangeUsers2.Text) = 0 And txtRangeUsers2.Visible Then
                txtRangeUsers2.SetFocus
                Err.Raise vbObjectError, , "You need to enter a second value to define a range"
            End If
            Set objFilterUsers = New clsFilterUserCount
            objFilterUsers.Build txtRangeUsers1.Text, IIf(txtRangeUsers2.Visible, txtRangeUsers2.Text, _
                                txtRangeUsers1.Text), optActionUsers(0).Value, True, txtFilterNameUsers.Text
            mobjFilters.Add objFilterUsers
            Set objFilterUsers = Nothing
        Case Else
            ' Ignore
    End Select
    RefreshFilterList
    
    Exit Sub
errhandler:
    MsgBox "Unable to add the new rule. Ensure that this filter does not duplicate one already in the list", vbOKOnly + vbExclamation, APP_NAME
End Sub

' ################################
' Mirrors Tab

Private Sub RefreshMirrorList()
    Dim objMirror As clsMirroredTracker
    Dim lsiThisMirror As ListItem
    
    lvwMirrors.ListItems.Clear
    For Each objMirror In mcolMirrors
        Set lsiThisMirror = lvwMirrors.ListItems.Add(, objMirror.Address, objMirror.Address)
        lsiThisMirror.Checked = objMirror.Enabled
    Next objMirror
    
End Sub

Private Sub cmdAddMirror_Click()
    Dim objNewMirror As clsMirroredTracker
    
    On Error GoTo errhandler
    
    Set objNewMirror = New clsMirroredTracker
    objNewMirror.Address = txtAddress.Text
    objNewMirror.Enabled = True
    mcolMirrors.Add objNewMirror, objNewMirror.Address
    txtAddress.Text = ""
    txtAddress.SetFocus
    RefreshMirrorList
    Set objNewMirror = Nothing
    
    Exit Sub
errhandler:
    If Err.Number = 457 Then
        MsgBox "A mirror with this address is already present in the grid"
        txtAddress.SetFocus
    Else
        ErrorReport Err.Number, MOD_NAME, "cmdAddMirror_Click", Err.Description, Erl
    End If
End Sub


Private Sub cmddelMirror_Click()
    Dim lsiThisMirror As ListItem

    On Error GoTo errhandler

    For Each lsiThisMirror In lvwMirrors.ListItems
        If lsiThisMirror.Selected Then
            mcolMirrors.Remove lsiThisMirror.Key
        End If
    Next lsiThisMirror
    
    RefreshMirrorList
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmddelMirror_Click", Err.Description, Erl
End Sub

Private Sub lvwMirrors_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errhandler

    mcolMirrors.Item(Item.Key).Enabled = Item.Checked
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "lvwMirrors_ItemCheck", Err.Description, Erl
    
End Sub

Private Sub optBlockOption_Click(Index As Integer)
    On Error GoTo errhandler

    txtBlockMessage.Enabled = (Index = 1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "optBlockOption_Click", Err.Description, Erl
    
End Sub

Private Sub optClientIP_Click(Index As Integer)
    On Error GoTo errhandler

    lblClientTo.Visible = (Index = 1)
    txtClientIPTo.Visible = (Index = 1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "optClientIP_Click", Err.Description, Erl
    
End Sub

Private Sub optMatchTypeIP_Click(Index As Integer)
    On Error GoTo errhandler

    txtRange2.Visible = (Index = 1)
    lblRange.Visible = (Index = 1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "optMatchTypeIP_Click", Err.Description, Erl
    
End Sub

Private Sub optMatchTypePort_Click(Index As Integer)
    On Error GoTo errhandler

    txtRangePort2.Visible = (Index = 1)
    lblPortTo.Visible = (Index = 1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "optMatchTypePort_Click", Err.Description, Erl
    
End Sub

Private Sub optMatchTypeUsers_Click(Index As Integer)
    On Error GoTo errhandler

    txtRangeUsers2.Visible = (Index = 1)
    lblUsersTo.Visible = (Index = 1)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "optMatchTypeUsers_Click", Err.Description, Erl
    
End Sub

Private Sub scrInterval_Change()
    On Error GoTo errhandler

    lblMirrorInterval.Caption = CStr(scrInterval.Value)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "scrInterval_Change", Err.Description, Erl
    
End Sub

Private Sub tbsFilterTypes_Click()
    On Error GoTo errhandler

    ShowFrame tbsFilterTypes.SelectedItem.Index, fraFilterTypes
    'DisplayCommonFilterElements tbsFilterTypes.SelectedItem.Index
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "tbsFilterTypes_Click", Err.Description, Erl
    
End Sub

Private Sub txtAddress_Change()
    On Error GoTo errhandler

    cmdAddMirror.Enabled = (Len(txtAddress.Text) > 0)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtAddress_Change", Err.Description, Erl
    
End Sub

' ################################
' General Tab

Private Sub chkReqPass_Click()
    On Error GoTo errhandler

    txtPass.BackColor = IIf(chkReqPass.Value = vbChecked, vbWhite, DISABLED_COL)
    txtPass.Enabled = (chkReqPass.Value = vbChecked)
    If Not mbLoading Then
        If txtPass.Enabled Then txtPass.SetFocus
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "chkReqPass_Click", Err.Description, Erl

End Sub

Private Sub cmdColour1_Click()
    Dim lColour As OLE_COLOR
    
    On Error GoTo errhandler
    
    CommonDialog1.ShowColor
    Text1.ForeColor = CommonDialog1.Color
    
    Exit Sub
errhandler:
    If Err.Number <> cdlCancel Then
        ErrorReport Err.Number, MOD_NAME, "cmdColour1_Click", Err.Description, Erl
    End If
    
End Sub

Private Sub cmdColour2_Click()
    Dim lColour As OLE_COLOR
    
    On Error GoTo errhandler
    
    CommonDialog1.ShowColor
    Text1.BackColor = CommonDialog1.Color
    
    Exit Sub
errhandler:
    If Err.Number <> cdlCancel Then
        ErrorReport Err.Number, MOD_NAME, "cmdColour2_Click", Err.Description, Erl
    End If
End Sub

Private Sub txtExpire_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If KeyAscii = 8 Or KeyAscii = 9 Then
    Else
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtExpire_KeyPress", Err.Description, Erl
    
End Sub

' ################################
' Fake Servers Tab

Private Sub cmdUpdate_Click()
    Dim objServer As clsServer
    On Error GoTo errhandler
    
    Set objServer = BuildServer
    If Not objServer Is Nothing Then
        mobjServerGrid.UpdateServer objServer
    End If
    
    Exit Sub
errhandler:
    If Err.Number <= 0 Then
        MsgBox Err.Description
    Else
        ErrorReport Err.Number, MOD_NAME, "cmdUpdate_Click", Err.Description, Erl
    End If
    
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errhandler
    mobjServerGrid.DeleteServer mobjServerGrid.SelectedServer
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdDelete_Click", Err.Description, Erl
End Sub

Private Sub cmdNew_Click()
    On Error GoTo errhandler
    NewServer
    txtName.SetFocus
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdNew_Click", Err.Description, Erl
End Sub

Private Function BuildServer() As clsServer
    Dim objServer As clsServer
    Dim arrIP() As String
    Dim sErrLocation As String
    Dim ctlFields As Control
    
    On Error GoTo errhandler
    
    Set objServer = BuildNewServerObject
    
    sErrLocation = "Server Name"
    Set ctlFields = txtName
    objServer.ServerName = txtName.Text
    
    sErrLocation = "Server Description"
    Set ctlFields = txtDescription
    objServer.Description = txtDescription.Text
    
    sErrLocation = "Port Number"
    Set ctlFields = txtPort
    If CLng(txtPort.Text) = 0 Then
        Err.Raise vbObjectError
    End If
    
    objServer.Port = CLng(txtPort.Text)
    
    sErrLocation = "User Count"
    Set ctlFields = txtUsers
    objServer.UserCount = CLng(txtUsers.Text)
    
    sErrLocation = "IP Address"
    Set ctlFields = txtIP1
    'If CLng(txtIP1.Text) = 0 Then
    '    Err.Raise vbObjectError
    'End If
    objServer.SetIP txtIP1.Text, txtIP2.Text, txtIP3.Text, txtIP4.Text
    
    objServer.ServerType = Fake
    objServer.LastCheckedIn = Now
    Set BuildServer = objServer
    Set objServer = Nothing
    
    Exit Function
errhandler:
    If Len(sErrLocation) > 0 Then
        MsgBox "Invalid " & sErrLocation
        ctlFields.SetFocus
    Else
        ErrorReport Err.Number, MOD_NAME, "BuildServer", Err.Description, Erl
    End If
    Set BuildServer = Nothing
    
End Function

Private Sub NewServer()
    Dim objServer As clsServer
    Dim arrIP() As String
    
    On Error GoTo errhandler
    Set objServer = BuildNewServerObject
    
    objServer.SetIP "0", "0", "0", "0"
    objServer.ServerName = "New Server"
    objServer.ServerType = Fake = True
    objServer.LastCheckedIn = Now
    
    mobjServerGrid.AddNewServer objServer
    
    Set objServer = Nothing
    
    Exit Sub
errhandler:
    If Err.Number = ERR_BASE Then
        MsgBox Err.Description
    Else
        ErrorReport Err.Number, MOD_NAME, "NewServer", Err.Description, Erl
    End If
End Sub

Private Sub mobjServerGrid_SelectionChanged(objServer As clsServer)
    Dim arrIP() As String
    On Error GoTo errhandler
    
    If Not objServer Is Nothing Then
        EnableFields True
        txtName.Text = objServer.ServerName
        txtDescription.Text = objServer.Description
        txtPort = objServer.Port
        txtUsers = objServer.UserCount
        arrIP = Split(objServer.IPString, ".")
        txtIP1.Text = arrIP(0)
        txtIP2.Text = arrIP(1)
        txtIP3.Text = arrIP(2)
        txtIP4.Text = arrIP(3)
    Else
        EnableFields False
        txtName.Text = ""
        txtDescription.Text = ""
        txtPort = ""
        txtUsers = ""
        txtIP1.Text = ""
        txtIP2.Text = ""
        txtIP3.Text = ""
        txtIP4.Text = ""
    End If
    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "mobjServerGrid_SelectionChanged", Err.Description, Erl

End Sub

Private Sub EnableFields(bEnable As Boolean)
    On Error GoTo errhandler

    txtName.Enabled = bEnable
    txtDescription.Enabled = bEnable
    txtIP1.Enabled = bEnable
    txtIP2.Enabled = bEnable
    txtIP3.Enabled = bEnable
    txtIP4.Enabled = bEnable
    txtPort.Enabled = bEnable
    txtUsers.Enabled = bEnable
    
    txtName.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtDescription.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP1.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP2.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP3.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP4.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtPort.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtUsers.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "EnableFields", Err.Description, Erl
    
End Sub

Private Sub txtDescription_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtDescription_Change", Err.Description, Erl
    
End Sub

Private Sub txtDescription_GotFocus()
    On Error GoTo errhandler

    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtDescription_GotFocus", Err.Description, Erl
    
End Sub

Private Sub txtIP1_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP1_Change", Err.Description, Erl
    
End Sub
Private Sub txtIP2_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP2_Change", Err.Description, Erl
    
End Sub
Private Sub txtIP3_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP3_Change", Err.Description, Erl
    
End Sub
Private Sub txtIP4_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP4_Change", Err.Description, Erl
    
End Sub

Private Sub txtIP1_GotFocus()
    On Error GoTo errhandler

    txtIP1.SelStart = 0
    txtIP1.SelLength = Len(txtIP1.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP1_GotFocus", Err.Description, Erl
    
End Sub
Private Sub txtIP2_GotFocus()
    On Error GoTo errhandler

    txtIP2.SelStart = 0
    txtIP2.SelLength = Len(txtIP2.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP2_GotFocus", Err.Description, Erl
    
End Sub
Private Sub txtIP3_GotFocus()
    On Error GoTo errhandler

    txtIP3.SelStart = 0
    txtIP3.SelLength = Len(txtIP3.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP3_GotFocus", Err.Description, Erl
    
End Sub
Private Sub txtIP4_GotFocus()
    On Error GoTo errhandler

    txtIP4.SelStart = 0
    txtIP4.SelLength = Len(txtIP4.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtIP4_GotFocus", Err.Description, Erl
    
End Sub

Private Sub txtName_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtName_Change", Err.Description, Erl
    
End Sub

Private Sub txtName_GotFocus()
    On Error GoTo errhandler

    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtName_GotFocus", Err.Description, Erl
    
End Sub

Private Sub txtPort_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtPort_Change", Err.Description, Erl
    
End Sub

Private Sub txtPort_GotFocus()
    On Error GoTo errhandler

    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtPort_GotFocus", Err.Description, Erl
    
End Sub


Private Sub txtRange1_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = Asc(".")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRange1_KeyPress", Err.Description, Erl
    
End Sub

Private Sub txtRange2_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = Asc(".")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRange2_KeyPress", Err.Description, Erl
    
End Sub

Private Sub txtRangePort1_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRangePort1_KeyPress", Err.Description, Erl
    
End Sub

Private Sub txtRangePort2_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRangePort2_KeyPress", Err.Description, Erl
    
End Sub


Private Sub txtRangeUsers1_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRangeUsers1_KeyPress", Err.Description, Erl
    
End Sub

Private Sub txtRangeUsers2_KeyPress(KeyAscii As Integer)
    On Error GoTo errhandler

    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
    Else
        KeyAscii = 0
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtRangeUsers2_KeyPress", Err.Description, Erl
    
End Sub

Private Sub txtUsers_Change()
    On Error GoTo errhandler

    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtUsers_Change", Err.Description, Erl
    
End Sub

Private Sub txtUsers_GotFocus()
    On Error GoTo errhandler

    txtUsers.SelStart = 0
    txtUsers.SelLength = Len(txtUsers.Text)
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "txtUsers_GotFocus", Err.Description, Erl
    
End Sub

Private Sub RefreshButtons()
    Dim objServer As clsServer
    Set objServer = mobjServerGrid.SelectedServer
    
    If Not objServer Is Nothing Then
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = (txtName.Text <> objServer.ServerName) Or _
                            (txtDescription.Text <> objServer.Description) Or _
                            (txtPort.Text <> CStr(objServer.Port)) Or _
                            (txtIP1.Text & "." & txtIP2.Text & "." & txtIP3.Text & "." & txtIP4.Text <> objServer.IPString) Or _
                            (txtUsers.Text <> CStr(objServer.UserCount))
    Else
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
    End If
    
End Sub

Private Sub LoadBlockedIPs()
    Dim objIPRange As clsIPRange
    Dim lviThisItem As ListItem
    
    lvwClients.ListItems.Clear
    For Each objIPRange In mobjTempBlockedIPs
        Set lviThisItem = lvwClients.ListItems.Add(, objIPRange.UniqueID, objIPRange.Describe)
        lviThisItem.Checked = objIPRange.Enabled
    Next objIPRange
    
End Sub

Private Sub LoadFakeServers()
    Dim objServer As clsServer
    Dim asData() As String
    Dim vntServer As Variant
    
  ' In case the reg entry contains duplicates or is corrupt...
    On Error Resume Next
    
    asData = Split(mobjSettings.FakeServers, Chr(0))
    For Each vntServer In asData
        Set objServer = BuildFakeServer(CStr(vntServer))
        If Not objServer Is Nothing Then
            mobjServerGrid.AddNewServer objServer
        End If
    Next vntServer
    
    'For Each objServer In mobjServers
    '    If objServer.ServerType = Fake Then mobjServerGrid.AddNewServer objServer
    
    'Next objServer
    
End Sub

' ################################

Private Sub cmdOK_Click()
    Dim sErr As String
    On Error GoTo errhandler
    
    sErr = CheckValues
    If Len(sErr) = 0 Then
        mobjSettings.FakeServers = mobjServerGrid.GetRegString
        mobjServerGrid.MergeServers mobjServers
    
1       mobjSettings.TextColour = Text1.ForeColor
2       mobjSettings.BackColour = Text1.BackColor
3       frmMain.txtData.ForeColor = mobjSettings.TextColour
4       frmMain.txtData.BackColor = mobjSettings.BackColour
5       mobjSettings.AlertExpire = (chkAlertExpire.Value = vbChecked)
6       mobjSettings.AlertNew = (chkAlertNewServer.Value = vbChecked)
7       mobjSettings.AlertList = (chkAlertRequest.Value = vbChecked)
8       mobjSettings.Float = (chkFloat.Value = vbChecked)
9       mobjSettings.Password = txtPass.Text
10      mobjSettings.RequirePassword = (chkReqPass.Value = vbChecked)
11      mobjSettings.ExpireInterval = CLng(txtExpire.Text)
12      SetWindowPos frmMain.hwnd, IIf(mobjSettings.Float, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
13      mobjSettings.DoHitLogging = (chkHitLogging.Value = vbChecked)
14      mobjSettings.DoServerLogging = (chkServerLogging.Value = vbChecked)
15      mobjSettings.DoDisplayLogging = (chkDisplayLogging = vbChecked)

16      Set mobjSettings.TrackerMirrors = CloneCollection(mcolMirrors)
17      Set mobjSettings.Filters = CloneCollection2(mobjFilters)
        Set mobjBlockedIPs = CloneCollection4(mobjTempBlockedIPs)
        mobjSettings.BlockedIPs = mobjBlockedIPs.BuildRegString
        
        mobjSettings.RefuseBlocked = optBlockOption(0).Value
        mobjSettings.BlockedMsg = txtBlockMessage.Text
        
18      mobjSettings.ShowLogTimestamp = (chkLogTimeStamp.Value = vbChecked)
19      mobjSettings.ShowLogServerAdd = (chkLogAddServer.Value = vbChecked)
20      mobjSettings.ShowLogServerExpire = (chkLogServerExpire.Value = vbChecked)
21      mobjSettings.ShowLogMirrors = (chkLogMirror.Value = vbChecked)
22      mobjSettings.ShowLogListReq = (chkLogListing.Value = vbChecked)
23      mobjSettings.ShowLogServerFiltered = (chkLogFilter.Value = vbChecked)
24      mobjSettings.ShowLogBadPass = (chkLogBadPassword.Value = vbChecked)

        mobjSettings.DoMirroring = (chkDoMirrors.Value = vbChecked)

25      mobjSettings.RestrictLogSize = (chkTrimLogs.Value = vbChecked)
        If Not IsNumeric(txtLogMax.Text) Then
            txtLogMax.Text = 0
        End If
26      mobjSettings.MaxLogSize = CLng(txtLogMax.Text)
27      mobjSettings.MirrorInterval = scrInterval.Value
        
        mobjSettings.FilterDefault = optSpecial(1).Value
' IDEA - warn if mirror interval > expire interval
        If mbFiltersChanged Then
            mobjServers.NotifyFiltersChanged
        End If
28      frmMain.ReCalculateCounts
29      Unload Me
    Else
        MsgBox sErr, vbExclamation, APP_NAME
    End If


    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdOK_Click", Err.Description & Erl
    Unload Me
End Sub

Private Function CheckValues() As String
    Dim lExpire As Long
    
    If Not IsNumeric(txtExpire.Text) Then
        CheckValues = "Invalid expiry interval"
        txtExpire.SetFocus
    End If
    lExpire = CLng(txtExpire.Text)
    If lExpire < MIN_EXPIRE Or lExpire > MAX_EXPIRE Then
        CheckValues = "Invalid expiry interval - enter a value between " & MIN_EXPIRE & " and " & MAX_EXPIRE
        txtExpire.SetFocus
    End If

End Function

Private Sub cmdCancel_Click()
    On Error GoTo errhandler
    'Debug.Print lvwFilters.ListItems(1).Width
    Unload Me
    'Debug.Print Me.lvwFilters.ColumnHeaders(1).Width
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdCancel_Click", Err.Description, Erl
End Sub
' todo - hide errorreports for release build
Private Sub Form_Load()
    On Error GoTo errhandler
    
    mbLoading = True
    PositionFrames
    ShowVersion
    ShowFrame 1, fraSettings
    ShowFrame 1, fraFilterTypes
    Width = tbsSettings.Width + FRM_HDIFF
    Height = tbsSettings.Height + FRM_VDIFF
    'lvwMirrors.ColumnHeaders(1).Width = lvwMirrors.Width - 100
    SetWindowPos Me.hwnd, IIf(mobjSettings.Float, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    CommonDialog1.CancelError = True
    Text1.ForeColor = mobjSettings.TextColour
    Text1.BackColor = mobjSettings.BackColour
    chkAlertExpire.Value = IIf(mobjSettings.AlertExpire, vbChecked, vbUnchecked)
    chkAlertNewServer.Value = IIf(mobjSettings.AlertNew, vbChecked, vbUnchecked)
    chkAlertRequest.Value = IIf(mobjSettings.AlertList, vbChecked, vbUnchecked)
    chkFloat.Value = IIf(mobjSettings.Float, vbChecked, vbUnchecked)
    txtPass.Text = mobjSettings.Password
    txtPass.BackColor = IIf(mobjSettings.RequirePassword, vbWhite, DISABLED_COL)
    txtPass.Enabled = mobjSettings.RequirePassword
    chkReqPass.Value = IIf(mobjSettings.RequirePassword, vbChecked, vbUnchecked)
    txtExpire.Text = mobjSettings.ExpireInterval
    chkHitLogging.Value = IIf(mobjSettings.DoHitLogging, vbChecked, vbUnchecked)
    chkServerLogging.Value = IIf(mobjSettings.DoServerLogging, vbChecked, vbUnchecked)
    chkDisplayLogging.Value = IIf(mobjSettings.DoDisplayLogging, vbChecked, vbUnchecked)
    chkDoMirrors.Value = IIf(mobjSettings.DoMirroring, vbChecked, vbUnchecked)
    cmdMirrorNow.Enabled = mobjSettings.DoMirroring
    
    lnkViewServerLog.LinkTarget = mobjLog.ServerLogPath
    lnkViewHitLog.LinkTarget = mobjLog.HitLogPath
    lnkViewWindowLog.LinkTarget = mobjLog.WindowLogPath
    
    lnkViewServerLog.Enabled = (Len(Dir(mobjLog.ServerLogPath)) > 0)
    lnkViewHitLog.Enabled = (Len(Dir(mobjLog.HitLogPath)) > 0)
    lnkViewWindowLog.Enabled = (Len(Dir(mobjLog.WindowLogPath)) > 0)
    
    chkLogTimeStamp.Value = IIf(mobjSettings.ShowLogTimestamp, vbChecked, vbUnchecked)
    chkLogAddServer.Value = IIf(mobjSettings.ShowLogServerAdd, vbChecked, vbUnchecked)
    chkLogServerExpire.Value = IIf(mobjSettings.ShowLogServerExpire, vbChecked, vbUnchecked)
    chkLogMirror.Value = IIf(mobjSettings.ShowLogMirrors, vbChecked, vbUnchecked)
    chkLogListing.Value = IIf(mobjSettings.ShowLogListReq, vbChecked, vbUnchecked)
    chkLogFilter.Value = IIf(mobjSettings.ShowLogServerFiltered, vbChecked, vbUnchecked)
    chkLogBadPassword.Value = IIf(mobjSettings.ShowLogBadPass, vbChecked, vbUnchecked)
    
    chkTrimLogs.Value = IIf(mobjSettings.RestrictLogSize, vbChecked, vbUnchecked)
    txtLogMax.Enabled = (chkTrimLogs.Value = vbChecked)
    txtLogMax.Text = CStr(mobjSettings.MaxLogSize)
    
    Set mcolMirrors = CloneCollection(mobjSettings.TrackerMirrors)
    RefreshMirrorList
    
    Set mobjFilters = CloneCollection2(mobjSettings.Filters)
    RefreshFilterList
    
    Set mobjTempBlockedIPs = CloneCollection4(mobjBlockedIPs)
    LoadBlockedIPs
    
    optBlockOption(0).Value = (mobjSettings.RefuseBlocked)
    optBlockOption(1).Value = Not (mobjSettings.RefuseBlocked)
    txtBlockMessage.Text = mobjSettings.BlockedMsg
    
    Set mobjServerGrid = New clsServerGrid
    Set mobjServerGrid.ListViewControl = lvwServers
    LoadFakeServers
    'lvwFilters.ListItems(1).Width = 4065
    lblMirrorInterval.Caption = CStr(mobjSettings.MirrorInterval)
    scrInterval = mobjSettings.MirrorInterval
    
    optSpecial(1).Value = (mobjSettings.FilterDefault)
    
    mobjServerGrid.RequestSelectionEvent
    
    UpdateUpTime
    mbLoading = False
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "Form_Load", Err.Description, Erl
End Sub

Private Sub ShowVersion()
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub PositionFrames()
    Dim fraThis As Frame
    
    For Each fraThis In fraSettings
        fraThis.Top = tbsSettings.Top + FRA_VOFFSET
        fraThis.Left = tbsSettings.Left + FRA_HOFFSET
    Next fraThis
    
    For Each fraThis In fraFilterTypes
        Set fraThis.Container = fraSettings(7)
        fraThis.Top = tbsFilterTypes.Top + FRA_VOFFSET
        fraThis.Left = tbsFilterTypes.Left + FRA_HOFFSET
    Next fraThis
    
    
End Sub

Private Sub tmrExpire_Timer()
    On Error GoTo errhandler

    UpdateUpTime
    Exit Sub
errhandler:
    tmrExpire.Interval = 0
    ErrorReport Err.Number, MOD_NAME, "Picture1_Click", Err.Description, Erl
End Sub
Private Sub UpdateUpTime()
    lblUpTime.Caption = GetUpTime
End Sub

Private Sub tbsSettings_Click()
    On Error GoTo errhandler

    Select Case tbsSettings.SelectedItem
        Case "Appearance"
            ShowFrame 1, fraSettings
        Case "Connections"
            ShowFrame 2, fraSettings
        Case "Alerts"
            ShowFrame 3, fraSettings
        Case "Logs"
            ShowFrame 4, fraSettings
        Case "Fake Servers"
            ShowFrame 5, fraSettings
        Case "Mirrors"
            ShowFrame 6, fraSettings
        Case "Filters"
            ShowFrame 7, fraSettings
        Case "About"
            ShowFrame 8, fraSettings
        Case Else
        '
    End Select
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "tbsSettings_Click", Err.Description, Erl
    
End Sub

Private Sub ShowFrame(nIndex As Integer, fraSet As Variant)
    Dim nThisTab As Integer
    Dim fraThis As Frame
    
    For Each fraThis In fraSet
        fraThis.Visible = (fraThis.Index = nIndex)
    Next fraThis
    
End Sub


