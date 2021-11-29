VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4596
   ClientLeft      =   8376
   ClientTop       =   2340
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4596
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOptDep 
      BorderStyle     =   0  'None
      Height          =   444
      Left            =   4468
      ScaleHeight     =   444
      ScaleWidth      =   2460
      TabIndex        =   37
      Top             =   4080
      Visible         =   0   'False
      Width           =   2460
      Begin VB.OptionButton optDepByForm 
         Caption         =   "By Form"
         Height          =   372
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   1140
      End
      Begin VB.OptionButton optDepByDep 
         Caption         =   "By Dep."
         Height          =   372
         Left            =   1224
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.PictureBox picOptFonts 
      BorderStyle     =   0  'None
      Height          =   444
      Left            =   4468
      ScaleHeight     =   444
      ScaleWidth      =   2460
      TabIndex        =   34
      Top             =   5184
      Visible         =   0   'False
      Width           =   2460
      Begin VB.OptionButton optFontsByForm 
         Caption         =   "By Form"
         Height          =   372
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   1140
      End
      Begin VB.OptionButton optFontsByFont 
         Caption         =   "By Font"
         Height          =   372
         Left            =   1224
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdCollapseTree 
      Caption         =   "Collapse all"
      Height          =   372
      Left            =   2064
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.PictureBox picScanning 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   876
      Left            =   2716
      ScaleHeight     =   876
      ScaleWidth      =   4380
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   4380
      Begin VB.Label lblScanning3 
         Alignment       =   1  'Right Justify
         Height          =   276
         Left            =   96
         TabIndex        =   31
         Top             =   600
         Width           =   4188
      End
      Begin VB.Label lblScanning2 
         Alignment       =   1  'Right Justify
         Height          =   276
         Left            =   96
         TabIndex        =   30
         Top             =   360
         Width           =   4188
      End
      Begin VB.Label lblScanning 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   96
         TabIndex        =   28
         Top             =   120
         Width           =   4188
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to clipboard"
      Enabled         =   0   'False
      Height          =   372
      Left            =   216
      TabIndex        =   6
      Top             =   4080
      Width           =   1572
   End
   Begin TabDlg.SSTab sst1 
      Height          =   3852
      Left            =   144
      TabIndex        =   0
      Top             =   120
      Width           =   6708
      _ExtentX        =   11832
      _ExtentY        =   6795
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Scan"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdScan"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSummary"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNote"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Dependencies"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "trvDepByForm"
      Tab(1).Control(1)=   "trvDepByDep"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Strings"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "trvStrings"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Fonts"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "trvFontsByForm"
      Tab(3).Control(1)=   "trvFontsByFont"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Find controls"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picTabContainer(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Replace fonts"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdReplaceFonts"
      Tab(5).Control(1)=   "tmrRefrehcbo"
      Tab(5).Control(2)=   "chkFontOfObject"
      Tab(5).Control(3)=   "txtFontProperties"
      Tab(5).Control(4)=   "cmdSelectFontProperties"
      Tab(5).Control(5)=   "cboNewFontName"
      Tab(5).Control(6)=   "cboNewFontSize"
      Tab(5).Control(7)=   "cboOrigFontName"
      Tab(5).Control(8)=   "cboOrigFontSize"
      Tab(5).Control(9)=   "cmdSelectObjects"
      Tab(5).Control(10)=   "txtObjects"
      Tab(5).Control(11)=   "cmsSelectControlTypes"
      Tab(5).Control(12)=   "txtControlTypes"
      Tab(5).Control(13)=   "Label10"
      Tab(5).Control(14)=   "Label9"
      Tab(5).Control(15)=   "Label8"
      Tab(5).Control(16)=   "Label7"
      Tab(5).Control(17)=   "Label6"
      Tab(5).Control(18)=   "Label5"
      Tab(5).Control(19)=   "Label4"
      Tab(5).Control(20)=   "Label3"
      Tab(5).ControlCount=   21
      TabCaption(6)   =   "Copy controls"
      TabPicture(6)   =   "frmMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picTabContainer(6)"
      Tab(6).ControlCount=   1
      Begin VB.CommandButton cmdNote 
         Caption         =   "?"
         Height          =   324
         Left            =   2432
         TabIndex        =   65
         Top             =   1440
         Width           =   564
      End
      Begin VB.PictureBox picTabContainer 
         BorderStyle     =   0  'None
         Height          =   3600
         Index           =   6
         Left            =   -74952
         ScaleHeight     =   3600
         ScaleWidth      =   6612
         TabIndex        =   51
         Top             =   408
         Width           =   6612
         Begin VB.CommandButton cmdAddControls 
            Caption         =   "Add controls"
            Height          =   372
            Left            =   4456
            TabIndex        =   62
            Top             =   3120
            Width           =   1572
         End
         Begin VB.ComboBox cboOrigControl 
            Height          =   336
            Left            =   1848
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1632
            Width           =   3600
         End
         Begin VB.ComboBox cboOrigObject 
            Height          =   336
            Left            =   1848
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1224
            Width           =   3600
         End
         Begin VB.TextBox txtDestinationObjects 
            Height          =   360
            Left            =   1848
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "All the others"
            Top             =   2040
            Width           =   3600
         End
         Begin VB.CommandButton cmdSelectDestinationObjects 
            Caption         =   "Select"
            Height          =   348
            Left            =   5580
            TabIndex        =   54
            Top             =   2040
            Width           =   1140
         End
         Begin VB.Label lblPicturepropertiesNote 
            Caption         =   "Note: properties having pictures will not be copied, if there are any, you'll have to copy them manually."
            Height          =   468
            Left            =   120
            TabIndex        =   61
            Top             =   2540
            Width           =   6996
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Origin form/usc):"
            Height          =   372
            Left            =   76
            TabIndex        =   58
            Top             =   1272
            Width           =   1640
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Objects (forms/usc):"
            Height          =   372
            Left            =   48
            TabIndex        =   57
            Top             =   2088
            Width           =   1668
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Control to copy:"
            Height          =   372
            Left            =   76
            TabIndex        =   56
            Top             =   1680
            Width           =   1640
         End
         Begin VB.Label Label13 
            Caption         =   "The control to be copied to other forms / usercontrols must already exist on some form / usercontrol. Please select it:"
            Height          =   468
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   6996
         End
         Begin VB.Label Label12 
            Caption         =   $"frmMain.frx":00C4
            ForeColor       =   &H000000C0&
            Height          =   708
            Left            =   120
            TabIndex        =   52
            Top             =   0
            Width           =   6948
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picTabContainer 
         BorderStyle     =   0  'None
         Height          =   3348
         Index           =   4
         Left            =   -74904
         ScaleHeight     =   3348
         ScaleWidth      =   6564
         TabIndex        =   40
         Top             =   360
         Width           =   6558
         Begin VB.ComboBox cboPropertyToCompare2 
            Height          =   336
            Left            =   3716
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   960
            Visible         =   0   'False
            Width           =   2828
         End
         Begin VB.ComboBox cboCriteria 
            Height          =   336
            ItemData        =   "frmMain.frx":015B
            Left            =   5584
            List            =   "frmMain.frx":0168
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   96
            Width           =   1480
         End
         Begin VB.CheckBox chkIgnoreCase 
            Caption         =   """A""=""a"""
            Height          =   336
            Left            =   96
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Ignore Case"
            Top             =   960
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   828
         End
         Begin ComctlLib.TreeView trvFind 
            Height          =   1380
            Left            =   96
            TabIndex        =   49
            Top             =   1440
            Visible         =   0   'False
            Width           =   5988
            _ExtentX        =   10562
            _ExtentY        =   2434
            _Version        =   327682
            Style           =   7
            Appearance      =   1
         End
         Begin VB.CommandButton cmdFindControls 
            Caption         =   "Go"
            Height          =   348
            Left            =   6664
            TabIndex        =   48
            Top             =   936
            Width           =   396
         End
         Begin VB.TextBox txtPropertyValue 
            Height          =   360
            Left            =   3716
            TabIndex        =   47
            Top             =   960
            Width           =   2828
         End
         Begin VB.ComboBox cboPropertyValueCondition 
            Height          =   336
            ItemData        =   "frmMain.frx":0196
            Left            =   1980
            List            =   "frmMain.frx":0198
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   960
            Width           =   1628
         End
         Begin VB.ComboBox cboPropertyToCompare 
            Height          =   336
            Left            =   1980
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   528
            Width           =   2828
         End
         Begin VB.ComboBox cboControlType 
            Height          =   336
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   96
            Width           =   2828
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Criteria:"
            Height          =   348
            Left            =   4900
            TabIndex        =   66
            Top             =   108
            Width           =   612
         End
         Begin VB.Label lblCondition 
            Alignment       =   1  'Right Justify
            Caption         =   "Condition:"
            Height          =   348
            Left            =   144
            TabIndex        =   46
            Top             =   972
            Width           =   1716
         End
         Begin VB.Label lblPropertyToCompare 
            Alignment       =   1  'Right Justify
            Caption         =   "Property to compare:"
            Height          =   348
            Left            =   144
            TabIndex        =   44
            Top             =   540
            Width           =   1716
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Control type:"
            Height          =   348
            Left            =   132
            TabIndex        =   42
            Top             =   108
            Width           =   1716
         End
      End
      Begin VB.CommandButton cmdReplaceFonts 
         Caption         =   "Replace fonts"
         Enabled         =   0   'False
         Height          =   348
         Left            =   -70500
         TabIndex        =   22
         Top             =   3536
         Width           =   1524
      End
      Begin VB.Timer tmrRefrehcbo 
         Interval        =   1
         Left            =   -74664
         Top             =   3192
      End
      Begin VB.CheckBox chkFontOfObject 
         Caption         =   "Replace object Font too"
         Height          =   324
         Left            =   -70016
         TabIndex        =   26
         ToolTipText     =   "Replaces the Font of the Form or UserControl itself"
         Top             =   1872
         Value           =   1  'Checked
         Width           =   2212
      End
      Begin VB.TextBox txtFontProperties 
         Height          =   360
         Left            =   -73152
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "All"
         Top             =   1440
         Width           =   1600
      End
      Begin VB.CommandButton cmdSelectFontProperties 
         Caption         =   "Select"
         Height          =   348
         Left            =   -71424
         TabIndex        =   23
         Top             =   1440
         Width           =   1140
      End
      Begin VB.ComboBox cboNewFontName 
         Height          =   336
         Left            =   -73152
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2760
         Width           =   2028
      End
      Begin VB.ComboBox cboNewFontSize 
         Enabled         =   0   'False
         Height          =   336
         ItemData        =   "frmMain.frx":019A
         Left            =   -70416
         List            =   "frmMain.frx":019C
         TabIndex        =   18
         Top             =   2760
         Width           =   732
      End
      Begin VB.ComboBox cboOrigFontName 
         Height          =   336
         Left            =   -73152
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2304
         Width           =   2028
      End
      Begin VB.ComboBox cboOrigFontSize 
         Height          =   336
         ItemData        =   "frmMain.frx":019E
         Left            =   -70416
         List            =   "frmMain.frx":01A0
         TabIndex        =   14
         Top             =   2304
         Width           =   732
      End
      Begin VB.CommandButton cmdSelectObjects 
         Caption         =   "Select"
         Height          =   348
         Left            =   -71424
         TabIndex        =   13
         Top             =   1848
         Width           =   1140
      End
      Begin VB.TextBox txtObjects 
         Height          =   360
         Left            =   -73152
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "All"
         Top             =   1848
         Width           =   1600
      End
      Begin VB.CommandButton cmsSelectControlTypes 
         Caption         =   "Select"
         Height          =   348
         Left            =   -71424
         TabIndex        =   10
         Top             =   1032
         Width           =   1140
      End
      Begin VB.TextBox txtControlTypes 
         Height          =   360
         Left            =   -73152
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "All"
         Top             =   1032
         Width           =   1600
      End
      Begin VB.TextBox txtSummary 
         Height          =   2412
         Left            =   168
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1770
         Width           =   5892
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan project"
         Height          =   372
         Left            =   192
         TabIndex        =   1
         Top             =   520
         Width           =   1572
      End
      Begin ComctlLib.TreeView trvDepByForm 
         Height          =   2292
         Left            =   -74904
         TabIndex        =   2
         Top             =   1068
         Visible         =   0   'False
         Width           =   5988
         _ExtentX        =   10562
         _ExtentY        =   4043
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.TreeView trvStrings 
         Height          =   2292
         Left            =   -74880
         TabIndex        =   4
         Top             =   792
         Width           =   5988
         _ExtentX        =   10562
         _ExtentY        =   4043
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.TreeView trvFontsByForm 
         Height          =   2292
         Left            =   -74880
         TabIndex        =   5
         Top             =   1008
         Visible         =   0   'False
         Width           =   5988
         _ExtentX        =   10562
         _ExtentY        =   4043
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.TreeView trvDepByDep 
         Height          =   2292
         Left            =   -74928
         TabIndex        =   32
         Top             =   816
         Width           =   5988
         _ExtentX        =   10562
         _ExtentY        =   4043
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.TreeView trvFontsByFont 
         Height          =   2292
         Left            =   -74832
         TabIndex        =   33
         Top             =   648
         Width           =   5988
         _ExtentX        =   10562
         _ExtentY        =   4043
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Caution:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   264
         TabIndex        =   63
         Top             =   984
         Width           =   696
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Properties:"
         Height          =   372
         Left            =   -74424
         TabIndex        =   25
         Top             =   1488
         Width           =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "New Font.Name:"
         Height          =   348
         Left            =   -74976
         TabIndex        =   21
         Top             =   2772
         Width           =   1716
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   348
         Left            =   -70992
         TabIndex        =   20
         Top             =   2784
         Width           =   468
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Original Font.Name:"
         Height          =   348
         Left            =   -74976
         TabIndex        =   17
         Top             =   2316
         Width           =   1716
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   348
         Left            =   -70992
         TabIndex        =   16
         Top             =   2328
         Width           =   468
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Objects (forms/usc):"
         Height          =   372
         Left            =   -74952
         TabIndex        =   11
         Top             =   1896
         Width           =   1668
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Control type:"
         Height          =   372
         Left            =   -74424
         TabIndex        =   8
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   $"frmMain.frx":01A2
         ForeColor       =   &H000000C0&
         Height          =   708
         Left            =   -74832
         TabIndex        =   7
         Top             =   408
         Width           =   6950
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Caption         =   $"frmMain.frx":0239
         Height          =   732
         Left            =   288
         TabIndex        =   64
         Top             =   984
         Width           =   6868
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function HashData Lib "shlwapi" (ByVal pbData As Long, ByVal cbData As Long, ByRef pbHash As Any, ByVal cbHash As Long) As Long

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private mDPIf As Single
Private mProject  As VBProject
Private mObjects_Name() As String
Private mObjects_ObjectOwnFont() As cPropFont
Private mObjects_ObjectOwnCaption() As String
Private mObjects_ControlNames() As Variant
Private mObjects_ControlTypes() As Variant
Private mObjects_ControlPropertiesFont() As Variant
Private mObjects_ControlPropertiesString() As Variant
Private mDependencies As Collection
Private mControlTypesGlobal As Collection
Private mControlTypesGlobal_PropTypes As Collection
Private mFontsGlobal As Collection
Private mProjectsNames As Collection

Private mSelectedControlTypes As Collection
Private mSelectedFontProperties As Collection
Private mSelectedObjects As Collection
Private mPropertyNamesWithFont As Collection

Private mIndentLevel As Long
Private mTreeText As String
Private mScanning As Boolean
Private mCanceled As Boolean
Private mUnloading As Boolean
Private mReplacingFonts As Boolean
Private mFindCriteria As String
Private mFinding As Boolean
Private mCopyingControls As Boolean
Private mLastProjectName As String

Private mDesignerWindowsVisibility As Collection
Private mDesignerWindowsVisibility_ProjectName As String
Private mDesignerWindowsZOrder As Collection

Private WithEvents mProjects As VBIDE.VBProjectsEvents
Attribute mProjects.VB_VarHelpID = -1
Private mInIDE As Boolean

Private Sub cboControlType_Click()
    Dim iCol As Collection
    Dim iPt As cPropType
    
    cboPropertyToCompare.Clear
    If cboControlType.ListIndex > 0 Then
        Set iCol = mControlTypesGlobal_PropTypes(cboControlType.ListIndex)
        For Each iPt In iCol
            cboPropertyToCompare.AddItem iPt.PropertyName
            cboPropertyToCompare.ItemData(cboPropertyToCompare.NewIndex) = iPt.ReturnType
            cboPropertyToCompare2.AddItem iPt.PropertyName
            cboPropertyToCompare2.ItemData(cboPropertyToCompare2.NewIndex) = iPt.ReturnType
        Next
        cboPropertyToCompare.AddItem "[Please select]", 0
        cboPropertyToCompare.ListIndex = 0
        cboPropertyToCompare2.AddItem "[Please select]", 0
        cboPropertyToCompare2.ListIndex = 0
    End If
End Sub

Private Sub cboCriteria_Click()
    Select Case cboCriteria.ListIndex
        Case 0 ' list all
            txtPropertyValue.Visible = True
            cboPropertyToCompare2.Visible = False
            lblPropertyToCompare.Enabled = False
            cboPropertyToCompare.Enabled = False
            chkIgnoreCase.Enabled = False
            lblCondition.Enabled = False
            cboPropertyValueCondition.Enabled = False
            txtPropertyValue.Enabled = False
        Case 1 ' compare property value
            txtPropertyValue.Visible = True
            cboPropertyToCompare2.Visible = False
            cboPropertyToCompare.Enabled = True
            chkIgnoreCase.Enabled = True
            lblCondition.Enabled = True
            cboPropertyValueCondition.Enabled = True
            txtPropertyValue.Enabled = True
        Case 2 ' comprare values of two properties
            txtPropertyValue.Visible = False
            cboPropertyToCompare2.Visible = True
            cboPropertyToCompare.Enabled = True
            chkIgnoreCase.Enabled = True
            lblCondition.Enabled = True
            cboPropertyValueCondition.Enabled = True
            txtPropertyValue.Enabled = True
    End Select
End Sub

Private Sub cboNewFontName_Click()
    EnableDisableFontReplacementButton
End Sub

Private Sub cboNewFontSize_Change()
    EnableDisableFontReplacementButton
End Sub

Private Sub cboNewFontSize_Click()
    cboNewFontSize_Change
End Sub

Private Sub cboOrigFontName_Click()
    EnableDisableFontReplacementButton
End Sub

Private Sub cboOrigFontSize_Change()
    If Val(cboOrigFontSize.Text) = 0 Then
        cboNewFontSize.Text = ""
        cboNewFontSize.Enabled = False
    Else
        If cboNewFontSize.Text = "" Then
            If Val(cboOrigFontSize.Text) >= 8 Then
                cboNewFontSize.Text = cboOrigFontSize.Text
            End If
        End If
        cboNewFontSize.Enabled = True
    End If
    EnableDisableFontReplacementButton
End Sub


Private Sub cboOrigFontSize_Click()
    cboOrigFontSize_Change
End Sub

Private Sub cboOrigObject_Click()
    Dim iComp As VBComponent
    Dim iDes As Object
    Dim iCtl As VBControl
    Dim iDesignerWindowVisible As Boolean
    Dim iIsDirty As Boolean
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
        If mProject.Name <> cboOrigObject.Tag Then
            MsgBox "Project changed. Please select again", vbExclamation
            FillcboOrigObject
            Exit Sub
        End If
        cboOrigControl.Clear
        If cboOrigObject.ListIndex > 0 Then
            Screen.MousePointer = vbHourglass
            For Each iComp In mProject.VBComponents
                If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
                    If iComp.Name = cboOrigObject.Text Then
                        iIsDirty = iComp.IsDirty
                        iDesignerWindowVisible = iComp.DesignerWindow.Visible
                        Set iDes = iComp.Designer
                        For Each iCtl In iDes.VBControls
                            cboOrigControl.AddItem iCtl.ControlObject.Name & " (" & iCtl.ProgId & ")"
                        Next
                        Set iCtl = Nothing
                        If iComp.IsDirty And (Not iIsDirty) Then
                            On Error Resume Next
                            iComp.Reload
                            On Error GoTo 0
                        End If
                        If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
                        Set iDes = Nothing
                    End If
                End If
                Set iComp = Nothing
            Next
            cboOrigControl.AddItem "[Please select]", 0
            cboOrigControl.ListIndex = 0
            Screen.MousePointer = vbDefault
        End If
    Else
        cboOrigControl.Clear
        MsgBox "The is no active project.", vbExclamation
    End If
End Sub

Private Sub cboPropertyToCompare_Click()
    Dim iPropType As EPropertyReturnType
    
    iPropType = cboPropertyToCompare.ItemData(cboPropertyToCompare.ListIndex)
    
    If cboPropertyToCompare.ListIndex > 0 Then
        If Val(cboPropertyValueCondition.Tag) <> iPropType Then
            cboPropertyValueCondition.Clear
            cboPropertyValueCondition.AddItem "[Please select]"
            If iPropType = eptNumeric Then
                chkIgnoreCase.Visible = False
                cboPropertyValueCondition.AddItem "="
                cboPropertyValueCondition.AddItem "<>"
                cboPropertyValueCondition.AddItem ">"
                cboPropertyValueCondition.AddItem ">="
                cboPropertyValueCondition.AddItem "<"
                cboPropertyValueCondition.AddItem "<="
            Else
                chkIgnoreCase.Visible = True
                cboPropertyValueCondition.AddItem "="
                cboPropertyValueCondition.AddItem "<>"
                cboPropertyValueCondition.AddItem "contains"
                cboPropertyValueCondition.AddItem "is contained in"
            End If
            cboPropertyValueCondition.Tag = iPropType
            cboPropertyValueCondition.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmdAddControls_Click()
    If mCopyingControls Then
        If MsgBox("Cancel? (controls already added won't be undone).", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then Exit Sub
        mCanceled = True
        cmdAddControls.Caption = "Add controls"
        cmdAddControls.Move 4456, 3120
    Else
        If cboOrigObject.ListIndex <= 0 Then
            MsgBox "Please select the object (form or unsercontrol) from where to copy.", vbExclamation
            cboOrigObject.SetFocus
            Exit Sub
        End If
        If cboOrigControl.ListIndex <= 0 Then
            MsgBox "Please select the control to copy.", vbExclamation
            cboOrigControl.SetFocus
            Exit Sub
        End If
        
        If Not VBInstance.ActiveVBProject Is Nothing Then
            Set mProject = VBInstance.ActiveVBProject
            If mProject.Name <> cboOrigObject.Tag Then
                MsgBox "Project changed. Please select again", vbExclamation
                FillcboOrigObject
                Exit Sub
            End If
            
            cmdAddControls.Move 360, 3220
            mCanceled = False
            cmdAddControls.Caption = "Cancel"
            Screen.MousePointer = vbArrowHourglass
            
            DoAddControls
            Screen.MousePointer = vbDefault
            If Not mUnloading Then
                EnableOtherTabs True
                picScanning.Visible = False
                cmdAddControls.Move 4456, 3120
                lblPicturepropertiesNote.Visible = True
                cmdAddControls.Caption = "Add controls"
            End If
            
        Else
            cboOrigControl.Clear
            MsgBox "The is no active project.", vbExclamation
        End If
    End If
End Sub

Private Sub cmdCollapseTree_Click()
    Dim trv As TreeView
    Dim i As Long
    
    Select Case sst1.Tab
        Case 1
            If optDepByForm.Value Then
                Set trv = trvDepByForm
            Else
                Set trv = trvDepByDep
            End If
        Case 2
            Set trv = trvStrings
        Case 3
            If optFontsByForm.Value Then
                Set trv = trvFontsByForm
            Else
                Set trv = trvFontsByFont
            End If
    End Select
    If Not trv Is Nothing Then
        For i = 1 To trv.Nodes.Count
            If trv.Nodes(i).Parent Is Nothing Then
                trv.Nodes(i).Expanded = False
            End If
        Next
        If trv.Nodes.Count > 0 Then trv.Nodes(1).EnsureVisible
    End If
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    mTreeText = Caption
    mIndentLevel = 0
    Select Case sst1.Tab
        Case 0
            mTreeText = mTreeText & vbCrLf & vbCrLf
            mTreeText = mTreeText & txtSummary.Text
        Case 1
            If optDepByForm.Value Then
                mTreeText = mTreeText & vbCrLf & vbCrLf
                If trvDepByForm.Nodes.Count > 0 Then
                    ParseTree trvDepByForm.Nodes(1)
                Else
                    mTreeText = mTreeText & "None."
                End If
            Else
                mTreeText = mTreeText & vbCrLf & vbCrLf
                If trvDepByDep.Nodes.Count > 0 Then
                    ParseTree trvDepByDep.Nodes(1)
                Else
                    mTreeText = mTreeText & "None."
                End If
            End If
        Case 2
            mTreeText = mTreeText & ", Strings in properties:" & vbCrLf & vbCrLf
            If trvStrings.Nodes.Count > 0 Then
                ParseTree trvStrings.Nodes(1)
            Else
                mTreeText = mTreeText & "None."
            End If
        Case 3
            mTreeText = mTreeText & ", Fonts" & ":" & vbCrLf & vbCrLf
            If optFontsByForm.Value Then
                If trvFontsByForm.Nodes.Count > 0 Then
                    ParseTree trvFontsByForm.Nodes(1)
                Else
                    mTreeText = mTreeText & "None."
                End If
            Else
                If trvFontsByFont.Nodes.Count > 0 Then
                    ParseTree trvFontsByFont.Nodes(1)
                Else
                    mTreeText = mTreeText & "None."
                End If
            End If
        Case 4
            mTreeText = mTreeText & ", " & mFindCriteria & ":" & vbCrLf & vbCrLf
            If trvFind.Nodes.Count > 0 Then
                ParseTree trvFind.Nodes(1)
            Else
                mTreeText = mTreeText & "None."
            End If
    End Select
    Clipboard.SetText mTreeText
    MsgBox "Copied", vbInformation
End Sub

Private Sub cmdFindControls_Click()
    If mFinding Then
        mCanceled = True
        cmdFindControls.FontName = "Segoe UI"
        cmdFindControls.Caption = "Go"
        mFinding = False
        picScanning.Visible = False
    Else
        If cboControlType.ListIndex <= 0 Then
            MsgBox "Please select the Control Type.", vbExclamation
            cboControlType.SetFocus
            Exit Sub
        End If
        If cboCriteria.ListIndex > 0 Then
            If cboPropertyToCompare.ListIndex <= 0 Then
                MsgBox "Please select a Property.", vbExclamation
                cboPropertyToCompare.SetFocus
                Exit Sub
            End If
            If cboPropertyValueCondition.ListIndex <= 0 Then
                MsgBox "Please select a Condition.", vbExclamation
                cboPropertyValueCondition.SetFocus
                Exit Sub
            End If
        End If
        If cboCriteria.ListIndex = 1 Then
            txtPropertyValue.Text = Trim(txtPropertyValue.Text)
       '     If txtPropertyValue.Text = "" Then
        '        MsgBox "Please enter a criteria.", vbExclamation
         '       txtPropertyValue.SetFocus
          '      Exit Sub
          '  End If
            If Val(cboPropertyValueCondition.Tag) = eptNumeric Then
                If Not IsNumeric(txtPropertyValue.Text) Then
                    MsgBox "The value must be numeric.", vbExclamation
                    txtPropertyValue.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf cboCriteria.ListIndex = 2 Then
            If cboPropertyToCompare2.ListIndex <= 0 Then
                MsgBox "Please select the second Property.", vbExclamation
                cboPropertyToCompare2.SetFocus
                Exit Sub
            End If
        End If
        
        cmdFindControls.FontName = "Wingdings 2"
        cmdFindControls.Caption = "X"
        
        FindControls
        
        mFinding = False
        If Not mUnloading Then
            trvFind.Height = 1300
            EnableOtherTabs True
            cmdFindControls.FontName = "Segoe UI"
            cmdFindControls.Caption = "Go"
            picScanning.Visible = False
            trvFind.Move 100 - picTabContainer(4).Left, sst1.TabHeight + 1160, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500 - 1400
            If Not mCanceled Then
                trvFind.Visible = True
                cmdCopy.Visible = True
            End If
        End If
    End If
End Sub

Private Sub cmdNote_Click()
    Const Message As String = "Option 1 (easier): From the Visual Basic menu, choose Tools and then Options, in the General Tab ensure that the Error Trapping section is set to ""Break on Unhandled Errors"". The IDE default is ""Break in Class Module""." & vbCrLf & vbCrLf & _
    "Option 2: Run the utility and if it hangs, it will tell at the bottom what Form has the UserControl, and most important what property was reading when it hanged." & vbCrLf & _
    "Modify the code in the UserControl so as not to generate errors when Ambient.UserMode is False." & vbCrLf & vbCrLf & _
    "When the IDE hangs, you'll have to close it from the Windows Task Manager. So ensure you have to project with any changes saved."
    
    MsgBox Message
End Sub

Private Sub cmdReplaceFonts_Click()
    If MsgBox("Replace fonts?", vbYesNo Or vbDefaultButton2 Or vbExclamation) = vbNo Then Exit Sub
    If mSelectedControlTypes Is Nothing Then SelectAllControlTypes
    If mPropertyNamesWithFont Is Nothing Then SetPropertyNamesWithFont
    If mSelectedFontProperties Is Nothing Then SelectAllFontproperties
    If mSelectedObjects Is Nothing Then SelectAllObjects

    If mReplacingFonts Then
        If MsgBox("Cancel Replace? (fonts already replaced won't be undone).", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then Exit Sub
        mCanceled = True
        cmdReplaceFonts.Caption = "Replace fonts"
        cmdReplaceFonts.Left = 4500
        Me.Refresh
    Else
        If Not VBInstance.ActiveVBProject Is Nothing Then
            cmdReplaceFonts.Left = 360
            mCanceled = False
            cmdReplaceFonts.Caption = "Cancel replace"
            Screen.MousePointer = vbArrowHourglass
            DoReplaceFonts
            Screen.MousePointer = vbDefault
            If Not mUnloading Then
                EnableOtherTabs True
                picScanning.Visible = False
                cmdReplaceFonts.Left = 4500
                cmdReplaceFonts.Caption = "Replace fonts"
            End If
        Else
            MsgBox "No current project."
        End If
    End If
End Sub

Private Sub cmdScan_Click()
    Dim c As Long
    
    If mScanning Then
        mCanceled = True
        cmdScan.Caption = "Scan project"
        txtSummary.Move 100, 1870, sst1.Width / mDPIf - 200, 1600
        txtSummary.Refresh
        Me.Refresh
    Else
        On Error Resume Next
        Set mProject = VBInstance.ActiveVBProject
        If Err.Number Then
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        If Not mLastProjectName = "" Then
            If mProject.Name <> mLastProjectName Then
                Set mSelectedControlTypes = Nothing
                Set mSelectedFontProperties = Nothing
                Set mSelectedObjects = Nothing
                Set mPropertyNamesWithFont = Nothing
            End If
        End If
        
        If Not mProject Is Nothing Then
            txtSummary.Move 100, 1870, sst1.Width / mDPIf - 200, 1600
            mCanceled = False
            cmdScan.Caption = "Cancel scan"
            Screen.MousePointer = vbArrowHourglass
            Me.Caption = "Project: " & mProject.Name
            Scan
            Screen.MousePointer = vbDefault
            On Error Resume Next
            mLastProjectName = mProject.Name
            If Err.Number Then
                mLastProjectName = ""
            End If
            On Error GoTo 0
            If Not mUnloading Then
                If Not mCanceled Then
                    EnableOtherTabs True
                    cmdCopy.Enabled = True
                    ShowResults
                    picScanning.Visible = False
                    txtSummary.Move 100, 1870, sst1.Width / mDPIf - 200, sst1.Height / mDPIf - 1870 - 150 - 500
                    FillcboOrigObject
                End If
                cmdScan.Caption = "Scan project"
            End If
        Else
            txtSummary.Text = "No current project."
        End If
    End If
End Sub

Private Sub Scan()
    Dim iComp As VBComponent
    Dim iDes As Object
    Dim iCtl As VBControl
    Dim iControlNames As Collection
    Dim iControlTypes As Collection
    Dim iAuxControlPropertiesFont As Collection
    Dim iAuxControlPropertiesString As Collection
    Dim i As Long
    Dim iStrs() As String
    Dim n As Long
    Dim iProp As Property
    Dim iObj As Object
    Dim iAuxPropsFont As Collection
    Dim iAuxPropsString As Collection
    Dim iFont As cPropFont
    Dim iString As cPropString
    Dim c As Long
    Dim iVar As Variant
    Dim iStr As String
    Dim iNullCount As Long
    Dim c2 As Long
    Dim iPropsWithStringsCount As Long
    Dim iPropsWithFontsCount As Long
    Dim p As Long
    Dim p2 As Long
    Dim pc As Long
    Dim pr As Long
    Dim iFormCount As Long
    Dim iUserControlCount As Long
    Dim iFormNumber As Long
    Dim iUserControlNumber As Long
    Dim iPrj As VBProject
    Dim iControlIsInSourceCode As Boolean
    Dim iVarType As Long
    Dim iDesignerWindowVisible As Boolean
    Dim iIsDirty As Boolean
    
    ReDim mObjects_Name(0)
    ReDim mObjects_ObjectOwnFont(0)
    ReDim mObjects_ObjectOwnCaption(0)
    ReDim mObjects_ControlNames(0)
    ReDim mObjects_ControlTypes(0)
    ReDim mObjects_ControlPropertiesFont(0)
    ReDim mObjects_ControlPropertiesString(0)
    Set mDependencies = New Collection
    Set mProjectsNames = New Collection
    Set mControlTypesGlobal = New Collection
    Set mControlTypesGlobal_PropTypes = New Collection
    Set mFontsGlobal = New Collection
    mScanning = True
    
    For Each iPrj In VBInstance.VBProjects
        mProjectsNames.Add iPrj.Name, iPrj.Name
    Next
    
    EnableOtherTabs False
    cmdCopy.Enabled = False
    On Error GoTo ErrorExit
    
    lblScanning.Caption = ""
    lblScanning2.Caption = ""
    lblScanning3.Caption = ""
    picScanning.Visible = True
    For Each iComp In mProject.VBComponents
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            If iComp.Type = vbext_ct_UserControl Then
                iUserControlCount = iUserControlCount + 1
            Else
                iFormCount = iFormCount + 1
            End If
            pc = pc + 1
        End If
        Set iComp = Nothing
    Next
    
    For Each iComp In mProject.VBComponents
        'Debug.Print iComp.Name
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            If iComp.Type = vbext_ct_UserControl Then
                iUserControlNumber = iUserControlNumber + 1
            Else
                iFormNumber = iFormNumber + 1
            End If
            p = p + 1
            txtSummary.Text = "Forms: " & iFormNumber & " of " & iFormCount & vbCrLf
            txtSummary.Text = txtSummary.Text & "UserControls: " & iUserControlNumber & " of " & iUserControlCount & vbCrLf
            picScanning.Line (0, 0)-(picScanning.ScaleWidth / pc * p, 60), vbGreen, BF
            picScanning.Line (picScanning.ScaleWidth / pc * p, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
            'Debug.Print vbTab & "Scanning: " & iComp.Name & "..."
            lblScanning.Caption = "Scanning: " & iComp.Name & "..."
            lblScanning.Refresh
            If mInIDE Then Debug.Print lblScanning.Caption
            DoEvents
            If mCanceled Then
                mScanning = False
                picScanning.Visible = False
                Exit Sub
            End If
            pr = 0
            iIsDirty = iComp.IsDirty
            iDesignerWindowVisible = iComp.DesignerWindow.Visible
            Set iDes = iComp.Designer
            Set iControlNames = New Collection
            Set iControlTypes = New Collection
            Set iAuxControlPropertiesFont = New Collection
            Set iAuxControlPropertiesString = New Collection
            
            For Each iCtl In iDes.VBControls
                'Debug.Print iCtl.ControlObject.Name
                lblScanning2.Caption = "Control: " & iCtl.ControlObject.Name & " type: " & TypeName(iCtl.ControlObject)
                lblScanning2.Refresh
                If mInIDE Then Debug.Print vbTab & lblScanning2.Caption
                lblScanning3.Caption = ""
                lblScanning3.Refresh
                If iCtl.ControlObject Is Nothing Then
                    mCanceled = True
                    mScanning = False
                    picScanning.Visible = False
                    Exit Sub
                End If
                Set iAuxPropsFont = New Collection
                Set iAuxPropsString = New Collection
                iControlNames.Add GetControlName(iCtl.ControlObject)
                iControlTypes.Add iCtl.ProgId
                iStrs = Split(iCtl.ProgId, ".")
                AddDependency iStrs(0)
                iControlIsInSourceCode = DependencyIsFromSourceProjects(iStrs(0))
                If Not ControlTypesGlobalExists(iCtl.ProgId) Then
                    AddControlTypeGlobal iCtl
                End If
                For Each iProp In iCtl.Properties
'                    Debug.Print iProp.Name
 '                   Stop
                    lblScanning3.Caption = "Property: " & iProp.Name
                    If mInIDE Then Debug.Print vbTab & vbTab & lblScanning3.Caption
                    lblScanning3.Refresh
                    'Debug.Print vbTab & iProp.Name
                    pr = pr + 1
                    If pr > 1000 Then
                        pr = 0
                        DoEvents
                        If mCanceled Then
                            mScanning = False
                            picScanning.Visible = False
                            Exit Sub
                        End If
                    End If
                    iVar = Empty
                    On Error Resume Next
                    iVar = iProp
                    On Error GoTo ErrorExit
                    iVarType = VarType(iVar)
                    If iVarType = vbObject Then
                        Set iObj = Nothing
                        On Error Resume Next
                        Set iObj = iProp.object
                        On Error GoTo ErrorExit
                        If Not iObj Is Nothing Then
                            If TypeName(iProp.object) = "Font" Then
                                Set iFont = New cPropFont
                                iFont.PropertyName = iProp.Name
                                iFont.FontName = iObj.Name
                                iFont.FontSize = iObj.Size
                                AddFontGlobal iFont
                                iAuxPropsFont.Add iFont
                                iPropsWithFontsCount = iPropsWithFontsCount + 1
                            End If
                        ElseIf Not iControlIsInSourceCode Then
                            If iProp.NumIndices = 1 Then
                                iVar = Empty
                                On Error Resume Next
                                iVar = iProp.IndexedValue(0)
                                On Error GoTo ErrorExit
                                If VarType(iVar) = vbString Then
                                    iStr = ""
                                    c = 0
                                    iNullCount = 0
                                    Do Until iNullCount > 100
                                        If iVar = "" Then
                                            iNullCount = iNullCount + 1
                                        Else
                                            If iNullCount > 0 Then
                                                For c2 = 1 To iNullCount
                                                    If iStr <> "" Then iStr = iStr & "|"
                                                Next
                                            End If
                                            iNullCount = 0
                                            If iStr <> "" Then iStr = iStr & "|"
                                            iStr = iStr & CStr(iVar)
                                        End If
                                        c = c + 1
                                        iVar = Empty
                                        On Error Resume Next
                                        iVar = iProp.IndexedValue(c)
                                        On Error GoTo ErrorExit
                                    Loop
                                    Set iString = New cPropString
                                    iString.PropertyName = iProp.Name & "()"
                                    iString.StringValue = iStr
                                    iAuxPropsString.Add iString
                                    iPropsWithStringsCount = iPropsWithStringsCount + 1
                                End If
                            End If
                        End If
                        Set iObj = Nothing
                    ElseIf iVarType = vbString Then
                        If (iProp.Name <> "Name") And (iProp.Name <> "FontName") Then
                            iStr = ""
                            On Error Resume Next
                            iStr = iProp.Value
                            On Error GoTo ErrorExit
                            If iStr <> "" Then
                                Set iString = New cPropString
                                iString.PropertyName = iProp.Name
                                iString.StringValue = iStr
                                iAuxPropsString.Add iString
                                iPropsWithStringsCount = iPropsWithStringsCount + 1
                            End If
                        End If
                    End If
                    iVar = Empty
                Next
                Set iProp = Nothing
                iAuxControlPropertiesFont.Add iAuxPropsFont
                iAuxControlPropertiesString.Add iAuxPropsString
            Next
            Set iCtl = Nothing
            i = UBound(mObjects_Name) + 1
            ReDim Preserve mObjects_Name(i)
            ReDim Preserve mObjects_ObjectOwnFont(i)
            ReDim Preserve mObjects_ObjectOwnCaption(i)
            ReDim Preserve mObjects_ControlNames(i)
            ReDim Preserve mObjects_ControlTypes(i)
            ReDim Preserve mObjects_ControlPropertiesFont(i)
            ReDim Preserve mObjects_ControlPropertiesString(i)
            mObjects_Name(i) = iComp.Name
            Set mObjects_ControlNames(i) = iControlNames
            Set mObjects_ControlTypes(i) = iControlTypes
            Set mObjects_ControlPropertiesFont(i) = iAuxControlPropertiesFont
            Set mObjects_ControlPropertiesString(i) = iAuxControlPropertiesString
            
            Set iFont = New cPropFont
            Set iProp = Nothing
            For p2 = 1 To iComp.Properties.Count
                Set iProp = iComp.Properties(p2)
                If iProp.Name = "Font" Then
                    Set iObj = Nothing
                    On Error Resume Next
                    Set iObj = iProp.object
                    On Error GoTo ErrorExit
                    If Not iObj Is Nothing Then
                        If TypeName(iProp.object) = "Font" Then
                            Set iFont = New cPropFont
                            iFont.PropertyName = iProp.Name
                            iFont.FontName = iObj.Name
                            iFont.FontSize = iObj.Size
                            AddFontGlobal iFont
                            iFont.PropertyName = "Font"
                        End If
                    End If
                    Set iObj = Nothing
                ElseIf iProp.Name = "Caption" Then
                    iStr = ""
                    On Error Resume Next
                    iStr = iProp.Value
                    On Error GoTo ErrorExit
                    mObjects_ObjectOwnCaption(i) = iStr
                End If
            Next
            Set iProp = Nothing 'Bug in the Add-In environment, if not set to Nothing VB chashes with UserControls when the Add-In is compiled
            Set mObjects_ObjectOwnFont(i) = iFont
            
            If iComp.IsDirty And (Not iIsDirty) Then
                On Error Resume Next
                iComp.Reload
                On Error GoTo ErrorExit
            End If
            If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
            Set iDes = Nothing
        End If
        Set iComp = Nothing
    Next
    txtSummary.Text = ""
    txtSummary.SelText = "Summary:" & vbCrLf & vbCrLf
    n = UBound(mObjects_Name)
    txtSummary.SelText = n & IIf(n = 1, " Object (" & iFormCount & " Form" & " / " & iUserControlCount & " UserControl" & ")", " Objects (" & iFormCount & IIf(iFormCount = 1, " Form", " Forms") & " / " & iUserControlCount & IIf(iUserControlCount = 1, " UserControl", " UserControls") & ")") & "." & vbCrLf
    n = mControlTypesGlobal.Count
    txtSummary.SelText = n & IIf(n = 1, " Control type", " Controls types")
    n = mDependencies.Count
    txtSummary.SelText = " in " & n & IIf(n = 1, " Component", " Components") & "." & vbCrLf
    txtSummary.SelText = "Number of properties with fonts: " & iPropsWithFontsCount & vbCrLf
    txtSummary.SelText = "Number of properties with strings: " & iPropsWithStringsCount & vbCrLf
    txtSummary.SelText = vbCrLf
    txtSummary.SelText = "Please go to other tab for details."
    lblScanning.Caption = ""
    lblScanning2.Caption = ""
    lblScanning3.Caption = ""
    Set mPropertyNamesWithFont = Nothing
    mScanning = False
    Exit Sub
    
ErrorExit:
    mScanning = False
    mCanceled = True
    picScanning.Visible = False
    txtSummary.Text = "Canceled"
End Sub

Private Sub AddDependency(nStr As String)
    On Error Resume Next
    mDependencies.Add nStr, nStr
End Sub

Private Function DependencyIsFromSourceProjects(nStr As String) As Boolean
    Dim iStr As String
    
    On Error GoTo TheExit
    iStr = mProjectsNames(nStr)
    DependencyIsFromSourceProjects = True
    
TheExit:
End Function

Private Function ControlTypesGlobalExists(nStr As String) As Boolean
    Dim iStr As String
    
    On Error GoTo TheExit
    iStr = mControlTypesGlobal(nStr)
    ControlTypesGlobalExists = True
TheExit:
End Function

Private Sub AddControlTypeGlobal(nCtl As VBControl)
    Dim iCol As Collection
    Dim iProp As Property
    Dim iVar As Variant
    Dim iPt As cPropType
    
    mControlTypesGlobal.Add nCtl.ProgId, nCtl.ProgId
    
    Set iCol = New Collection
    For Each iProp In nCtl.Properties
        If mInIDE Then Debug.Print vbTab & vbTab & "Property: " & iProp.Name
        iVar = Empty
        On Error Resume Next
        iVar = iProp
        On Error GoTo 0
        Select Case VarType(iVar)
            Case vbString
                Set iPt = New cPropType
                iPt.ReturnType = eptString
                iPt.PropertyName = iProp.Name
                iCol.Add iPt
            Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbDecimal, vbDate
                Set iPt = New cPropType
                iPt.ReturnType = eptNumeric
                iPt.PropertyName = iProp.Name
                iCol.Add iPt
        End Select
    Next
    mControlTypesGlobal_PropTypes.Add iCol
End Sub

Private Sub AddFontGlobal(nFont As cPropFont)
    Dim iStr As String
    Dim iFnt As cPropFont
    
    iStr = nFont.FontName & "|" & nFont.FontSize
    If Not FontGlobalExists(iStr) Then
        Set iFnt = New cPropFont
        iFnt.FontName = nFont.FontName
        iFnt.FontSize = nFont.FontSize
        mFontsGlobal.Add iFnt, iStr
    End If
End Sub

Private Function FontGlobalExists(nStr As String) As Boolean
    Dim iObj As Object
    
    On Error GoTo TheExit
    Set iObj = mFontsGlobal(nStr)
    FontGlobalExists = True
TheExit:
End Function

Private Sub cmdSelectDestinationObjects_Click()
    Dim ifrmSI As New frmSelectItems
    Dim i As Long
    
    If mSelectedObjects Is Nothing Then SelectAllObjects
    For i = 1 To UBound(mObjects_Name)
        If mObjects_Name(i) <> cboOrigObject.Text Then
            ifrmSI.AddItem mObjects_Name(i), mObjects_Name(i), ObjectSelected(mObjects_Name(i))
        End If
    Next
    ifrmSI.Show vbModal, Me
    
    If ifrmSI.OKPressed Then
        Set mSelectedObjects = New Collection
        For i = 1 To ifrmSI.ItemCount
            If ifrmSI.Selected(i) Then
                mSelectedObjects.Add ifrmSI.Item(i), ifrmSI.Item(i)
            End If
        Next
        mSelectedObjects.Add cboOrigObject.Text, cboOrigObject.Text
        ShowCurrentSelectionOnCopyControls
    End If
End Sub

Private Sub cmdSelectObjects_Click()
    Dim ifrmSI As New frmSelectItems
    Dim i As Long
    
    If mSelectedObjects Is Nothing Then SelectAllObjects
    For i = 1 To UBound(mObjects_Name)
        ifrmSI.AddItem mObjects_Name(i), mObjects_Name(i), ObjectSelected(mObjects_Name(i))
    Next
    ifrmSI.Show vbModal, Me
    
    If ifrmSI.OKPressed Then
        Set mSelectedObjects = New Collection
        For i = 1 To ifrmSI.ItemCount
            If ifrmSI.Selected(i) Then
                mSelectedObjects.Add ifrmSI.Item(i), ifrmSI.Item(i)
            End If
        Next
        ShowCurrentSelectionOnReplaceFont
    End If
End Sub

Private Sub cmdSelectFontProperties_Click()
    Dim ifrmSI As New frmSelectItems
    Dim i As Long
    Dim iVar As Variant
    
    If mPropertyNamesWithFont Is Nothing Then SetPropertyNamesWithFont
    If mSelectedFontProperties Is Nothing Then SelectAllFontproperties
    For Each iVar In mPropertyNamesWithFont
        ifrmSI.AddItem iVar, iVar, FontPropertySelected(iVar)
    Next
    ifrmSI.Show vbModal, Me
    
    If ifrmSI.OKPressed Then
        Set mSelectedFontProperties = New Collection
        For i = 1 To ifrmSI.ItemCount
            If ifrmSI.Selected(i) Then
                mSelectedFontProperties.Add ifrmSI.Item(i), ifrmSI.Item(i)
            End If
        Next
        ShowCurrentSelectionOnReplaceFont
    End If
    
    ' mFontProperties_ControlPropertiesFont
    
End Sub

Private Sub SetPropertyNamesWithFont()
    Dim i As Long
    Dim iAuxControlPropertiesFont As Collection
    Dim iAuxPropsFont As Collection
    Dim iFont As cPropFont
    
    Set mPropertyNamesWithFont = New Collection
    For i = 1 To UBound(mObjects_ControlPropertiesFont)
        Set iAuxControlPropertiesFont = mObjects_ControlPropertiesFont(i)
        For Each iAuxPropsFont In iAuxControlPropertiesFont
            For Each iFont In iAuxPropsFont
                If Not PropertyNameWithFontExists(iFont.PropertyName) Then
                    mPropertyNamesWithFont.Add iFont.PropertyName, iFont.PropertyName
                End If
            Next
        Next
    Next
End Sub

Private Function PropertyNameWithFontExists(nPropertyname As String) As Boolean
    Dim iStr As String
    
    On Error GoTo TheExit
    iStr = mPropertyNamesWithFont(nPropertyname)
    PropertyNameWithFontExists = True
    
TheExit:
End Function

Private Sub SelectAllFontproperties()
    Dim iVar As Variant

    Set mSelectedFontProperties = New Collection
    For Each iVar In mPropertyNamesWithFont
        mSelectedFontProperties.Add iVar, iVar
    Next
End Sub

Private Sub cmsSelectControlTypes_Click()
    Dim ifrmSI As New frmSelectItems
    Dim iVar As Variant
    Dim i As Long
    
    If mSelectedControlTypes Is Nothing Then SelectAllControlTypes
    For Each iVar In mControlTypesGlobal
        ifrmSI.AddItem iVar, iVar, ControlTypeSelected(iVar)
    Next
    ifrmSI.Show vbModal, Me
    
    If ifrmSI.OKPressed Then
        Set mSelectedControlTypes = New Collection
        For i = 1 To ifrmSI.ItemCount
            If ifrmSI.Selected(i) Then
                mSelectedControlTypes.Add ifrmSI.Item(i), ifrmSI.Item(i)
            End If
        Next
        ShowCurrentSelectionOnReplaceFont
    End If
End Sub

Private Sub ShowCurrentSelectionOnReplaceFont()
    Dim iVar As Variant
    Dim i As Long
    
    txtControlTypes.Text = "All"
    cmsSelectControlTypes.Enabled = mControlTypesGlobal.Count > 1
    If Not mSelectedControlTypes Is Nothing Then
        If mSelectedControlTypes.Count = 0 Then
            txtControlTypes.Text = "None"
        ElseIf mSelectedControlTypes.Count = 1 Then
            txtControlTypes.Text = mSelectedControlTypes(1)
        Else
            For Each iVar In mControlTypesGlobal
                If Not ControlTypeSelected(iVar) Then
                    txtControlTypes.Text = "Selection"
                    Exit For
                End If
            Next
        End If
    ElseIf mControlTypesGlobal.Count = 1 Then
        txtControlTypes.Text = mControlTypesGlobal(1)
    End If
    
    txtFontProperties.Text = "All"
    If mPropertyNamesWithFont Is Nothing Then SetPropertyNamesWithFont
    cmdSelectFontProperties.Enabled = mPropertyNamesWithFont.Count > 1
    If Not mSelectedFontProperties Is Nothing Then
        If mSelectedFontProperties.Count = 0 Then
            txtControlTypes.Text = "None"
        ElseIf mSelectedFontProperties.Count = 1 Then
            txtFontProperties.Text = mSelectedFontProperties(1)
        Else
            For Each iVar In mPropertyNamesWithFont
                If Not FontPropertySelected(iVar) Then
                    txtFontProperties.Text = "Selection"
                    Exit For
                End If
            Next
        End If
    ElseIf mPropertyNamesWithFont.Count = 1 Then
        txtFontProperties.Text = mPropertyNamesWithFont(1)
    End If
    
    txtObjects.Text = "All"
    cmdSelectObjects.Enabled = UBound(mObjects_Name) > 1
    If Not mSelectedObjects Is Nothing Then
        If mSelectedObjects.Count = 0 Then
            txtObjects.Text = "None"
        ElseIf mSelectedObjects.Count = 1 Then
            txtObjects.Text = mSelectedObjects(1)
        Else
            For i = 1 To UBound(mObjects_Name)
                If Not ObjectSelected(mObjects_Name(i)) Then
                    txtObjects.Text = "Selection"
                    Exit For
                End If
            Next
        End If
    ElseIf UBound(mObjects_Name) = 1 Then
        txtObjects.Text = mObjects_Name(1)
    End If
End Sub

Private Sub ShowCurrentSelectionOnCopyControls()
    Dim i As Long
    
    txtDestinationObjects.Text = "All the others"
    cmdSelectDestinationObjects.Enabled = UBound(mObjects_Name) > 1
    If Not mSelectedObjects Is Nothing Then
        If mSelectedObjects.Count = 1 Then
            txtDestinationObjects.Text = "None"
        ElseIf mSelectedObjects.Count = 2 Then
            txtDestinationObjects.Text = mSelectedObjects(1)
        Else
            For i = 1 To UBound(mObjects_Name)
                If Not ObjectSelected(mObjects_Name(i)) Then
                    txtDestinationObjects.Text = "Selection"
                    Exit For
                End If
            Next
        End If
    Else
        txtDestinationObjects.Text = "All the others"
    End If
End Sub

Private Function ControlTypeSelected(nType As Variant) As Boolean
    Dim iStr As String
    
    If mSelectedControlTypes Is Nothing Then
        ControlTypeSelected = True
    Else
        On Error GoTo TheExit
        iStr = mSelectedControlTypes(nType)
        ControlTypeSelected = True
    End If
    
TheExit:
End Function

Private Function FontPropertySelected(nType As Variant) As Boolean
    Dim iStr As String
    
    If mSelectedFontProperties Is Nothing Then
        FontPropertySelected = True
    Else
        On Error GoTo TheExit
        iStr = mSelectedFontProperties(nType)
        FontPropertySelected = True
    End If
    
TheExit:
End Function

Private Sub SelectAllControlTypes()
    Dim iVar As Variant

    Set mSelectedControlTypes = New Collection
    For Each iVar In mControlTypesGlobal
        mSelectedControlTypes.Add iVar, iVar
    Next
End Sub

Private Function ObjectSelected(nName As Variant) As Boolean
    Dim iStr As String
    
    If mSelectedObjects Is Nothing Then
        ObjectSelected = True
    Else
        On Error GoTo TheExit
        iStr = mSelectedObjects(nName)
        ObjectSelected = True
    End If
    
TheExit:
End Function

Private Sub SelectAllObjects()
    Dim i As Long

    Set mSelectedObjects = New Collection
    For i = 1 To UBound(mObjects_Name)
        mSelectedObjects.Add mObjects_Name(i), mObjects_Name(i)
    Next
End Sub

Private Sub Form_Load()
    Dim c As Long
    
    mUnloading = False
    mInIDE = InIDE
    Set mProjects = Nothing
    Set mProjects = VBInstance.Events.VBProjectsEvents
    Me.Caption = App.Title
    mDPIf = GetTrueTwipsPerPixelX / Screen.TwipsPerPixelX
    sst1.Move 0, 60, Me.ScaleWidth * mDPIf, Me.ScaleHeight * mDPIf
    If sst1.Tab = 0 Then
        txtSummary.Move 100, 1870, sst1.Width / mDPIf - 200, sst1.Height / mDPIf - 1870 - 150 - 500
    End If
    sst1.Tab = 0
    For c = 1 To sst1.Tabs - 1
        sst1.TabEnabled(c) = False
    Next
    cmdScan.Caption = "Scan project"
    cmdCopy.Top = Me.ScaleHeight - 500
    LoadFontCombos
    cboCriteria.ListIndex = 1
    picOptFonts.Move picOptDep.Left, picOptDep.Top
End Sub

Private Sub ShowResults()
    Dim o As Long
    Dim iNode As Node
    Dim iType As Variant
    Dim iControlTypes As Collection
    Dim iCtlType As Variant
    Dim iFound As Boolean
    Dim iControlNames As Collection
    Dim i As Long
    Dim iDep As Variant
    Dim iStrs() As String
    Dim iAuxPropsString As Collection
    Dim iString As cPropString
    Dim iAuxControlPropertiesFont As Collection
    Dim iAuxControlPropertiesString As Collection
    Dim iAuxPropsFont As Collection
    Dim iFont As cPropFont
    Dim iOKey As String
    Dim iTKey As String
    Dim iDKey As String
    Dim iCKey As String
    Dim iFKey As String
    Dim iAuxCol As Collection
    Dim iVar As Variant
    
    ' Dependencies By Form
    trvDepByForm.Nodes.Clear
    
    picScanning.Line (0, 0)-(picScanning.ScaleWidth / 6 * 1, 60), vbGreen, BF
    picScanning.Line (picScanning.ScaleWidth / 6 * 1, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
    lblScanning2.Caption = "Loading: By Form..."
    lblScanning2.Refresh
    picScanning.Refresh
    
    For o = 1 To UBound(mObjects_Name)
        iOKey = SimpleHash("o_" & mObjects_Name(o))
        trvDepByForm.Nodes.Add , , iOKey, mObjects_Name(o)
        Set iControlTypes = mObjects_ControlTypes(o)
        Set iControlNames = mObjects_ControlNames(o)
        For Each iType In mControlTypesGlobal
            iFound = False
            For Each iCtlType In iControlTypes
                If iCtlType = iType Then
                    iFound = True
                    Exit For
                End If
            Next
            If iFound Then
                iTKey = SimpleHash("t_" & mObjects_Name(o) & "_" & iType)
                trvDepByForm.Nodes.Add iOKey, tvwChild, iTKey, iType
                For i = 1 To iControlTypes.Count
                    iCtlType = iControlTypes(i)
                    If iCtlType = iType Then
                        trvDepByForm.Nodes.Add iTKey, tvwChild, , iControlNames(i)
                    End If
                Next
            End If
        Next
    Next
    For i = 1 To trvDepByForm.Nodes.Count
        If trvDepByForm.Nodes(i).Parent Is Nothing Then
            trvDepByForm.Nodes(i).Expanded = True
        End If
    Next
    If trvDepByForm.Nodes.Count > 0 Then trvDepByForm.Nodes(1).EnsureVisible
    
    picScanning.Line (0, 0)-(picScanning.ScaleWidth / 6 * 2, 60), vbGreen, BF
    picScanning.Line (picScanning.ScaleWidth / 6 * 2, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
    lblScanning2.Caption = "Loading: By Dependency..."
    lblScanning2.Refresh
    picScanning.Refresh
    
    ' Dependencies By Dependency
    trvDepByDep.Nodes.Clear
    
    For Each iDep In mDependencies
        iDKey = SimpleHash("d_" & iDep)
        trvDepByDep.Nodes.Add , , iDKey, iDep
        For o = 1 To UBound(mObjects_Name)
            Set iControlTypes = mObjects_ControlTypes(o)
            Set iControlNames = mObjects_ControlNames(o)
            iFound = False
            For Each iCtlType In iControlTypes
                iStrs = Split(iCtlType, ".")
                If iStrs(0) = iDep Then
                    iFound = True
                    Exit For
                End If
            Next
            If iFound Then
                iOKey = SimpleHash("o_" & iDep & "_" & mObjects_Name(o))
                trvDepByDep.Nodes.Add iDKey, tvwChild, iOKey, mObjects_Name(o)
                For Each iType In mControlTypesGlobal
                    iStrs = Split(iType, ".")
                    If iStrs(0) = iDep Then
                        iFound = False
                        For i = 1 To iControlTypes.Count
                            iCtlType = iControlTypes(i)
                            If iCtlType = iType Then
                                iFound = True
                            End If
                        Next
                        If iFound Then
                            iTKey = SimpleHash("t_" & iDep & "_" & mObjects_Name(o) & "_" & iType)
                            trvDepByDep.Nodes.Add iOKey, tvwChild, iTKey, iType
                            For i = 1 To iControlTypes.Count
                                iCtlType = iControlTypes(i)
                                If iCtlType = iType Then
                                    trvDepByDep.Nodes.Add iTKey, tvwChild, , iControlNames(i)
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        Next
    Next
    For i = 1 To trvDepByDep.Nodes.Count
        If trvDepByDep.Nodes(i).Parent Is Nothing Then
            trvDepByDep.Nodes(i).Expanded = True
        End If
    Next
    If trvDepByDep.Nodes.Count > 0 Then trvDepByDep.Nodes(1).EnsureVisible
    
    picScanning.Line (0, 0)-(picScanning.ScaleWidth / 6 * 3, 60), vbGreen, BF
    picScanning.Line (picScanning.ScaleWidth / 6 * 3, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
    lblScanning2.Caption = "Loading: Strings..."
    lblScanning2.Refresh
    picScanning.Refresh
    
    ' Strings
    trvStrings.Nodes.Clear
    For o = 1 To UBound(mObjects_Name)
        Set iAuxControlPropertiesString = mObjects_ControlPropertiesString(o)
        For Each iAuxPropsString In iAuxControlPropertiesString
            If iAuxPropsString.Count > 0 Then
                iFound = True
                Exit For
            End If
        Next
        If iFound Or (mObjects_ObjectOwnCaption(o) <> "") Then
            iOKey = SimpleHash("o_" & mObjects_Name(o))
            trvStrings.Nodes.Add , , iOKey, mObjects_Name(o)
            If mObjects_ObjectOwnCaption(o) <> "" Then
                trvStrings.Nodes.Add iOKey, tvwChild, , mObjects_Name(o) & ".Caption: """ & mObjects_ObjectOwnCaption(o) & """"
            End If
            If iFound Then
                Set iControlTypes = mObjects_ControlTypes(o)
                Set iControlNames = mObjects_ControlNames(o)
                For Each iType In mControlTypesGlobal
                    iFound = False
                    For i = 1 To iControlTypes.Count
                        iCtlType = iControlTypes(i)
                        If iCtlType = iType Then
                            Set iAuxPropsString = iAuxControlPropertiesString(i)
                            If iAuxPropsString.Count > 0 Then
                                iFound = True
                                Exit For
                            End If
                        End If
                    Next
                    If iFound Then
                        iTKey = SimpleHash("t_" & mObjects_Name(o) & "_" & iType)
                        trvStrings.Nodes.Add iOKey, tvwChild, iTKey, iType
                        For i = 1 To iControlTypes.Count
                            iCtlType = iControlTypes(i)
                            If iCtlType = iType Then
                                iCKey = SimpleHash("c_" & mObjects_Name(o) & "_" & iType & "_" & iControlNames(i))
                                Set iAuxPropsString = iAuxControlPropertiesString(i)
                                If iAuxPropsString.Count > 0 Then
                                    trvStrings.Nodes.Add iTKey, tvwChild, iCKey, iControlNames(i)
                                    For Each iString In iAuxPropsString
                                        trvStrings.Nodes.Add iCKey, tvwChild, , iString.PropertyName & " Property" & ": """ & iString.StringValue & """"
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End If
    Next
    For i = 1 To trvStrings.Nodes.Count
        If trvStrings.Nodes(i).Parent Is Nothing Then
            trvStrings.Nodes(i).Expanded = True
        End If
    Next
    If trvStrings.Nodes.Count > 0 Then trvStrings.Nodes(1).EnsureVisible
    
    picScanning.Line (0, 0)-(picScanning.ScaleWidth / 6 * 4, 60), vbGreen, BF
    picScanning.Line (picScanning.ScaleWidth / 6 * 4, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
    lblScanning2.Caption = "Loading: Fonts..."
    lblScanning2.Refresh
    picScanning.Refresh
    
    ' Fonts By Form
    trvFontsByForm.Nodes.Clear
    For o = 1 To UBound(mObjects_Name)
        Set iAuxControlPropertiesFont = mObjects_ControlPropertiesFont(o)
        For Each iAuxPropsFont In iAuxControlPropertiesFont
            If iAuxPropsFont.Count > 0 Then
                iFound = True
                Exit For
            End If
        Next
        iOKey = SimpleHash("o_" & mObjects_Name(o))
        trvFontsByForm.Nodes.Add , , iOKey, mObjects_Name(o)
        If mObjects_ObjectOwnFont(o).FontName <> "" Then
            trvFontsByForm.Nodes.Add iOKey, tvwChild, , mObjects_Name(o) & ".Font: " & mObjects_ObjectOwnFont(o).FontName & "  " & mObjects_ObjectOwnFont(o).FontSize & " pt"
        End If
        If iFound Then
            Set iControlTypes = mObjects_ControlTypes(o)
            Set iControlNames = mObjects_ControlNames(o)
            For Each iType In mControlTypesGlobal
                iFound = False
                For i = 1 To iControlTypes.Count
                    iCtlType = iControlTypes(i)
                    If iCtlType = iType Then
                        Set iAuxPropsFont = iAuxControlPropertiesFont(i)
                        If iAuxPropsFont.Count > 0 Then
                            iFound = True
                            Exit For
                        End If
                    End If
                Next
                If iFound Then
                    iTKey = "t_" & mObjects_Name(o) & "_" & iType
                    trvFontsByForm.Nodes.Add iOKey, tvwChild, iTKey, iType
                    For i = 1 To iControlTypes.Count
                        iCtlType = iControlTypes(i)
                        If iCtlType = iType Then
                            Set iAuxPropsFont = iAuxControlPropertiesFont(i)
                            If iAuxPropsFont.Count > 0 Then
                                iCKey = "c_" & mObjects_Name(o) & "_" & iType & "_" & iControlNames(i)
                                trvFontsByForm.Nodes.Add iTKey, tvwChild, iCKey, iControlNames(i)
                                For Each iFont In iAuxPropsFont
                                    trvFontsByForm.Nodes.Add iCKey, tvwChild, , iFont.PropertyName & " Property" & ": " & iFont.FontName & "  " & iFont.FontSize & " pt"
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Next
    For i = 1 To trvFontsByForm.Nodes.Count
        If trvFontsByForm.Nodes(i).Parent Is Nothing Then
            trvFontsByForm.Nodes(i).Expanded = True
        End If
    Next
    If trvFontsByForm.Nodes.Count > 0 Then trvFontsByForm.Nodes(1).EnsureVisible

    picScanning.Line (0, 0)-(picScanning.ScaleWidth / 6 * 5, 60), vbGreen, BF
    picScanning.Line (picScanning.ScaleWidth / 6 * 5, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
    lblScanning2.Caption = "Loading: Fonts..."
    lblScanning2.Refresh
    picScanning.Refresh
    
    ' Fonts By Font
    trvFontsByFont.Nodes.Clear
    For Each iFont In mFontsGlobal
        iFKey = SimpleHash("f_" & iFont.FontName & "_" & iFont.FontSize)
        trvFontsByFont.Nodes.Add , , iFKey, iFont.FontName & "  " & iFont.FontSize & " pt"
        For o = 1 To UBound(mObjects_Name)
            iOKey = SimpleHash("f_" & iFont.FontName & "_" & iFont.FontSize) & "_" & mObjects_Name(o)
            Set iNode = trvFontsByFont.Nodes.Add(iFKey, tvwChild, iOKey, mObjects_Name(o))
            iNode.Tag = "o"
        Next o
    Next
    
    For o = 1 To UBound(mObjects_Name)
        Set iAuxControlPropertiesFont = mObjects_ControlPropertiesFont(o)
        For Each iAuxPropsFont In iAuxControlPropertiesFont
            If iAuxPropsFont.Count > 0 Then
                iFound = True
                Exit For
            End If
        Next
        Set iFont = mObjects_ObjectOwnFont(o)
        iOKey = SimpleHash("f_" & iFont.FontName & "_" & iFont.FontSize) & "_" & mObjects_Name(o)
        trvFontsByFont.Nodes.Add iOKey, tvwChild, , mObjects_Name(o) & ".Font"
        If iFound Then
            Set iControlTypes = mObjects_ControlTypes(o)
            Set iControlNames = mObjects_ControlNames(o)
            For Each iType In mControlTypesGlobal
                iFound = False
                For i = 1 To iControlTypes.Count
                    iCtlType = iControlTypes(i)
                    If iCtlType = iType Then
                        Set iAuxPropsFont = iAuxControlPropertiesFont(i)
                        If iAuxPropsFont.Count > 0 Then
                            iFound = True
                            Exit For
                        End If
                    End If
                Next
                If iFound Then
                    For i = 1 To iControlTypes.Count
                        iCtlType = iControlTypes(i)
                        If iCtlType = iType Then
                            Set iAuxPropsFont = iAuxControlPropertiesFont(i)
                            If iAuxPropsFont.Count > 0 Then
                                For Each iFont In iAuxPropsFont
                                    iOKey = SimpleHash("f_" & iFont.FontName & "_" & iFont.FontSize) & "_" & mObjects_Name(o)
                                    trvFontsByFont.Nodes.Add iOKey, tvwChild, , iControlNames(i) & "." & iFont.PropertyName & " (" & iType & ")"
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Next
    Set iAuxCol = New Collection
    For Each iNode In trvFontsByFont.Nodes
        If iNode.Tag = "o" Then
            If iNode.Children = 0 Then
                iAuxCol.Add iNode
            End If
        End If
    Next
    For Each iNode In iAuxCol
        trvFontsByFont.Nodes.Remove iNode.Index
    Next
    For i = 1 To trvFontsByFont.Nodes.Count
        If trvFontsByFont.Nodes(i).Parent Is Nothing Then
            trvFontsByFont.Nodes(i).Expanded = True
        End If
    Next
    If trvFontsByFont.Nodes.Count > 0 Then trvFontsByFont.Nodes(1).EnsureVisible
    
    cboControlType.Clear
    cboControlType.AddItem "[Please select]"
    For i = 1 To mControlTypesGlobal.Count
        cboControlType.AddItem mControlTypesGlobal(i)
    Next
    cboControlType.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mCanceled = True
    mUnloading = True
End Sub

Private Sub mProjects_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Dim c As Long
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
        If VBInstance.ActiveVBProject.Name <> mLastProjectName Then
            For c = 1 To sst1.Tabs - 1
                sst1.TabEnabled(c) = False
            Next
            sst1.Tag = sst1.Tab
            sst1.Tab = 0
        Else
            For c = 1 To sst1.Tabs - 1
                sst1.TabEnabled(c) = True
            Next
            sst1.Tab = Val(sst1.Tag)
        End If
    Else
        For c = 1 To sst1.Tabs - 1
            sst1.TabEnabled(c) = False
        Next
        sst1.Tag = sst1.Tab
        sst1.Tab = 0
    End If
End Sub

Private Sub optDepByDep_Click()
    trvDepByForm.Visible = False
    trvDepByDep.Visible = True
End Sub

Private Sub optDepByForm_Click()
    trvDepByForm.Visible = True
    trvDepByDep.Visible = False
End Sub

Private Sub optFontsByFont_Click()
    trvFontsByForm.Visible = False
    trvFontsByFont.Visible = True
End Sub

Private Sub optFontsByForm_Click()
    trvFontsByForm.Visible = True
    trvFontsByFont.Visible = False
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    Select Case sst1.Tab
        Case 0
            txtSummary.Move 100, 1870, sst1.Width / mDPIf - 200, sst1.Height / mDPIf - 1870 - 150 - 500
        Case 1
            trvDepByForm.Move 100, sst1.TabHeight + 100, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500
            trvDepByDep.Move 100, sst1.TabHeight + 100, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500
        Case 2
            trvStrings.Move 100, sst1.TabHeight + 100, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500
        Case 3
            trvFontsByFont.Move 100, sst1.TabHeight + 100, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500
            trvFontsByForm.Move 100, sst1.TabHeight + 100, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500
        Case 4
            picTabContainer(4).Width = sst1.Width - 100
            picTabContainer(4).Height = sst1.Height - sst1.TabHeight - 60
            trvFind.Move 100 - picTabContainer(4).Left, sst1.TabHeight + 1160, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500 - 1400
        Case 5
            ShowCurrentSelectionOnReplaceFont
            tmrRefrehcbo.Enabled = True
        Case 6
            picTabContainer(6).Width = sst1.Width - 100
            If Not VBInstance.ActiveVBProject Is Nothing Then
                If cboOrigObject.Tag <> VBInstance.ActiveVBProject.Name Then
                    FillcboOrigObject
                End If
            Else
                FillcboOrigObject
            End If
            ShowCurrentSelectionOnCopyControls
    End Select
    If sst1.Tab = 4 Then
        cmdCopy.Visible = trvFind.Visible
    Else
        cmdCopy.Visible = sst1.Tab < 4
    End If
    cmdCollapseTree.Visible = (sst1.Tab > 0) And (sst1.Tab < 4)
    picOptDep.Visible = sst1.Tab = 1
    picOptFonts.Visible = sst1.Tab = 3
    
End Sub

Private Sub LoadFontCombos()
    Dim c As Long
    
    cboOrigFontName.Clear
    cboOrigFontName.AddItem "[Please select]"
    cboOrigFontName.AddItem "MS Sans Serif"
    cboOrigFontName.AddItem "Arial"
    cboOrigFontName.AddItem "Tahoma"
    cboOrigFontName.AddItem "Microsoft Sans Serif"
    cboOrigFontName.AddItem "Segoe UI"
    
    For c = 0 To Screen.FontCount - 1
        Select Case Screen.Fonts(c)
            Case "MS Sans Serif", "Arial", "Tahoma", "Microsoft Sans Serif", "Segoe UI"
            Case Else
                cboOrigFontName.AddItem Screen.Fonts(c)
        End Select
    Next
    cboOrigFontName.ListIndex = 0
    
    cboOrigFontSize.Clear
    cboOrigFontSize.AddItem "*"
    cboOrigFontSize.AddItem "8"
    cboOrigFontSize.AddItem "9"
    cboOrigFontSize.AddItem "10"
    cboOrigFontSize.AddItem "11"
    cboOrigFontSize.AddItem "12"
    cboOrigFontSize.AddItem "14"
    cboOrigFontSize.AddItem "16"
    cboOrigFontSize.AddItem "18"
    cboOrigFontSize.AddItem "20"
    cboOrigFontSize.AddItem "22"
    cboOrigFontSize.AddItem "24"
    cboOrigFontSize.Text = "*"
    
    cboOrigFontSize.ListIndex = 0
    
    cboNewFontName.Clear
    cboNewFontName.AddItem "[Please select]"
    cboNewFontName.AddItem "MS Sans Serif"
    cboNewFontName.AddItem "Arial"
    cboNewFontName.AddItem "Tahoma"
    cboNewFontName.AddItem "Microsoft Sans Serif"
    cboNewFontName.AddItem "Segoe UI"
    
    For c = 0 To Screen.FontCount - 1
        Select Case Screen.Fonts(c)
            Case "MS Sans Serif", "Arial", "Tahoma", "Microsoft Sans Serif", "Segoe UI"
            Case Else
                cboNewFontName.AddItem Screen.Fonts(c)
        End Select
    Next
    cboNewFontName.ListIndex = 0
    
    cboNewFontSize.Clear
    cboNewFontSize.AddItem "8"
    cboNewFontSize.AddItem "9"
    cboNewFontSize.AddItem "10"
    cboNewFontSize.AddItem "11"
    cboNewFontSize.AddItem "12"
    cboNewFontSize.AddItem "14"
    cboNewFontSize.AddItem "16"
    cboNewFontSize.AddItem "18"
    cboNewFontSize.AddItem "20"
    cboNewFontSize.AddItem "22"
    cboNewFontSize.AddItem "24"
    cboNewFontSize.Text = ""
    
End Sub

Private Sub ParseTree(nNode As Node)
    mTreeText = mTreeText & String$(mIndentLevel, vbTab) & nNode.Text & vbCrLf
    If nNode.Children > 0 Then
        mIndentLevel = mIndentLevel + 1
        ParseTree nNode.Child
    End If
    Set nNode = nNode.Next
    If TypeName(nNode) <> "Nothing" Then
        ParseTree nNode
    Else
        mIndentLevel = mIndentLevel - 1
    End If
End Sub

Private Function SimpleHash(ByVal nData As Variant, Optional pNumberOfHasCharacters_MustBeEvenAndLessThan514 As Long = 38) As String
    Dim iHashBytes() As Byte
    Dim c As Long
    Dim n As Long
    Dim iStr As String
    Dim iVarType As Long
    Dim iDataBytes() As Byte

    If pNumberOfHasCharacters_MustBeEvenAndLessThan514 Mod 2 <> 0 Then Err.Raise 1142, App.Title & ".SimpleHash", "pNumberOfHasCharacters_MustBeEvenAndLessThan514 must be even."
    If pNumberOfHasCharacters_MustBeEvenAndLessThan514 < 2 Then Err.Raise 1142, App.Title & ".SimpleHash", "pNumberOfHasCharacters_MustBeEvenAndLessThan514 must 2 or more."
    If pNumberOfHasCharacters_MustBeEvenAndLessThan514 > 512 Then Err.Raise 1142, App.Title & ".SimpleHash", "pNumberOfHasCharacters_MustBeEvenAndLessThan514 must 512 or less."

    n = (pNumberOfHasCharacters_MustBeEvenAndLessThan514 / 2)
    ReDim iHashBytes(n - 1)
    iVarType = VarType(nData)
    If iVarType = vbString Then
        iStr = nData
        HashData StrPtr(iStr), 2 * Len(iStr), iHashBytes(0), n
    Else
        If iVarType <> vbArray + vbByte Then
            Err.Raise 2345, , "Invalid data type"
            Exit Function
        Else
            iDataBytes = nData
            HashData VarPtr(iDataBytes(0)), UBound(iDataBytes) + 1, iHashBytes(0), n
        End If
    End If
    For c = 0 To UBound(iHashBytes)
        iStr = Hex$(iHashBytes(c))
        If Len(iStr) = 1 Then
            iStr = "0" & iStr
        End If
        SimpleHash = SimpleHash & iStr
    Next c
    'SimpleHash = Right$("00000000000000" & Hex$(iHashBytes.l1) & Hex$(iHashBytes.l2), 16)
End Function

Private Function GetControlName(nCtrl As Object) As String
    Dim iIndex As Long
    
    GetControlName = nCtrl.Name
    iIndex = -1
    On Error Resume Next
    iIndex = nCtrl.Index
    On Error GoTo 0
    If iIndex > -1 Then
        GetControlName = GetControlName & "(" & iIndex & ")"
    End If
End Function

Private Sub tmrRefrehcbo_Timer()
'    cboOrigFontSize.SelStart = 0
    cboOrigFontSize.SelLength = 0
    cboOrigFontSize.Refresh
    cboNewFontSize.Refresh
    tmrRefrehcbo.Enabled = False
End Sub

Private Sub EnableDisableFontReplacementButton()
    cmdReplaceFonts.Enabled = False
    If cboOrigFontName.ListIndex > 0 Then
        If cboNewFontName.ListIndex > 0 Then
            If Val(cboOrigFontSize.Text) = 0 Then
                cmdReplaceFonts.Enabled = (cboOrigFontName.ListIndex <> cboNewFontName.ListIndex)
            ElseIf Val(cboNewFontSize.Text) <> Val(cboOrigFontSize.Text) Then
                cmdReplaceFonts.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub DoReplaceFonts()
    Dim iComp As VBComponent
    Dim pc As Long
    Dim p As Long
    Dim iCtl As VBControl
    Dim pr As Long
    Dim iVar As Variant
    Dim iObj As Object
    Dim iDes As Object
    Dim iProp As Property
    Dim iVarType As Long
    Dim iOrigFontName As String
    Dim iNewFontName As String
    Dim iOrigFontSize As Long
    Dim iNewFontSize As Long
    Dim iChangeFontName As Boolean
    Dim iChangeFontSize As Boolean
    Dim p2 As Long
    Dim iChanged As Boolean
    Dim iCount As Long
    Dim iDesignerWindowVisible As Boolean
    Dim iIsDirty As Boolean
    Dim iFontsReplacedInComp As Boolean
    
    iOrigFontName = cboOrigFontName.Text
    iNewFontName = cboNewFontName.Text
    iOrigFontSize = Val(cboOrigFontSize.Text)
    iNewFontSize = Val(cboNewFontSize.Text)
    iChangeFontName = iOrigFontName <> iNewFontName
    iChangeFontSize = cboNewFontSize.Text <> ""
    
    EnableOtherTabs False
    On Error GoTo ErrorExit
    
    For Each iComp In mProject.VBComponents
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            pc = pc + 1
        End If
        Set iComp = Nothing
    Next
    
    lblScanning.Caption = ""
    lblScanning2.Caption = ""
    lblScanning3.Caption = ""
    picScanning.Visible = True
    mReplacingFonts = True
    mCanceled = False
    For Each iComp In mProject.VBComponents
        'Debug.Print iComp.Name
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            If ObjectSelected(CVar(iComp.Name)) Then
                iFontsReplacedInComp = False
                p = p + 1
                picScanning.Line (0, 0)-(picScanning.ScaleWidth / pc * p, 60), vbGreen, BF
                picScanning.Line (picScanning.ScaleWidth / pc * p, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
                lblScanning.Caption = "Scanning: " & iComp.Name & "..."
                lblScanning.Refresh
                DoEvents
                If mCanceled Then
                    mReplacingFonts = False
                    picScanning.Visible = False
                    Exit Sub
                End If
                
                iIsDirty = iComp.IsDirty
                iDesignerWindowVisible = iComp.DesignerWindow.Visible
                Set iDes = iComp.Designer
                For Each iCtl In iDes.VBControls
                    If ControlTypeSelected(CVar(iCtl.ProgId)) Then
                        lblScanning2.Caption = "Control: " & iCtl.ControlObject.Name & " type: " & TypeName(iCtl.ControlObject)
                        lblScanning2.Refresh
                        If iCtl.ControlObject Is Nothing Then
                            mCanceled = True
                            mReplacingFonts = False
                            picScanning.Visible = False
                            Exit Sub
                        End If
                        For Each iProp In iCtl.Properties
                            If FontPropertySelected(iProp.Name) Then
                                lblScanning3.Caption = "Property: " & iProp.Name
                                lblScanning3.Refresh
                                pr = pr + 1
                                If pr > 1000 Then
                                    pr = 0
                                    DoEvents
                                    If mCanceled Then
                                        mReplacingFonts = False
                                        picScanning.Visible = False
                                        Exit Sub
                                    End If
                                End If
                                iVar = Empty
                                On Error Resume Next
                                iVar = iProp
                                On Error GoTo ErrorExit
                                iVarType = VarType(iVar)
                                If iVarType = vbObject Then
                                    Set iObj = Nothing
                                    On Error Resume Next
                                    Set iObj = iProp.object
                                    On Error GoTo ErrorExit
                                    If Not iObj Is Nothing Then
                                        If TypeName(iObj) = "Font" Then
                                            If iObj.Name = iOrigFontName Then
                                                iChanged = False
                                                If iChangeFontName Then
                                                    iObj.Name = iNewFontName
                                                    iChanged = True
                                                End If
                                                If iChangeFontSize Then
                                                    If Round(iObj.Size) = iOrigFontSize Then
                                                        iObj.Size = iNewFontSize
                                                        iChanged = True
                                                    End If
                                                End If
                                                If iChanged Then
                                                    iCount = iCount + 1
                                                    iFontsReplacedInComp = True
                                                End If
                                            End If
                                        End If
                                    End If
                                    Set iObj = Nothing
                                End If
                            End If
                            iVar = Empty
                        Next
                        Set iProp = Nothing
                    End If
                Next
                Set iCtl = Nothing
                
                If chkFontOfObject.Value Then
                    Set iProp = Nothing
                    For p2 = 1 To iComp.Properties.Count
                        Set iProp = iComp.Properties(p2)
                        If iProp.Name = "Font" Then
                            Set iObj = Nothing
                            On Error Resume Next
                            Set iObj = iProp.object
                            On Error GoTo ErrorExit
                            If Not iObj Is Nothing Then
                                If TypeName(iProp.object) = "Font" Then
                                    If iObj.Name = iOrigFontName Then
                                        iChanged = False
                                        If iChangeFontName Then
                                            iObj.Name = iNewFontName
                                            iChanged = True
                                        End If
                                        If iChangeFontSize Then
                                            If Round(iObj.Size) = iOrigFontSize Then
                                                iObj.Size = iNewFontSize
                                                iChanged = True
                                            End If
                                        End If
                                        If iChanged Then
                                            iCount = iCount + 1
                                            iFontsReplacedInComp = True
                                        End If
                                    End If
                                End If
                            End If
                            Set iObj = Nothing
                        End If
                    Next
                    Set iProp = Nothing 'Bug in the Add-In environment, if not set to Nothing VB chashes with UserControls when the Add-In is compiled
                End If
                If Not iFontsReplacedInComp Then
                    If iComp.IsDirty And (Not iIsDirty) Then
                        On Error Resume Next
                        iComp.Reload
                        On Error GoTo ErrorExit
                    End If
                Else
                    iComp.IsDirty = True
                End If
                If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
                Set iDes = Nothing
            End If
        End If
        Set iComp = Nothing
    Next
    mReplacingFonts = False
    MsgBox iCount & " fonts replaced." & vbCrLf & "To see them reflected on the treeviews run the scan again from the first tab.", vbInformation
    Exit Sub
    
ErrorExit:
    MsgBox "Error. Canceled", vbCritical
    mReplacingFonts = False
    mCanceled = True
    picScanning.Visible = False
End Sub
    
Private Sub FindControls()
    Dim iComp As VBComponent
    Dim pc As Long
    Dim p As Long
    Dim iCtl As VBControl
    Dim iVar As Variant
    Dim iObj As Object
    Dim iDes As Object
    Dim iProp As Property
    Dim iVarType As Long
    Dim iCount As Long
    Dim iControlType As String
    Dim iPropertyToCompare As String
    Dim iCompareCondition As String
    Dim iCompareToString As String
    Dim iCompareToNumber As Double
    Dim iMatchCriteria As Boolean
    Dim iOKey As String
    Dim i As Long
    Dim cc As Long
    Dim iIgnoreCase As Boolean
    Dim iStrValue As String
    Dim iDesignerWindowVisible As Boolean
    Dim iIsDirty As Boolean
    Dim iCriteria As Long
    
    iCriteria = cboCriteria.ListIndex
    If iCriteria = 0 Then ' list all
        mFindCriteria = "controls " & cboControlType.Text
    ElseIf iCriteria = 1 Then ' property value
        mFindCriteria = "controls " & cboControlType.Text & " with property " & cboPropertyToCompare.Text & " " & cboPropertyValueCondition.Text & " " & IIf(cboPropertyToCompare.ItemData(cboPropertyToCompare.ListIndex) = 2, """", "") & txtPropertyValue.Text & IIf(cboPropertyToCompare.ItemData(cboPropertyToCompare.ListIndex) = 2, """", "")
    ElseIf iCriteria = 2 Then ' property value
        mFindCriteria = "controls " & cboControlType.Text & " with property " & cboPropertyToCompare.Text & " " & cboPropertyValueCondition.Text & " " & cboPropertyToCompare2.Text
    End If
    
    iControlType = cboControlType.Text
    iPropertyToCompare = cboPropertyToCompare.Text
    iCompareCondition = cboPropertyValueCondition.Text
    iIgnoreCase = chkIgnoreCase.Value
    If Val(cboPropertyValueCondition.Tag) = eptNumeric Then
        iCompareToNumber = Val(txtPropertyValue.Text)
    Else
        If iIgnoreCase Then
            iCompareToString = LCase$(txtPropertyValue.Text)
        Else
            iCompareToString = txtPropertyValue.Text
        End If
    End If
    
    EnableOtherTabs False
    On Error GoTo ErrorExit
    
    For Each iComp In mProject.VBComponents
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            pc = pc + 1
        End If
        Set iComp = Nothing
    Next
    
    lblScanning.Caption = ""
    lblScanning2.Caption = ""
    lblScanning3.Caption = ""
    picScanning.Visible = True
    mFinding = True
    mCanceled = False
    trvFind.Visible = False
    trvFind.Nodes.Clear
    For Each iComp In mProject.VBComponents
        'Debug.Print iComp.Name
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            iOKey = ""
            p = p + 1
            picScanning.Line (0, 0)-(picScanning.ScaleWidth / pc * p, 60), vbGreen, BF
            picScanning.Line (picScanning.ScaleWidth / pc * p, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
            lblScanning.Caption = "Scanning: " & iComp.Name & "..."
            lblScanning.Refresh
            DoEvents
            If mCanceled Then
                mFinding = False
                picScanning.Visible = False
                Exit Sub
            End If
            
            iIsDirty = iComp.IsDirty
            iDesignerWindowVisible = iComp.DesignerWindow.Visible
            Set iDes = iComp.Designer
            For Each iCtl In iDes.VBControls
                If iCtl.ProgId = iControlType Then
                    lblScanning2.Caption = "Control: " & iCtl.ControlObject.Name ' & " type: " & TypeName(iCtl.ControlObject)
                    lblScanning2.Refresh
                    If iCriteria = 0 Then ' list all
                        If iOKey = "" Then
                            iOKey = SimpleHash("o_" & iComp.Name)
                            trvFind.Nodes.Add , , iOKey, iComp.Name
                        End If
                        trvFind.Nodes.Add iOKey, tvwChild, , iCtl.ControlObject.Name
                        cc = cc + 1
                    Else
                        If iCriteria = 2 Then ' compare to other property
                            If Val(cboPropertyValueCondition.Tag) = eptNumeric Then
                                iCompareToNumber = Val(iCtl.Properties(cboPropertyToCompare2.Text))
                            Else
                                If iIgnoreCase Then
                                    iCompareToString = LCase$(iCtl.Properties(cboPropertyToCompare2.Text))
                                Else
                                    iCompareToString = iCtl.Properties(cboPropertyToCompare2.Text)
                                End If
                            End If
                        End If
                        For Each iProp In iCtl.Properties
                            If iProp.Name = iPropertyToCompare Then
                            '    lblScanning3.Caption = "Property: " & iProp.Name
                             '   lblScanning3.Refresh
                                iVar = Empty
                                On Error Resume Next
                                iVar = iProp
                                On Error GoTo ErrorExit
                                iVarType = VarType(iVar)
                                
                                iMatchCriteria = False
                                If iVarType = vbString Then
                                    If iIgnoreCase Then
                                        iStrValue = Replace$(LCase$(iProp.Value), "&", "")
                                    Else
                                        iStrValue = Replace$(iProp.Value, "&", "")
                                    End If
                                    Select Case iCompareCondition
                                        Case "="
                                            iMatchCriteria = iStrValue = iCompareToString
                                        Case "<>"
                                            iMatchCriteria = iStrValue <> iCompareToString
                                        Case "contains"
                                            iMatchCriteria = InStr(iStrValue, iCompareToString)
                                        Case "is contained in"
                                            iMatchCriteria = InStr(iCompareToString, iStrValue)
                                    End Select
                                Else ' numeric
                                    Select Case iCompareCondition
                                        Case "="
                                            iMatchCriteria = iProp.Value = iCompareToNumber
                                        Case "<>"
                                            iMatchCriteria = iProp.Value <> iCompareToNumber
                                        Case ">"
                                            iMatchCriteria = iProp.Value > iCompareToNumber
                                        Case ">="
                                            iMatchCriteria = iProp.Value >= iCompareToNumber
                                        Case "<"
                                            iMatchCriteria = iProp.Value < iCompareToNumber
                                        Case "<="
                                            iMatchCriteria = iProp.Value <= iCompareToNumber
                                    End Select
                                End If
                                If iMatchCriteria Then
                                    If iOKey = "" Then
                                        iOKey = SimpleHash("o_" & iComp.Name)
                                        trvFind.Nodes.Add , , iOKey, iComp.Name
                                    End If
                                    If iCriteria = 1 Then
                                        If iVarType = vbString Then
                                            trvFind.Nodes.Add iOKey, tvwChild, , iCtl.ControlObject.Name & "." & iProp.Name & " = """ & iProp.Value & """"
                                        Else
                                            trvFind.Nodes.Add iOKey, tvwChild, , iCtl.ControlObject.Name & "." & iProp.Name & " = " & iProp.Value
                                        End If
                                    ElseIf iCriteria = 2 Then
                                        If iVarType = vbString Then
                                            trvFind.Nodes.Add iOKey, tvwChild, , iCtl.ControlObject.Name & "." & iProp.Name & " = """ & iProp.Value & """, " & cboPropertyToCompare2.Text & " = """ & iCompareToString & """"
                                        Else
                                            trvFind.Nodes.Add iOKey, tvwChild, , iCtl.ControlObject.Name & "." & iProp.Name & " = " & iProp.Value & ", " & cboPropertyToCompare2.Text & " = " & iCompareToNumber
                                        End If
                                    End If
                                    cc = cc + 1
                                    Exit For
                                End If
                            End If
                            iVar = Empty
                        Next
                        Set iProp = Nothing
                    End If
                End If
            Next
            Set iCtl = Nothing
            If iComp.IsDirty And (Not iIsDirty) Then
                On Error Resume Next
                iComp.Reload
                On Error GoTo ErrorExit
            End If
            If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
            Set iDes = Nothing
        End If
        Set iComp = Nothing
    Next
    For i = 1 To trvFind.Nodes.Count
        If trvFind.Nodes(i).Parent Is Nothing Then
            trvFind.Nodes(i).Expanded = True
        End If
    Next
    If trvFind.Nodes.Count > 0 Then trvFind.Nodes(1).EnsureVisible
    
    picScanning.Visible = False
    trvFind.Move 100 - picTabContainer(4).Left, sst1.TabHeight + 1160, sst1.Width - 200, sst1.Height - 260 - sst1.TabHeight - 500 - 1400
    If cc > 0 Then
        MsgBox cc & " control" & IIf(cc = 1, "", "s") & " found.", vbInformation
    Else
        MsgBox "No controls found.", vbInformation
    End If
    Exit Sub
    
ErrorExit:
    MsgBox "Error. Canceled", vbCritical
    mFinding = False
    mCanceled = True
    picScanning.Visible = False
    
End Sub

Private Sub FillcboOrigObject()
    Dim iComp As VBComponent
    
    cboOrigObject.Clear
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Screen.MousePointer = vbHourglass
        cboOrigObject.Tag = mProject.Name
        For Each iComp In mProject.VBComponents
            If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
                cboOrigObject.AddItem iComp.Name
            End If
            Set iComp = Nothing
        Next
        cboOrigObject.AddItem "[Please select]", 0
        cboOrigObject.ListIndex = 0
        Screen.MousePointer = vbDefault
    End If
    cboOrigControl.Clear
End Sub

Private Sub DoAddControls()
    Dim iComp As VBComponent
    Dim iForm As VBForm
    Dim iCtl As VBControl
    Dim iControlToCopy As VBControl
    Dim iNewControl As VBControl
    Dim iPropOrig As Property
    Dim iPropNew As Property
    Dim iVar As Variant
    Dim iVarType As Long
    Dim iObj As Object
    Dim cc As Long
    Dim i As Long
    Dim pc As Long
    Dim p As Long
    Dim nf As Object
    Dim iDesignerWindowVisible As Boolean
    Dim iIsDirty As Boolean
    Dim iControlAddedToComp As Boolean
    
    For Each iComp In mProject.VBComponents
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            If iComp.Name = cboOrigObject.Text Then
                iIsDirty = iComp.IsDirty
                iDesignerWindowVisible = iComp.DesignerWindow.Visible
                Set iForm = iComp.Designer
                For Each iCtl In iForm.VBControls
                    If iCtl.ControlObject.Name = Left$(cboOrigControl.Text, InStr(cboOrigControl.Text, " ") - 1) Then
                        Set iControlToCopy = iCtl
                    End If
                Next
                Set iCtl = Nothing
                If iComp.IsDirty And (Not iIsDirty) Then
                    On Error Resume Next
                    iComp.Reload
                    On Error GoTo 0
                End If
                If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
                Set iForm = Nothing
            End If
            pc = pc + 1
        End If
        Set iComp = Nothing
    Next
    
    Set iComp = Nothing
    Set iForm = Nothing
    Set iCtl = Nothing
    If iControlToCopy Is Nothing Then
        MsgBox "Original control not found.", vbCritical
        Exit Sub
    End If
    
    EnableOtherTabs False
    lblScanning.Caption = ""
    lblScanning2.Caption = ""
    lblScanning3.Caption = ""
    picScanning.Visible = True
    mCopyingControls = True
    mCanceled = False
    lblPicturepropertiesNote.Visible = False
    
    For Each iComp In mProject.VBComponents
        If (iComp.Type = vbext_ct_VBForm) Or (iComp.Type = vbext_ct_UserControl) Or (iComp.Type = vbext_ct_VBMDIForm) Then
            p = p + 1
            If (iComp.Name <> cboOrigObject.Text) And ObjectSelected(CVar(iComp.Name)) Then
                iControlAddedToComp = False
                picScanning.Line (0, 0)-(picScanning.ScaleWidth / pc * p, 60), vbGreen, BF
                picScanning.Line (picScanning.ScaleWidth / pc * p, 0)-(picScanning.ScaleWidth, 60), vbWhite, BF
                lblScanning.Caption = "Adding to: " & iComp.Name & "..."
                lblScanning.Refresh
                DoEvents
                If mCanceled Then
                    mCopyingControls = False
                    picScanning.Visible = False
                    Exit Sub
                End If
                iIsDirty = iComp.IsDirty
                iDesignerWindowVisible = iComp.DesignerWindow.Visible
                Set iForm = iComp.Designer
                Set iNewControl = Nothing
                On Error Resume Next
                Set iNewControl = iForm.VBControls.Add(iControlToCopy.ProgId)
                On Error GoTo ErrorExit
                If Not iNewControl Is Nothing Then
                    iControlAddedToComp = True
                    If Not ControlNameExistsInForm(iForm, iControlToCopy.ControlObject.Name) Then
                        iNewControl.ControlObject.Name = iControlToCopy.ControlObject.Name
                    End If
                    cc = cc + 1
                    For i = 1 To iControlToCopy.Properties.Count
                        Set iPropOrig = iControlToCopy.Properties(i)
                        Set iPropNew = iNewControl.Properties(iPropOrig.Name)
                        On Error Resume Next
                        Err.Clear
                        If iPropNew.Value <> iPropOrig.Value Then
                            If Err.Number = 0 Then
                                iPropNew.Value = iPropOrig.Value
                            End If
                        End If
                        On Error GoTo ErrorExit
                        iVar = Empty
                        On Error Resume Next
                        iVar = iPropOrig
                        On Error GoTo ErrorExit
                        iVarType = VarType(iVar)
                        If iVarType = vbObject Then
                            Set iObj = Nothing
                            On Error Resume Next
                            Set iObj = iPropOrig.object
                            On Error GoTo ErrorExit
                            If Not iObj Is Nothing Then
                                If TypeName(iObj) = "Font" Then
                                    Set nf = iPropNew.object
                                    nf.Name = iObj.Name
                                    nf.Size = iObj.Size
                                    nf.Bold = iObj.Bold
                                    nf.Charset = iObj.Charset
                                    nf.Italic = iObj.Italic
                                    nf.Strikethrough = iObj.Strikethrough
                                    nf.Underline = iObj.Underline
                                    nf.Weight = iObj.Weight
                                End If
                            End If
                        End If
                    Next
                End If
                If Not iControlAddedToComp Then
                    If iComp.IsDirty And (Not iIsDirty) Then
                        On Error Resume Next
                        iComp.Reload
                        On Error GoTo ErrorExit
                    End If
                End If
                If Not iDesignerWindowVisible Then iComp.DesignerWindow.Close
                Set iForm = Nothing
            End If
        End If
        Set iComp = Nothing
    Next
    mCopyingControls = False
    picScanning.Visible = False
    MsgBox cc & " controls added.", vbInformation
    Exit Sub
    
ErrorExit:
    MsgBox "Error. Canceled", vbCritical
    mCopyingControls = False
    mCanceled = True
    picScanning.Visible = False
End Sub

Private Function ControlNameExistsInForm(nForm As VBForm, ByVal nName As String) As Boolean
    Dim iCtl As VBControl
    
    nName = LCase$(nName)
    For Each iCtl In nForm.VBControls
        If LCase$(iCtl.ControlObject.Name) = nName Then
            ControlNameExistsInForm = True
            Exit For
        End If
    Next
End Function

Private Sub EnableOtherTabs(nValue As Boolean)
    Dim c As Long
    
    If nValue Then
        For c = 0 To sst1.Tabs - 1
            sst1.TabEnabled(c) = True
        Next
    Else
        For c = 0 To sst1.Tabs - 1
            If c <> sst1.Tab Then
                sst1.TabEnabled(c) = False
            End If
        Next
    End If
End Sub

Private Function GetTrueTwipsPerPixelX() As Single
    Dim h As Long
    Const LOGPIXELSX As Long = 88
    
    h = GetDC(0)
    If h <> 0 Then
        GetTrueTwipsPerPixelX = 1440 / GetDeviceCaps(h, LOGPIXELSX)
        ReleaseDC 0, h
    Else
        GetTrueTwipsPerPixelX = Screen.TwipsPerPixelX
    End If
End Function

Private Property Get InIDE() As Boolean
    Debug.Assert MakeTrue(InIDE)
End Property
 
Private Function MakeTrue(bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

