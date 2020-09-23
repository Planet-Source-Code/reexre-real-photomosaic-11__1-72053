VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FolderBrowser 
   Appearance      =   0  'Flat
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   PropertyPages   =   "Browse.ctx":0000
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ToolboxBitmap   =   "Browse.ctx":0024
   Begin VB.PictureBox picBorder 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   3975
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3975
      Begin VB.Frame fraDetails 
         Caption         =   "Selected Folder: "
         Height          =   1455
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Information about currently selected Folder or File (Double-click here to show more details)"
         Top             =   2400
         Width           =   3975
         Begin VB.TextBox txtDetails 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Information about currently selected Folder or File (Double-click here to show more details)"
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.CheckBox chkShowNetwork 
         Caption         =   "Show &Network"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Show network drives in Folder Browser."
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton btnNewFolder 
         Caption         =   "New Folder ..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkShowFiles 
         Caption         =   "Show F&iles"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Show files in Folder Browser (only Folders can be selected)"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.TreeView tvBrowse 
         Height          =   1575
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2778
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ilBrowse"
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList ilBrowse 
         Left            =   1440
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":0336
               Key             =   "drive-net"
               Object.Tag             =   "drive-net"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":0739
               Key             =   "computer"
               Object.Tag             =   "computer"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":0B27
               Key             =   "domain"
               Object.Tag             =   "domain"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":0F0F
               Key             =   "drive"
               Object.Tag             =   "drive"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":12C6
               Key             =   "world"
               Object.Tag             =   "world"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":1718
               Key             =   "network"
               Object.Tag             =   "network"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":1B45
               Key             =   "file-hidden"
               Object.Tag             =   "file-hidden"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":1F19
               Key             =   "file"
               Object.Tag             =   "file"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":22ED
               Key             =   "folder-shared"
               Object.Tag             =   "folder-shared"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":26E6
               Key             =   "folder-hidden"
               Object.Tag             =   "folder-hidden"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":2AC1
               Key             =   "folder-hidden-shared"
               Object.Tag             =   "folder-hidden-shared"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":2EB6
               Key             =   "provider"
               Object.Tag             =   "provider"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Browse.ctx":32CC
               Key             =   "folder"
               Object.Tag             =   "folder"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBrowse 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a &Folder:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyFolder 
         Caption         =   "Copy Folder Name"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyFile 
         Caption         =   "Copy File Name"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuNewFolder 
         Caption         =   "Create New &Folder Here"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "FolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
' Copyright © 2004-2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/
' Version 1.15 7/23/2005

Private Const BROWSER_DESC = "Karen's Folder Browser"
Private Const BROWSER_COPYRIGHT = "Copyright © 2004-2005 Karen Kenworthy\nAll Rights Reserved"
Private Const BROWSER_VER = "1.15"
Private Const BROWSER_COMMENTS = "This is Karen's Folder Browser, discussed in her Power Tools Newsletter.\n\nIt allows you to select any accessible Folder or File, on a local drive or a network share. It can optionally display hidden files or folders."

'Private Const SEL_TIP = "Information about currently selected Folder or File -- %SEL%"
Private Const SEL_TIP = "%SEL%"

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3
Private Const RESOURCE_RECENT = &H4
Private Const RESOURCE_CONTEXT = &H5

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_RESERVED = &H8
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_NOLOCALDEVICE = &H4
Private Const RESOURCEUSAGE_ATTACHED = &H10
Private Const RESOURCEUSAGE_SIBLING = &H8
'Private Const RESOURCEUSAGE_ALL = &H0
Private Const RESOURCEUSAGE_ALL = (RESOURCEUSAGE_CONNECTABLE Or RESOURCEUSAGE_CONTAINER Or RESOURCEUSAGE_ATTACHED)
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEDISPLAYTYPE_FILE = &H4
Private Const RESOURCEDISPLAYTYPE_GROUP = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT = &H7
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
Private Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9
Private Const RESOURCEDISPLAYTYPE_TREE = &HA
Private Const RESOURCEDISPLAYTYPE_NDSCONTAINER = &HB

Private Const NO_ERROR = 0
Private Const ERROR_NO_NETWORK = 1222
Private Const ERROR_INVALID_HANDLE = 6
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_EXTENDED_ERROR = 1208
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const NETRES_BUFLEN = 2048

Private Const MWN_NAME = "Microsoft Windows Network"

Private KEY_FOLDER As String
Private MWN_PRESENT As Boolean

Private Const KEY_LOCALROOT = "L"
Private Const KEY_SPECIALROOT = "S"
Private Const KEY_NETROOT = "N"
Private Const KEY_CONTAINER = "C"
Private Const KEY_CONTAINER_MWN = "CMWN"
Private Const KEY_FOLDER_DEF = "F"
Private Const KEY_FOLDER_ALT = "G"
Private Const KEY_FOLDER_HIDDEN = "H"
Private Const KEY_SPECIAL = "I"
Private Const KEY_FILE = "X"
Private Const KEY_DUMMY = "D"

Private Const IMAGE_NETROOT = "network"
Private Const IMAGE_PROVIDER = "provider"
Private Const IMAGE_NETWORK = "network"
Private Const IMAGE_DOMAIN = "domain"
Private Const IMAGE_SERVER = "computer"
Private Const IMAGE_COMPUTER = "computer"
Private Const IMAGE_DRIVE_NET = "drive-net"
Private Const IMAGE_LOCALROOT = "computer"
Private Const IMAGE_DRIVE = "drive"
Private Const IMAGE_FOLDER = "folder"
Private Const IMAGE_FOLDER_HIDDEN = "folder-hidden"
Private Const IMAGE_FOLDER_SHARED = "folder-shared"
Private Const IMAGE_FOLDER_HIDDEN_SHARED = "folder-hidden-shared"
Private Const IMAGE_FILE = "file"
Private Const IMAGE_FILE_HIDDEN = "file-hidden"
Private Const IMAGE_DUMMY = "dummy"

Private Const TAG_NETROOT = "neighborhood"
Private Const TAG_NETWORK = "network"
Private Const TAG_DOMAIN = "domain"
Private Const TAG_SERVER = "server"
Private Const TAG_LOCALROOT = "local"
Private Const TAG_FOLDER = "folder"
Private Const TAG_FOLDER_HIDDEN = "folder-hidden"
Private Const TAG_SPECIAL = "special"
Private Const TAG_FILE = "file"
Private Const TAG_DUMMY = "dummy"

Public Enum BROWSE_ERROR
    PROPERTY_READONLY = vbObjectError + 1
    SERIALIZE_EMPTY
    SERIALIZE_INVALID
    SERIALIZE_BAD_VERSION
    JOBTYPE_INVALID
End Enum

Private Enum SERIAL_POS
    POS_SVERSION = 0
    POS_Enabled
    POS_Visible
    POS_ShowNetwork
    POS_ShowLocal
    POS_ShowFiles
    POS_Folder
    POS_File
    POS_ShowFilesVisible
    POS_ShowNetworkVisible
    POS_NewFolderVisible
    POS_FileFilter
    POS_FolderFilter
    POS_ShowHidden
    POS_ShowSystem
    POS_ShowReadOnly

    POS_Width
    POS_Height
    POS_Border
    POS_Caption
    POS_Appearance
    POS_BackStyle

    POS_FontName
    POS_FontSize
    POS_FontBold
    POS_FontItalic
    POS_FontUnderline
    POS_FontStrikethru
    POS_ForeColor
    POS_Dummy1
    POS_Dummy2
    POS_Dummy3
    POS_Dummy4
    POS_Dummy5

    POS_UBOUND = POS_Dummy5
End Enum

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Private Type NETRESOURCEA
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment  As Long
    lpProvider As Long
End Type

Private Type NETRESOURCEW
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment  As Long
    lpProvider As Long
End Type

Private Enum PT_NET_TYPE
    NET_TYPE_GENERIC = &H0
    NET_TYPE_DOMAIN = &H1
    NET_TYPE_SERVER = &H2
    NET_TYPE_SHARE = &H3
    NET_TYPE_FILE = &H4
    NET_TYPE_GROUP = &H5
    NET_TYPE_NETWORK = &H6
    NET_TYPE_ROOT = &H7
    NET_TYPE_SHAREADMIN = &H8
    NET_TYPE_DIRECTORY = &H9
    NET_TYPE_TREE = &HA
    NET_TYPE_NDSCONTAINER = &HB
End Enum

Private Enum PT_NET_DISP_TYPE
    NET_DISP_TYPE_GENERIC = &H0
    NET_DISP_TYPE_DOMAIN = &H1
    NET_DISP_TYPE_SERVER = &H2
    NET_DISP_TYPE_SHARE = &H3
    NET_DISP_TYPE_FILE = &H4
    NET_DISP_TYPE_GROUP = &H5
    NET_DISP_TYPE_NETWORK = &H6
    NET_DISP_TYPE_ROOT = &H7
    NET_DISP_TYPE_SHAREADMIN = &H8
    NET_DISP_TYPE_DIRECTORY = &H9
    NET_DISP_TYPE_TREE = &HA
    NET_DISP_TYPE_NDSCONTAINER = &HB
End Enum

Private Type PT_NET_ITEM
    IsContainer As Boolean
    Type As PT_NET_TYPE
    Provider As String
'    Network As String
'    Domain As String
'    Server As String
    FullPath As String
    RemoteName As String
    LocalName As String
    DispType As PT_NET_DISP_TYPE
    Usage As Long
    Spec As String ' spec = provider & vbtab & container & vbtab & displaytype & vbtab & usage
End Type

Private Type PT_NET_INFO
    Cnt As Long
    Resource() As PT_NET_ITEM
End Type

Private Type BROWSE_SORT
    Key As String
    Hidden As Boolean
End Type

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
    Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function GetDriveType Lib "kernel32" _
    Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function WNetOpenEnumA Lib "mpr.dll" ( _
    ByVal dwScope As Long, _
    ByVal dwType As Long, _
    ByVal dwUsage As Long, _
    ByVal lpNetResource As Long, _
    ByRef lphEnum As Long) As Long

Private Declare Function WNetOpenEnumW Lib "mpr.dll" ( _
    ByVal dwScope As Long, _
    ByVal dwType As Long, _
    ByVal dwUsage As Long, _
    ByVal lpNetResource As Long, _
    ByRef lphEnum As Long) As Long

Private Declare Function WNetEnumResourceA Lib "mpr.dll" ( _
    ByVal hEnum As Long, _
    ByRef lpcCount As Long, _
    ByVal lpBuffer As Long, _
    ByRef lpBufferSize As Long) As Long

Private Declare Function WNetEnumResourceW Lib "mpr.dll" ( _
    ByVal hEnum As Long, _
    ByRef lpcCount As Long, _
    ByVal lpBuffer As Long, _
    ByRef lpBufferSize As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" ( _
    ByVal hEnum As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private mTerminate As Boolean
Private MinWidth As Single
Private MinHeight As Single
Private VertGap As Single
Private ShowFilesLeft As Single
Private BrowseAbort As Boolean

Private mEnabled As Boolean
Private mVisible As Boolean
Private mBackColor As OLE_COLOR
Private mShowNetwork As Boolean
Private mShowLocal As Boolean
Private mShowFiles As Boolean
Private mIsSpecialFolder As Boolean
Private mSpecialFolderDesc As String
Private mFolder As String
Private mFile As String
Private mShowFilesVisible As Boolean
Private mShowNetworkVisible As Boolean
Private mNewFolderVisible As Boolean
Private mDetailsVisible As Boolean
Private mFileFilter As String
Private mFolderFilter As String
Private mShowHidden As Boolean
Private mShowSystem As Boolean
Private mShowReadOnly As Boolean
Private mTip As String
Private mWidth As Single
Private mHeight As Single
Private mBorder As Boolean
Private mCaption As String
Private mAppearance As AppearanceConstants
Private mBackStyle As Long
Private mSpeedy As Boolean

Private InitDone As Boolean
Private NOEW_OK As Boolean
Private NERW_OK As Boolean
Private UNICODE_OK As Boolean
Private PCPXW_OK As Boolean
Private PCPXA_OK As Boolean
Private PCPW_OK As Boolean
Private PCPA_OK As Boolean
Private KeyCnt As Long

Private LocalRoot As MSComctlLib.Node
Private SpecialRoot As MSComctlLib.Node

Private FileCnts As New Collection

Public Event FolderSelection(Folder As String)
Attribute FolderSelection.VB_Description = "Fired when a folder is selected."
Attribute FolderSelection.VB_MemberFlags = "200"
Public Event FileSelection(File As String)
Attribute FileSelection.VB_Description = "Fired whenever a file is selected."
Public Property Get Serialize(Optional SerialFormat As SERIAL_FORMAT = SERIAL_FORMAT_TSV) As String
    Dim sa(POS_UBOUND) As String

    sa(POS_SVERSION) = CStr(0)
    sa(POS_Enabled) = mEnabled
    sa(POS_Visible) = mVisible
    sa(POS_ShowNetwork) = mShowNetwork
    sa(POS_ShowLocal) = mShowLocal
    sa(POS_ShowFiles) = mShowFiles
    sa(POS_Folder) = mFolder
    sa(POS_File) = mFile
    sa(POS_ShowFilesVisible) = mShowFilesVisible
    sa(POS_ShowNetworkVisible) = mShowNetworkVisible
    sa(POS_NewFolderVisible) = mNewFolderVisible
    sa(POS_FileFilter) = mFileFilter
    sa(POS_FolderFilter) = mFolderFilter
    sa(POS_ShowHidden) = mShowHidden
    sa(POS_ShowSystem) = mShowSystem
    sa(POS_ShowReadOnly) = mShowReadOnly

    sa(POS_Width) = UserControl.Width
    sa(POS_Height) = UserControl.Height
    sa(POS_Border) = mBorder
    sa(POS_Caption) = lblBrowse.Caption
    sa(POS_Appearance) = UserControl.Appearance
    sa(POS_BackStyle) = UserControl.BackStyle

    sa(POS_FontName) = UserControl.FontName
    sa(POS_FontSize) = UserControl.FontSize
    sa(POS_FontBold) = UserControl.FontBold
    sa(POS_FontItalic) = UserControl.FontItalic
    sa(POS_FontUnderline) = UserControl.FontUnderline
    sa(POS_FontStrikethru) = UserControl.FontStrikethru
    sa(POS_ForeColor) = UserControl.ForeColor

    Serialize = Join(sa, vbTab)
End Property
Public Property Let Serialize(Optional SerialFormat As SERIAL_FORMAT = SERIAL_FORMAT_TSV, SerialData As String)
    Dim sa() As String
    Dim fa() As String
    Dim FieldCnt As Long
    Dim Fields As Collection
    Dim i As Long

    If Len(SerialData) <= 0 Then
        Err.Raise SERIALIZE_EMPTY
        Exit Property
    End If

    sa = Split(SerialData, vbTab)
    If UBound(sa) < POS_UBOUND Then
        Err.Raise SERIALIZE_INVALID
        Exit Property
    End If
    If sa(POS_SVERSION) <> CStr(0) Then
        Err.Raise SERIALIZE_BAD_VERSION
        Exit Property
    End If

    Enabled = sa(POS_Enabled)
    Visible = sa(POS_Visible)
    ShowNetwork = sa(POS_ShowNetwork)
    ShowLocal = sa(POS_ShowLocal)
    ShowFiles = sa(POS_ShowFiles)
    Folder = sa(POS_Folder)
'    File = sa(POS_File)
    ShowFilesVisible = sa(POS_ShowFilesVisible)
    ShowNetworkVisible = sa(POS_ShowNetworkVisible)
    NewFolderVisible = sa(POS_NewFolderVisible)
    FileFilter = sa(POS_FileFilter)
    FolderFilter = sa(POS_FolderFilter)
    ShowHidden = sa(POS_ShowHidden)
    ShowSystem = sa(POS_ShowSystem)
    ShowReadOnly = sa(POS_ShowReadOnly)

    Width = sa(POS_Width)
    Height = sa(POS_Height)
    Border = sa(POS_Border)
    Caption = sa(POS_Caption)
    Appearance = sa(POS_Appearance)
'    BackStyle = sa(POS_BackStyle)

    FontName = sa(POS_FontName)
    FontSize = sa(POS_FontSize)
'    FontBold = sa(POS_FontBold)
'    FontItalic = sa(POS_FontItalic)
'    FontUnderline = sa(POS_FontUnderline)
'    FontStrikethru = sa(POS_FontStrikethru)
'    ForeColor = sa(POS_ForeColor)
End Property
Private Sub BrowseInit()
    NOEW_OK = APIExists("mpr.dll", "WNetOpenEnumW")
    NERW_OK = APIExists("mpr.dll", "WNetEnumResourceW")
    UNICODE_OK = NOEW_OK And NERW_OK

    PCPXW_OK = APIExists("shlwapi.dll", "PathCompactPathExW")
    PCPXA_OK = APIExists("shlwapi.dll", "PathCompactPathExA")
    PCPW_OK = APIExists("shlwapi.dll", "PathCompactPathW")
    PCPA_OK = APIExists("shlwapi.dll", "PathCompactPathA")

    InitDone = True
End Sub
Private Function EnumNetContainer(Optional Spec As String = "", Optional Path As String = "") As PT_NET_INFO
    Dim nr As NETRESOURCE
    Dim ni As PT_NET_INFO
    Dim buflen As Long
    Dim result As Long
    Dim hEnum As Long
    Dim ResCnt As Long
    Dim lpbuff As Long
    Dim p As Long
    Dim i As Long
    Dim ParentNr As Long
    Dim Encoding As PT_TEXT_ENCODING
    Dim Provider As String
    Dim Container As String
    Dim DispType As Long
    Dim Usage As Long
    Dim sa() As String
    Dim PrevPath As String

' spec = provider & vbtab & container & vbtab & displaytype & vbtab & usage
    If Not InitDone Then BrowseInit
    PrevPath = Path

    ni.Cnt = 0
    EnumNetContainer = ni

    If UNICODE_OK Then
        Encoding = PT_TEXT_UNICODE
    Else
        Encoding = PT_TEXT_ASCII
    End If

    If Len(Spec) > 0 Then
        sa = Split(Spec, vbTab)
        Provider = sa(0)
        Container = sa(1)
        DispType = CLng(sa(2))
        Usage = CLng(sa(3))
        buflen = NETRES_BUFLEN
        ParentNr = GlobalAlloc(GPTR, buflen)
        If ParentNr = 0 Then Exit Function

        ApiMemoryZero ParentNr, buflen
        nr.dwDisplayType = DispType ' RESOURCEDISPLAYTYPE_NETWORK ' RESOURCEDISPLAYTYPE_DOMAIN
        nr.dwUsage = Usage ' RESOURCEUSAGE_CONTAINER Or RESOURCEUSAGE_RESERVED
        nr.dwScope = RESOURCE_GLOBALNET
        nr.dwType = RESOURCETYPE_ANY
        nr.lpLocalName = 0
        nr.lpProvider = 0

        p = ParentNr + LenB(nr)

        If UNICODE_OK Then
            nr.lpRemoteName = p
            p = p + ApiTextWrite(p, Container, PT_TEXT_UNICODE)
            nr.lpProvider = p
            p = p + ApiTextWrite(p, Provider, PT_TEXT_UNICODE)
        Else
            nr.lpRemoteName = p
            p = p + ApiTextWrite(p, Container, PT_TEXT_ASCII)
            nr.lpProvider = p
            p = p + ApiTextWrite(p, Provider, PT_TEXT_ASCII)
        End If
        ApiMemoryCopy ParentNr, VarPtr(nr), LenB(nr)
        If UNICODE_OK Then
            result = WNetOpenEnumW(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                RESOURCEUSAGE_ALL, ParentNr, hEnum)
        Else
            result = WNetOpenEnumA(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                RESOURCEUSAGE_ALL, ParentNr, hEnum)
        End If
        GlobalFree ParentNr
    Else    ' get root of network
        If UNICODE_OK Then
            result = WNetOpenEnumW(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                RESOURCEUSAGE_ALL, 0, hEnum)
        Else
            result = WNetOpenEnumA(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, _
                RESOURCEUSAGE_ALL, 0, hEnum)
        End If
    End If
    If result <> NO_ERROR Then Exit Function

    buflen = NETRES_BUFLEN
    lpbuff = GlobalAlloc(GPTR, buflen)
    If lpbuff = 0 Then Exit Function

    ApiMemoryZero lpbuff, buflen
    ResCnt = -1 ' ask for all
    If UNICODE_OK Then
        result = WNetEnumResourceW(hEnum, ResCnt, lpbuff, buflen)
    Else
        result = WNetEnumResourceA(hEnum, ResCnt, lpbuff, buflen)
    End If
    If result <> NO_ERROR Then ' buffer too small?
        If lpbuff <> 0 Then GlobalFree lpbuff
        If buflen <= 0 Then Exit Function
        lpbuff = GlobalAlloc(GPTR, buflen)
        If lpbuff = 0 Then Exit Function ' can't get buffer
    End If

    If result = NO_ERROR Then
        p = lpbuff
        For i = 1 To ResCnt
            DoEvents
            If BrowseAbort Then
                If lpbuff <> 0 Then GlobalFree lpbuff
                WNetCloseEnum hEnum
                Exit Function
            End If
            ApiMemoryCopy VarPtr(nr), p, LenB(nr)

            p = p + LenB(nr)
            If AddItem(nr) Then
                ReDim Preserve ni.Resource(ni.Cnt)
                ni.Resource(ni.Cnt).DispType = nr.dwDisplayType
                ni.Resource(ni.Cnt).Type = nr.dwType
                ni.Resource(ni.Cnt).Usage = nr.dwUsage
                ni.Resource(ni.Cnt).RemoteName = ApiTextCopy(nr.lpRemoteName, Encoding)
                ni.Resource(ni.Cnt).FullPath = PrevPath & "\" & ni.Resource(ni.Cnt).RemoteName
                ni.Resource(ni.Cnt).LocalName = ApiTextCopy(nr.lpLocalName, Encoding)
                ni.Resource(ni.Cnt).Provider = ApiTextCopy(nr.lpProvider, Encoding)
                If (nr.dwUsage And RESOURCEUSAGE_CONTAINER) Then
                    ni.Resource(ni.Cnt).IsContainer = True
'                Else
'                    If InStr(1, ni.Resource(ni.Cnt).RemoteName, "epson", vbTextCompare) > 0 Then Stop
                End If
                ni.Resource(ni.Cnt).Spec = ni.Resource(ni.Cnt).Provider & vbTab & ni.Resource(ni.Cnt).RemoteName & vbTab & nr.dwDisplayType & vbTab & nr.dwUsage

                'If Not (BrowseForm Is Nothing) Then BrowseForm.CallBack ni.Resource(ni.Cnt).LocalName
                DoEvents
                ni.Cnt = ni.Cnt + 1
            End If
        Next i
    End If

    If lpbuff <> 0 Then GlobalFree lpbuff
    WNetCloseEnum hEnum
    EnumNetContainer = ni
End Function
Private Function AddItem(nr As NETRESOURCE) As Boolean
    AddItem = True

    If nr.dwUsage And RESOURCEUSAGE_CONTAINER Then Exit Function
    If nr.dwType And RESOURCETYPE_DISK Then Exit Function

    AddItem = False
End Function
Private Sub DumpHex(Title As String, Addr As Long, Length As Long)
    Dim Ptr As Long
    Dim b As Byte
    Dim i As Long
    Dim s As String
    Dim Offset As Long

    If Addr = 0 Then
        Debug.Print Title
        Exit Sub
    End If

    Offset = Addr Mod 32
    s = Title
    Ptr = Addr
    For i = 0 To Length - 1
        If ((Addr + i) Mod 32) = Offset Then
            Debug.Print s
            Debug.Print Format(i, "00000") & ": ";
            s = ""
        End If
        ApiMemoryCopy VarPtr(b), Ptr, 1
        If b = 0 Then
            Debug.Print ".";
        ElseIf (b < 32) Or (b > 127) Then
            Debug.Print "·";
        Else
            Debug.Print Chr(b);
        End If
        Ptr = Ptr + 1
    Next i
    Debug.Print
End Sub
Private Sub btnNewFolder_Click()
    NewFolder
End Sub
Private Sub NewFolder()
    Dim FolderName As String
    Dim n As MSComctlLib.Node

    FolderName = InputBox("Name of new folder:", "Create new folder in " & mFolder, "New Folder")
    FolderName = Trim(FolderName)
    If Len(FolderName) > 0 Then
        MkDir mFolder & "\" & FolderName
        Set n = AddChild(tvBrowse.SelectedItem, KEY_FOLDER, FolderName, IMAGE_FOLDER, TAG_FOLDER)
        AddDummy n
    End If
End Sub
Private Sub chkShowNetwork_Click()
    If chkShowNetwork.Value = vbChecked Then
        If Not mShowNetwork Then ShowNetwork = True
    Else
        If mShowNetwork Then ShowNetwork = False
    End If
    If Not (tvBrowse.SelectedItem Is Nothing) Then tvBrowse.SelectedItem.EnsureVisible
End Sub
Private Sub fraDetails_DblClick()
    DispFolderInfo tvBrowse.SelectedItem, True
End Sub
Private Sub mnuCopyFile_Click()
    Clipboard.Clear
    Clipboard.SetText ExtractPath(tvBrowse.SelectedItem)
End Sub
Private Sub mnuCopyFolder_Click()
    Dim s As String

    Clipboard.Clear
    s = ExtractPath(tvBrowse.SelectedItem)
    If Right(s, 1) <> "\" Then s = s & "\"
    Clipboard.SetText s
End Sub
Private Sub mnuNewFolder_Click()
    NewFolder
End Sub
Private Sub tvBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Node As MSComctlLib.Node
    Dim c As String

    If (Button And vbRightButton) = 0 Then Exit Sub

    Set Node = tvBrowse.HitTest(X, Y)
    If Not (Node Is Nothing) Then MakeSelection Node

    If btnNewFolder.Visible Then
        mnuNewFolder.Visible = True
        If btnNewFolder.Enabled Then
            If Not (Node Is Nothing) Then
                mnuNewFolder.Enabled = True
            Else
                mnuNewFolder.Enabled = False
            End If
        Else
            mnuNewFolder.Enabled = False
        End If
    Else
        mnuNewFolder.Visible = False
    End If

    If Node Is Nothing Then
        mnuCopyFolder.Enabled = False
        mnuCopyFile.Enabled = False
    Else
        c = ExtractKeyCode(Node)
        Select Case c
            Case KEY_FILE:
                mnuCopyFile.Enabled = True
                mnuCopyFolder.Enabled = False
            Case KEY_FOLDER, KEY_FOLDER_HIDDEN, KEY_SPECIAL:
                mnuCopyFolder.Enabled = True
                mnuCopyFile.Enabled = False
            Case Else
                mnuCopyFolder.Enabled = False
                mnuCopyFile.Enabled = False
        End Select
    End If

    PopupMenu mnuPopup, , , , mnuCancel
End Sub
Private Sub txtDetails_DblClick()
    DispFolderInfo tvBrowse.SelectedItem, True
End Sub
Private Sub UserControl_Initialize()
    Dim n As MSComctlLib.Node

    mSpeedy = True
    KEY_FOLDER = KEY_FOLDER_DEF
    picBorder.BackColor = UserControl.BackColor
    fraDetails.BackColor = UserControl.BackColor
    txtDetails.BackColor = UserControl.BackColor
    chkShowFiles.BackColor = UserControl.BackColor
    chkShowNetwork.BackColor = UserControl.BackColor

    MinWidth = UserControl.Width
    MinHeight = UserControl.Height
    VertGap = fraDetails.top - (btnNewFolder.top + btnNewFolder.Height)
    ShowFilesLeft = chkShowFiles.left

    BrowseInit

    tvBrowse.Nodes.Clear
    txtDetails.Text = "Click Folder name for more info"
    fraDetails.Visible = mDetailsVisible
    KeyCnt = 0
End Sub
Private Sub UserControl_InitProperties()
    BackColor = UserControl.BackColor
'    picBorder.BackColor = UserControl.BackColor
'    fraDetails.BackColor = UserControl.BackColor
'    txtDetails.BackColor = UserControl.BackColor
'    chkShowFiles.BackColor = UserControl.BackColor
'    chkShowNetwork.BackColor = UserControl.BackColor
    tvBrowse.ToolTipText = Extender.ToolTipText

    ShowHidden = True
    ShowSystem = True
    ShowReadOnly = True

    FileFilter = ""
    FolderFilter = ""

    mSpeedy = True

    ShowLocal = True
    If chkShowNetwork.Value = vbChecked Then ShowNetwork = True
    If chkShowFiles.Value = vbChecked Then ShowFiles = True

    mShowNetworkVisible = chkShowNetwork.Visible
    mShowFilesVisible = chkShowFiles.Visible
    mNewFolderVisible = btnNewFolder.Visible
    mDetailsVisible = fraDetails.Visible

    txtDetails.Text = "Click Folder name for more info"
    Enabled = True
    Visible = True
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", lblBrowse.Caption
    PropBag.WriteProperty "Speedy", mSpeedy
    PropBag.WriteProperty "ShowHidden", mShowHidden
    PropBag.WriteProperty "ShowSystem", mShowSystem
    PropBag.WriteProperty "ShowReadOnly", mShowReadOnly
    PropBag.WriteProperty "ShowLocal", mShowLocal
    PropBag.WriteProperty "ShowNetwork", mShowNetwork
    PropBag.WriteProperty "ShowNetworkVisible", mShowNetworkVisible
    PropBag.WriteProperty "ShowFilesVisible", mShowFilesVisible
    PropBag.WriteProperty "NewFolderVisible", mNewFolderVisible
    PropBag.WriteProperty "DetailsVisible", mDetailsVisible
    PropBag.WriteProperty "Height", UserControl.Height
    PropBag.WriteProperty "Width", UserControl.Width
'    PropBag.WriteProperty "BackStyle", 1
    PropBag.WriteProperty "Border", Border
'    Call PropBag.WriteProperty("BackColor", mbackcolor, SystemColorConstants.vbButtonFace)
    PropBag.WriteProperty "BackColor", mBackColor ' SystemColorConstants.vbButtonFace
    PropBag.WriteProperty "Appearance", UserControl.Appearance
    PropBag.WriteProperty "Visible", mVisible
    PropBag.WriteProperty "Enabled", mEnabled
    PropBag.WriteProperty "ShowFiles", ShowFiles
    PropBag.WriteProperty "FileFilter", mFileFilter
    PropBag.WriteProperty "FolderFilter", mFolderFilter
    PropBag.WriteProperty "Font", Font
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Caption = PropBag.ReadProperty("Caption")
    mSpeedy = PropBag.ReadProperty("Speedy")
    mShowHidden = PropBag.ReadProperty("ShowHidden")
    mShowSystem = PropBag.ReadProperty("ShowSystem")
    mShowReadOnly = PropBag.ReadProperty("ShowReadOnly")
    ShowLocal = PropBag.ReadProperty("ShowLocal")
    ShowNetwork = PropBag.ReadProperty("ShowNetwork")
    ShowNetworkVisible = PropBag.ReadProperty("ShowNetworkVisible")
    ShowFilesVisible = PropBag.ReadProperty("ShowFilesVisible")
    NewFolderVisible = PropBag.ReadProperty("NewFolderVisible")
    DetailsVisible = PropBag.ReadProperty("DetailsVisible")
    Height = PropBag.ReadProperty("Height")
    Width = PropBag.ReadProperty("Width")
'    UserControl.BackStyle = 1
    Border = PropBag.ReadProperty("Border")
'    chkShowFiles.BackColor = UserControl.Extender.BackColor
''    chkShowNetwork.BackColor = UserControl.Extender.BackColor
    lblBrowse.BackColor = UserControl.BackColor
    tvBrowse.ToolTipText = UserControl.Extender.ToolTipText
'    UserControl.BackColor = &H8000000F 'PropBag.ReadProperty("BackColor", &H8000000F)
    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)   'mBackColor ' SystemColorConstants.vbButtonFace ' PropBag.ReadProperty("BackColor")
    Appearance = PropBag.ReadProperty("Appearance")
    Enabled = PropBag.ReadProperty("Enabled")
    Visible = PropBag.ReadProperty("Visible")
    Enabled = PropBag.ReadProperty("Enabled")
    ShowFiles = PropBag.ReadProperty("ShowFiles")
    FileFilter = PropBag.ReadProperty("FileFilter")
    FolderFilter = PropBag.ReadProperty("FolderFilter")
    Set Font = PropBag.ReadProperty("Font")
End Sub
Private Sub UserControl_Resize()
    If UserControl.Width < MinWidth Then
        UserControl.Width = MinWidth
        PropertyChanged "Width"
        Exit Sub
    End If

    If UserControl.Height < MinHeight Then
        UserControl.Height = MinHeight
        PropertyChanged "Height"
        Exit Sub
    End If

    ControlResize
End Sub
Private Sub ControlResize()
    Dim ControlWidth As Single
    Dim ControlHeight As Single
    Dim LeftOver As Single

'    ControlWidth = UserControl.Width / Screen.TwipsPerPixelX
'    ControlHeight = UserControl.Height / Screen.TwipsPerPixelY
    ControlWidth = UserControl.Width
    ControlHeight = UserControl.Height

    picBorder.Width = ControlWidth
    picBorder.Height = ControlHeight

    lblBrowse.Width = ControlWidth
    tvBrowse.Width = ControlWidth
    fraDetails.Width = ControlWidth
'    txtDetails.Width = (ControlWidth * Screen.TwipsPerPixelX) - 240
    txtDetails.Width = ControlWidth - 240

    btnNewFolder.left = ControlWidth - btnNewFolder.Width
    If mShowFilesVisible And (Not mShowNetworkVisible) Then
        chkShowFiles.left = chkShowNetwork.left
    Else
        chkShowFiles.left = ShowFilesLeft
    End If

    LeftOver = ControlHeight
    If mDetailsVisible Then
        fraDetails.top = LeftOver - fraDetails.Height
        LeftOver = fraDetails.top - VertGap
    End If

    If mNewFolderVisible Then
        btnNewFolder.top = LeftOver - btnNewFolder.Height
        chkShowFiles.top = btnNewFolder.top + ((btnNewFolder.Height - chkShowFiles.top) / 2)
        chkShowNetwork.top = chkShowFiles.top
        LeftOver = btnNewFolder.top - VertGap
    ElseIf mShowNetworkVisible Or mShowFilesVisible Then
'        LeftOver = LeftOver - VertGap
        chkShowFiles.top = LeftOver - chkShowFiles.Height
        chkShowNetwork.top = chkShowFiles.top
        LeftOver = chkShowFiles.top - VertGap
'        LeftOver = LeftOver - VertGap
    End If
    tvBrowse.Height = LeftOver - tvBrowse.top
End Sub
Private Sub ExpandLocal(BaseNode As MSComctlLib.Node)
    Dim s As String
    Dim drv() As String
    Dim i As Long
    Dim n As MSComctlLib.Node
    Dim Desc As String
    Dim result As Long

    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove (BaseNode.Child.Index)
    Wend
    DoEvents

    Desc = "Special Folders"
    Set n = AddChild(BaseNode, KEY_SPECIALROOT, Desc, IMAGE_FOLDER, TAG_SPECIAL)
    Set SpecialRoot = n
    AddDummy n

    s = String(4000, " ")
    result = GetLogicalDriveStrings(Len(s), s)
    s = Trim(s)
    drv = Split(s, Chr(0))

    For i = 0 To UBound(drv)
        If Len(drv(i)) > 0 Then
            Desc = left(drv(i), 1) & ":"
            Set n = AddChild(BaseNode, KEY_FOLDER, Desc, IMAGE_DRIVE, TAG_LOCALROOT)
            AddDummy n
        End If
    Next i
End Sub
Private Sub tvBrowse_NodeClick(ByVal Node As MSComctlLib.Node)
    Static PrevNode As MSComctlLib.Node

    MakeSelection Node
    If PrevNode Is Node Then DispFolderInfo Node, True
    Set PrevNode = Node
End Sub
Private Sub MakeSelection(Node As MSComctlLib.Node)
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim fi As PT_FILE_INFO
    Dim TempFolder As String
    Dim TempFile As String

    Set tvBrowse.SelectedItem = Node

    If Node Is Nothing Then
        txtDetails.ToolTipText = ""
        mFolder = ""
        If mFolder <> "" Then
            mFolder = ""
            RaiseEvent FolderSelection(mFolder)
        End If
        If mFile <> "" Then
            mFile = ""
            RaiseEvent FileSelection(mFile)
        End If
        btnNewFolder.Enabled = False
        Exit Sub
    End If

    c = ExtractKeyCode(Node)
    mIsSpecialFolder = False
    mSpecialFolderDesc = ""
    Select Case c
        Case KEY_FILE:
            btnNewFolder.Enabled = False
            s = ExtractPath(Node)
            TempFile = s
            i = InStrRev(s, "\")
            s = left(s, i - 1)
            If Right(s, 1) <> "\" Then s = s & "\"
            TempFolder = s
'            fi = ApiFileInfo(TempFile)
            s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, "File: " & TempFile)
'            If Not fi.Error Then
'                s = s & vbCrLf & "Size: " & FormatNumber(fi.Size, 0) & " bytes"
'                If fi.Creation.Time64 <> 0 Then
'                    s = s & vbCrLf & "Created: " & FormatDateTime(fi.Creation.Time, vbGeneralDate)
'                End If
'                If fi.LastWrite.Time64 <> 0 Then
'                    s = s & vbCrLf & "Modified: " & FormatDateTime(fi.LastWrite.Time, vbGeneralDate)
'                End If
'                If fi.LastAccess.Time64 <> 0 Then
'                    s = s & vbCrLf & "Accessed: " & FormatDateTime(fi.LastAccess.Time, vbGeneralDate)
'                End If
'            End If
            txtDetails.Text = s
            If StrComp(mFolder, TempFolder, vbTextCompare) <> 0 Then
                mFolder = TempFolder
                RaiseEvent FolderSelection(mFolder)
            Else
                mFolder = TempFolder ' in case case has changed
            End If
            If StrComp(mFile, TempFile, vbTextCompare) <> 0 Then
                mFile = TempFile
                RaiseEvent FileSelection(mFile)
            Else
                mFile = TempFile
            End If

        Case KEY_FOLDER, KEY_FOLDER_HIDDEN:
            s = ExtractPath(Node)
            If Right(s, 1) <> "\" Then s = s & "\"
            TempFolder = s
            If mShowFiles Then
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, "Folder: " & s)
            Else
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, s)
            End If
            txtDetails.Text = "Working ..."
            DoEvents
            txtDetails.Text = s & FileCnt(Node)
            If StrComp(mFolder, TempFolder, vbTextCompare) <> 0 Then
                mFolder = TempFolder
                RaiseEvent FolderSelection(mFolder)
            Else
                mFolder = TempFolder
            End If
            If mFile <> "" Then
                mFile = ""
                RaiseEvent FileSelection(mFile)
            End If
            btnNewFolder.Enabled = True
            btnNewFolder.ToolTipText = "Create a new folder in " & mFolder

        Case KEY_SPECIAL:
            mIsSpecialFolder = True
            mSpecialFolderDesc = Node.Text
            s = ExtractPath(Node)
            If Right(s, 1) <> "\" Then s = s & "\"
            TempFolder = s
            If mShowFiles Then
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, "Folder: " & s)
            Else
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, s)
            End If
            txtDetails.Text = "Working ..."
            DoEvents
            txtDetails.Text = s & FileCnt(Node)
            If StrComp(mFolder, TempFolder, vbTextCompare) <> 0 Then
                mFolder = TempFolder
                RaiseEvent FolderSelection(mFolder)
            Else
                mFolder = TempFolder
            End If
            If mFile <> "" Then
                mFile = ""
                RaiseEvent FileSelection(mFile)
            End If
            btnNewFolder.Enabled = True
            btnNewFolder.ToolTipText = "Create a new folder in " & mFolder

        Case Else:
            txtDetails.Text = ""
            btnNewFolder.Enabled = False
            If mFolder <> "" Then
                mFolder = ""
                RaiseEvent FolderSelection(mFolder)
            End If
            If mFile <> "" Then
                mFile = ""
                RaiseEvent FileSelection(mFile)
            End If

    End Select

    txtDetails.ToolTipText = Replace(SEL_TIP, "%SEL%", mFolder, , , vbTextCompare)
End Sub
Private Sub DispFolderInfo(Node As MSComctlLib.Node, Optional Force As Boolean = False)
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim fi As PT_FILE_INFO
    Dim TempFolder As String
    Dim TempFile As String

    If Node Is Nothing Then
        txtDetails.Text = ""
        txtDetails.ToolTipText = ""
        Exit Sub
    End If

    tvBrowse.Enabled = False
    txtDetails.Enabled = False
    fraDetails.Enabled = False
    DoEvents

    c = ExtractKeyCode(Node)
    Select Case c
        Case KEY_FILE:
            s = ExtractPath(Node)
            TempFile = s
            i = InStrRev(s, "\")
            s = left(s, i - 1)
            If Right(s, 1) <> "\" Then s = s & "\"
            TempFolder = s
            fi = ApiFileInfo(TempFile)
            s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, "File: " & TempFile)
            If Not fi.Error Then
                s = s & vbCrLf & "Size: " & FormatNumber(fi.Size, 0) & " bytes"
                If fi.Creation.Time64 <> 0 Then
                    s = s & vbCrLf & "Created: " & FormatDateTime(fi.Creation.TimeLocal, vbGeneralDate)
                End If
                If fi.LastWrite.Time64 <> 0 Then
                    s = s & vbCrLf & "Modified: " & FormatDateTime(fi.LastWrite.TimeLocal, vbGeneralDate)
                End If
                If fi.LastAccess.Time64 <> 0 Then
                    s = s & vbCrLf & "Accessed: " & FormatDateTime(fi.LastAccess.TimeLocal, vbGeneralDate)
                End If
            End If
            txtDetails.Text = s

        Case KEY_FOLDER, KEY_FOLDER_HIDDEN, KEY_SPECIAL:
            s = ExtractPath(Node)
            If Right(s, 1) <> "\" Then s = s & "\"
            TempFolder = s
            If mShowFiles Then
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, "Folder: " & s)
            Else
                s = ApiTextPathFit(UserControl.hdc, txtDetails.Width, s)
            End If
            txtDetails.Text = "Working ..."
            txtDetails.Text = s & FileCnt(Node, Force)
        Case Else:

    End Select

    txtDetails.Enabled = True
    fraDetails.Enabled = True
    tvBrowse.Enabled = True
    On Error Resume Next
    If tvBrowse.Visible Then tvBrowse.SetFocus
    Err.Clear
End Sub
Private Function ExtractKeyCode(ByRef BaseNode As MSComctlLib.Node) As String
    Dim c As String

    c = left(BaseNode.Key, 1)
    ExtractKeyCode = c
End Function
Private Function ExtractPath(ByRef BaseNode As MSComctlLib.Node) As String
    Dim Prefix As String
    Dim i As Long
    Dim FullPath As String
    Dim n As MSComctlLib.Node
    Dim c As String

    c = ExtractKeyCode(BaseNode)
    If c = KEY_SPECIAL Then
        i = InStr(1, BaseNode.Tag, "-")
        ExtractPath = Mid(BaseNode.Tag, i + 1)
        Exit Function
    End If

    FullPath = BaseNode.Text
    Set n = BaseNode.Parent
    Do While Not (n Is Nothing)
        c = ExtractKeyCode(n)
        Select Case c
            Case KEY_LOCALROOT:
                ExtractPath = CleanupPath(BaseNode.FullPath)
                Exit Function
            Case KEY_NETROOT:
                ExtractPath = CleanupPath(BaseNode.FullPath)
                Exit Function
            Case KEY_SPECIALROOT:
                ExtractPath = FullPath
                Exit Function
            Case KEY_FOLDER, KEY_FOLDER_HIDDEN:
                FullPath = n.Text & FullPath
            Case KEY_SPECIAL:
                i = InStr(1, n.Tag, "-")
                If i > 1 Then FullPath = Mid(n.Tag, i + 1) & FullPath
            Case KEY_CONTAINER:
                ExtractPath = CleanupPath(BaseNode.FullPath)
                Exit Function
        End Select
        Set n = n.Parent
    Loop

    ExtractPath = FullPath
End Function
Private Function CleanupPath(FullPath As String) As String
    Dim i As Long
    Dim s As String
    Dim ep As String

    s = FullPath
    i = InStrRev(s, "\\\")
    If i > 0 Then ' share
        ep = Mid(s, i + 1)
    ElseIf InStr(1, s, ":") <= 0 Then ' not local either
        ep = ""
    Else
        i = InStr(1, s, "\")
        ep = Mid(s, i + 1)
    End If
    
    CleanupPath = ep
End Function
Private Sub tvBrowse_Expand(ByVal Node As MSComctlLib.Node)
    Dim c As String

    tvBrowse.Enabled = False
    DoEvents

    c = ExtractKeyCode(Node)
    Select Case c
        Case KEY_LOCALROOT:
            ExpandLocal Node
        Case KEY_NETROOT:
            ExpandNetRoot Node
        Case KEY_SPECIALROOT:
            ExpandSpecialRoot Node
        Case KEY_FOLDER, KEY_FOLDER_HIDDEN:
            ExpandFolder Node
        Case KEY_SPECIAL:
            ExpandSpecial Node
        Case KEY_CONTAINER:
            ExpandContainer Node
    End Select

'    MakeSelection Node
    tvBrowse.Enabled = True
    On Error Resume Next
    If tvBrowse.Visible Then tvBrowse.SetFocus
    Err.Clear
End Sub
Private Sub ExpandNetRoot(BaseNode As MSComctlLib.Node)
    Dim ni As PT_NET_INFO
    Dim i As Long
    Dim s As String
    Dim Desc As String
    Dim n As MSComctlLib.Node
    Dim Key As String

    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove (BaseNode.Child.Index)
    Wend
    DoEvents

    ni = EnumNetContainer()
    If ni.Cnt > 0 Then
        For i = 0 To UBound(ni.Resource)
            s = ni.Resource(i).RemoteName
            If Len(s) > 0 Then
                Desc = ni.Resource(i).RemoteName
                If ni.Resource(i).IsContainer Then
                    If StrComp(s, MWN_NAME, vbTextCompare) = 0 Then
                        Key = KEY_CONTAINER_MWN
                        MWN_PRESENT = True
                    Else
                        Key = KEY_CONTAINER
                    End If
                Else
                    Key = KEY_FOLDER
                End If
                Set n = AddChild(BaseNode, Key, Desc, IMAGE_PROVIDER)
                n.Tag = ni.Resource(i).Spec
                AddDummy n
            End If
        Next i
    End If
End Sub
Private Sub ExpandContainer(BaseNode As MSComctlLib.Node)
    Dim ni As PT_NET_INFO
    Dim nix As Long
    Dim s As String
    Dim Desc As String
    Dim n As MSComctlLib.Node
    Dim Key As String
    Dim i As Long
    Dim Img As String
    Dim Attrs As PT_ATTR

    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove (BaseNode.Child.Index)
    Wend
    DoEvents

    ni = EnumNetContainer(BaseNode.Tag, BaseNode.FullPath)
    If ni.Cnt > 0 Then
        For nix = 0 To UBound(ni.Resource)
            s = ni.Resource(nix).RemoteName
            If Len(s) > 0 Then
                If ni.Resource(nix).IsContainer Then
                    Key = KEY_CONTAINER
                    Desc = ni.Resource(nix).RemoteName
                Else
                    Key = KEY_FOLDER
                    Desc = ni.Resource(nix).RemoteName
                    If left(Desc, 2) = "\\" Then
                        i = InStr(3, Desc, "\")
                        If i > 0 Then Desc = Mid(Desc, i + 1)
                    End If
                End If

                Select Case ni.Resource(nix).DispType
                    Case NET_DISP_TYPE_ROOT
                        Img = IMAGE_NETROOT
                    Case NET_DISP_TYPE_NETWORK:
                        Img = IMAGE_NETWORK
                    Case NET_DISP_TYPE_DOMAIN:
                        Img = IMAGE_DOMAIN
                    Case NET_DISP_TYPE_DIRECTORY:
                        If ApiAttrsGet(CleanupPath(ni.Resource(nix).FullPath)) And PT_ATTR_HIDDEN Then
                            Img = IMAGE_FOLDER_HIDDEN
                        Else
                            Img = IMAGE_FOLDER
                        End If
                    Case NET_DISP_TYPE_SHARE:
                        If ApiAttrsGet(CleanupPath(ni.Resource(nix).FullPath)) And PT_ATTR_HIDDEN Then
                            Img = IMAGE_FOLDER_HIDDEN_SHARED
                        Else
                            Img = IMAGE_FOLDER_SHARED
                        End If
                    Case NET_DISP_TYPE_FILE:
                        If ApiAttrsGet(CleanupPath(ni.Resource(nix).FullPath)) And PT_ATTR_HIDDEN Then
                            Img = IMAGE_FILE_HIDDEN
                        Else
                            Img = IMAGE_FILE
                        End If
                    Case Else:
                        Img = IMAGE_COMPUTER
                End Select

                Set n = AddChild(BaseNode, Key, Desc, Img)
                n.Tag = ni.Resource(nix).Spec
                AddDummy n
            End If
        Next nix
    End If
End Sub
Private Function FileCnt(Node As MSComctlLib.Node, Optional Force As Boolean = False) As String
    Dim fd As PT_FILE_INFO ' FIND_DATA
    Dim hnd As Long
    Dim result As Long
    Dim s As String
    Dim SizeFiles As Variant
    Dim CntFile As Long
    Dim CntDir As Long
    Dim i As Long
    Dim DoCnt As Long
    Dim FullPath As String

    FullPath = ExtractPath(Node)

    If Not Force Then
        On Error Resume Next
        s = FileCnts(FullPath)
        If (Err.Number = 0) And (Len(s) > 0) Then
            FileCnt = s
            Exit Function
        End If
        On Error GoTo 0

        If mSpeedy Then
            FileCnt = ""
            Exit Function
        End If
    End If

    s = FullPath
    If Right(s, 1) <> "\" Then s = s & "\"
'    s = s & "*.*"

    SizeFiles = CDec(0)
    hnd = ApiFindFirstFile(fd, s, "*.*")
    If (hnd <> 0) And (hnd <> INVALID_HANDLE_VALUE) Then
        Do
            If (fd.Attrs And PT_ATTR_DIRECTORY) <> 0 Then
                If FilterFolder(fd) Then CntDir = CntDir + 1
            Else
                If FilterFile(fd) Then
                    CntFile = CntFile + 1
                    SizeFiles = SizeFiles + fd.Size
                End If
            End If
            result = ApiFindNextFile(fd)
            If DoCnt >= 10 Then
                DoEvents
                DoCnt = 0
            Else
                DoCnt = DoCnt + 1
            End If
        Loop While result <> 0
        result = ApiFindClose(fd)
    End If

    s = vbCrLf & "Sub-Folders: " & FormatNumber(CntDir, 0) & vbCrLf & "Files: " & FormatNumber(CntFile, 0) & vbCrLf & "Total Size: " & FormatNumber(SizeFiles, 0) & " bytes"
    On Error Resume Next
    FileCnts.Remove FullPath
    FileCnts.Add s, FullPath
    FileCnt = s
End Function
Private Sub chkShowFiles_Click()
    Dim b As Boolean

    If chkShowFiles.Value = vbChecked Then
        b = True
    Else
        b = False
    End If
    If b = mShowFiles Then Exit Sub

    mShowFiles = b
    DispFiles
    If Not (tvBrowse.SelectedItem Is Nothing) Then tvBrowse.SelectedItem.EnsureVisible
End Sub
Private Sub DispFiles()
    Dim SelNode As MSComctlLib.Node
    Dim SelKey As String
    Dim i As Long
    Dim n As MSComctlLib.Node
    Dim lim As Long
    Dim ExpKey() As String
    Dim ExpCnt As Long
    Dim c As String

    Set SelNode = tvBrowse.SelectedItem
    If SelNode Is Nothing Then
        Exit Sub
    End If

    SelKey = SelNode.Key
    lim = tvBrowse.Nodes.Count

    If mShowFiles Then
        'If Not (SelNode Is Nothing) Then tvBrowse_Expand SelNode
        ' call tvBrowse_expand for all expanded nodes?
        ExpCnt = 0
        For i = 1 To lim
            Set n = tvBrowse.Nodes(i)
            c = ExtractKeyCode(n)
            If (c = KEY_FOLDER) Or (c = KEY_FOLDER_HIDDEN) Then
            'If StrComp(n.Tag, TAG_FOLDER, vbTextCompare) = 0 Then
                If n.Expanded Then
                    ReDim Preserve ExpKey(ExpCnt)
                    ExpKey(ExpCnt) = n.Key
                    ExpCnt = ExpCnt + 1
                End If
            End If
        Next i
        If ExpCnt > 0 Then
            lim = UBound(ExpKey)
            For i = lim To 0 Step -1
                AddFiles tvBrowse.Nodes(ExpKey(i))
            Next i
        End If
    Else
        For i = lim To 1 Step -1
            If i <= tvBrowse.Nodes.Count Then
                Set n = tvBrowse.Nodes(i)
                If StrComp(n.Tag, TAG_FILE, vbTextCompare) = 0 Then
                    If n.Key = SelKey Then
                        SelKey = n.Parent.Key
                    End If
                    tvBrowse.Nodes.Remove n.Index
                End If
            End If
        Next i
    End If

    On Error Resume Next
    tvBrowse.Nodes(SelKey).Selected = True
    If tvBrowse.Visible Then tvBrowse.SetFocus
    Err.Clear
    tvBrowse_NodeClick SelNode
End Sub
Private Sub ExpandSpecialRoot(BaseNode As MSComctlLib.Node)
    Dim result As Long
    Dim s As String
    Dim NewNode As MSComctlLib.Node
    Dim i As Long
    Dim CommonFolder As String
    Dim PersonalFolder As String

    tvBrowse.Enabled = False
    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove BaseNode.Child.Index
    Wend
    DoEvents

    PersonalFolder = ApiSpecialFolder(CSIDL_DESKTOPDIRECTORY)
    CommonFolder = ApiSpecialFolder(CSIDL_DESKTOP)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Desktop"

    PersonalFolder = ApiSpecialFolder(CSIDL_MYDOCUMENTS)
    CommonFolder = ApiSpecialFolder(CSIDL_PERSONAL)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Documents"

    PersonalFolder = ApiSpecialFolder(CSIDL_MYPICTURES)
    CommonFolder = ApiSpecialFolder(CSIDL_COMMON_PICTURES)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Pictures"

    PersonalFolder = ApiSpecialFolder(CSIDL_MYMUSIC)
    CommonFolder = ApiSpecialFolder(CSIDL_COMMON_MUSIC)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Music"

    PersonalFolder = ApiSpecialFolder(CSIDL_MYVIDEO)
    CommonFolder = ApiSpecialFolder(CSIDL_COMMON_VIDEO)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Video"

    PersonalFolder = ApiSpecialFolder(CSIDL_FAVORITES)
    CommonFolder = ApiSpecialFolder(CSIDL_COMMON_FAVORITES)
    AddSpecial BaseNode, PersonalFolder, CommonFolder, "Favorites"

    CommonFolder = ApiSpecialFolder(CSIDL_WINDOWS)
    If Len(CommonFolder) > 0 Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "Windows", IMAGE_FOLDER, TAG_SPECIAL & "-" & CommonFolder)
        AddDummy NewNode
    End If

    CommonFolder = ApiSpecialFolder(CSIDL_SYSTEM)
    If Len(CommonFolder) > 0 Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "System", IMAGE_FOLDER, TAG_SPECIAL & "-" & CommonFolder)
        AddDummy NewNode
    End If

    CommonFolder = ApiSpecialFolder(CSIDL_PROGRAM_FILES)
    If Len(CommonFolder) > 0 Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "Program Files", IMAGE_FOLDER, TAG_SPECIAL & "-" & CommonFolder)
        AddDummy NewNode
    End If

    CommonFolder = ApiSpecialFolder(CSIDL_FONTS)
    If Len(CommonFolder) > 0 Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "Fonts", IMAGE_FOLDER, TAG_SPECIAL & "-" & CommonFolder)
        AddDummy NewNode
    End If

    If chkShowFiles.Value = vbChecked Then AddFiles BaseNode
    tvBrowse.Enabled = True
    DoEvents
End Sub
Private Sub AddSpecial(BaseNode As MSComctlLib.Node, PersonalFolder As String, CommonFolder As String, FolderDesc As String)
    Dim NewNode As MSComctlLib.Node
    Dim PersonalTag As String
    Dim CommonTag As String

    PersonalTag = TAG_SPECIAL & "-" & PersonalFolder
    CommonTag = TAG_SPECIAL & "-" & CommonFolder

    If StrComp(PersonalFolder, CommonFolder, vbTextCompare) = 0 Then
        If Len(PersonalFolder) > 0 Then
            Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "My " & FolderDesc, IMAGE_FOLDER, PersonalTag)
            AddDummy NewNode
        End If
    ElseIf (Len(PersonalFolder) > 0) And (Len(CommonFolder) <= 0) Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "My " & FolderDesc, IMAGE_FOLDER, PersonalTag)
        AddDummy NewNode
    ElseIf (Len(PersonalFolder) <= 0) And (Len(CommonFolder) > 0) Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "My " & FolderDesc, IMAGE_FOLDER, CommonTag)
        AddDummy NewNode
    ElseIf (Len(PersonalFolder) > 0) And (Len(CommonFolder) > 0) Then
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "My " & FolderDesc, IMAGE_FOLDER, PersonalTag)
        AddDummy NewNode
        Set NewNode = AddChild(BaseNode, KEY_SPECIAL, "Shared " & FolderDesc, IMAGE_FOLDER, CommonTag)
        AddDummy NewNode
    End If

End Sub
Private Sub ExpandSpecial(BaseNode As MSComctlLib.Node)
    Dim result As Long
    Dim s As String
    Dim NewNode As MSComctlLib.Node
    Dim i As Long

    tvBrowse.Enabled = False
    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove BaseNode.Child.Index
    Wend
    DoEvents

    AddFolders BaseNode
    If chkShowFiles.Value = vbChecked Then AddFiles BaseNode
    tvBrowse.Enabled = True
    DoEvents
End Sub
Private Sub ExpandFolder(BaseNode As MSComctlLib.Node)
    Dim result As Long
    Dim s As String
    Dim NewNode As MSComctlLib.Node
    Dim i As Long

    tvBrowse.Enabled = False
    While BaseNode.Children > 0
        tvBrowse.Nodes.Remove BaseNode.Child.Index
    Wend
    DoEvents

    AddFolders BaseNode
    If chkShowFiles.Value = vbChecked Then AddFiles BaseNode
    tvBrowse.Enabled = True
    DoEvents
End Sub
Private Sub AddFolders(BaseNode As MSComctlLib.Node)
    Dim fd As PT_FILE_INFO
    Dim hnd As Long
    Dim result As Long
    Dim s As String
    Dim NewNode As MSComctlLib.Node
    Dim i As Long
    Dim SortTable() As BROWSE_SORT
    Dim SortCnt As Long
    Dim FullPath As String

    FullPath = ExtractPath(BaseNode)
    s = FullPath
    If Right(s, 1) <> "\" Then s = s & "\"
'    s = s & "*.*"

    SortCnt = 0
    Erase SortTable
    hnd = ApiFindFirstFile(fd, s, "*.*")
    Do While fd.Continue
        If FilterFolder(fd) Then
            ReDim Preserve SortTable(SortCnt)
            SortTable(SortCnt).Key = fd.Name
            If (fd.Attrs And PT_ATTR_HIDDEN) > 0 Then
                SortTable(SortCnt).Hidden = True
            Else
                SortTable(SortCnt).Hidden = False
            End If
            SortCnt = SortCnt + 1
            If (SortCnt Mod 50) = 0 Then DoEvents
        End If
        result = ApiFindNextFile(fd)
    Loop
    result = ApiFindClose(fd)
    DoEvents

    If SortCnt <= 0 Then Exit Sub
    BrowseSort SortTable, True
    For i = 0 To SortCnt - 1
        If SortTable(i).Hidden Then
            Set NewNode = AddChild(BaseNode, KEY_FOLDER_HIDDEN, SortTable(i).Key, IMAGE_FOLDER_HIDDEN, TAG_FOLDER_HIDDEN)
            AddDummy NewNode
        Else
            Set NewNode = AddChild(BaseNode, KEY_FOLDER, SortTable(i).Key, IMAGE_FOLDER, TAG_FOLDER)
            AddDummy NewNode
        End If
        KeyCnt = KeyCnt + 1
        If (KeyCnt Mod 50) = 0 Then DoEvents
    Next i
End Sub
Private Sub BrowseSort(ByRef SortTable() As BROWSE_SORT, Optional Ascending As Boolean = True)
    Dim lim As Long
    Dim i As Long
    Dim hold As BROWSE_SORT
    Dim Sorted As Boolean
    Dim Passes As Long

    lim = UBound(SortTable)
    Sorted = False
    Passes = 0

    If Ascending Then
        While Not Sorted
            Sorted = True
            For i = 1 To lim
'                If SortTable(i).Key < SortTable(i - 1).Key Then
                If StrComp(SortTable(i).Key, SortTable(i - 1).Key, vbTextCompare) < 0 Then
                    hold = SortTable(i)
                    SortTable(i) = SortTable(i - 1)
                    SortTable(i - 1) = hold
                    Sorted = False
                End If
            Next i
        Wend
    Else
        While Not Sorted
            Sorted = True
            For i = 1 To lim
                If SortTable(i).Key > SortTable(i - 1).Key Then
                    hold = SortTable(i)
                    SortTable(i) = SortTable(i - 1)
                    SortTable(i - 1) = hold
                    Sorted = False
                End If
            Next i
            Passes = Passes + 1
        Wend
    End If
End Sub
Private Function FilterFolder(fd As PT_FILE_INFO) As Boolean
    Dim s As String

    s = fd.Name
    FilterFolder = False

    If (s = ".") Or (s = "..") Then Exit Function
    If (fd.Attrs And PT_ATTR_DIRECTORY) = 0 Then Exit Function
    If Not ApiTextMatch(s, mFolderFilter) Then Exit Function
    If (fd.Attrs And PT_ATTR_HIDDEN) And (Not mShowHidden) Then Exit Function
    If (fd.Attrs And PT_ATTR_SYSTEM) And (Not mShowSystem) Then Exit Function
    If (fd.Attrs And PT_ATTR_READONLY) And (Not mShowReadOnly) Then Exit Function

    FilterFolder = True
End Function
Private Function AddChild(Parent As MSComctlLib.Node, Key As String, Text As String, Img As String, Optional Tag As String = "") As MSComctlLib.Node
    Dim NewNode As MSComctlLib.Node
    Dim NewKey As String

    If Len(Key) > 1 Then
        NewKey = Key
    Else
        NewKey = Key & CStr(KeyCnt)
    End If
    NewKey = NewKey & "X"

    On Error Resume Next
    Set NewNode = tvBrowse.Nodes.Add(Parent, tvwChild, NewKey, Text, Img)
    If (NewNode Is Nothing) Or (Err.Number <> 0) Then
        If Key = KEY_FOLDER_DEF Then
            KEY_FOLDER = KEY_FOLDER_ALT
            Key = KEY_FOLDER
            Set NewNode = tvBrowse.Nodes.Add(Parent, tvwChild, NewKey, Text, Img)
        End If
    End If

    On Error GoTo 0
    If (NewNode Is Nothing) Or (Err.Number <> 0) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

    If Len(Tag) > 0 Then
        NewNode.Tag = Tag
    Else
        NewNode.Tag = Img
    End If
    KeyCnt = KeyCnt + 1

    Set AddChild = NewNode
End Function
Private Function AddDummy(BaseNode As MSComctlLib.Node) As MSComctlLib.Node
    Dim NewNode As MSComctlLib.Node

    Set NewNode = tvBrowse.Nodes.Add(BaseNode, tvwChild, KEY_DUMMY & CStr(KeyCnt) & "X", "")
    NewNode.Tag = TAG_DUMMY
    KeyCnt = KeyCnt + 1

    Set AddDummy = NewNode
End Function
Private Sub AddFiles(BaseNode As MSComctlLib.Node)
    Dim fd As PT_FILE_INFO
    Dim hnd As Long
    Dim result As Long
    Dim s As String
    Dim NewNode As MSComctlLib.Node
    Dim i As Long
    Dim SortTable() As BROWSE_SORT
    Dim SortCnt As Long
    Dim SizeFiles As Variant
    Dim CntFile As Long
    Dim CntDir As Long
    Dim FullPath As String

    FullPath = ExtractPath(BaseNode)
    s = FullPath
    If Right(s, 1) <> "\" Then s = s & "\"
'    s = s & "*.*"

    SortCnt = 0
    SizeFiles = CDec(0)
    Erase SortTable
    hnd = ApiFindFirstFile(fd, s, "*.*")
    Do While fd.Continue
        If (fd.Attrs And PT_ATTR_DIRECTORY) <> 0 Then
            If FilterFolder(fd) Then CntDir = CntDir + 1
        ElseIf FilterFile(fd) Then
            CntFile = CntFile + 1
            SizeFiles = SizeFiles + fd.Size
            ReDim Preserve SortTable(SortCnt)
            SortTable(SortCnt).Key = fd.Name
            If (fd.Attrs And PT_ATTR_HIDDEN) > 0 Then
                SortTable(SortCnt).Hidden = True
            Else
                SortTable(SortCnt).Hidden = False
            End If
            SortCnt = SortCnt + 1
            If (SortCnt Mod 50) = 0 Then DoEvents
        End If
        result = ApiFindNextFile(fd)
    Loop
    result = ApiFindClose(fd)
    DoEvents

    s = vbCrLf & "Sub-Folders: " & FormatNumber(CntDir, 0) & vbCrLf & "Files: " & FormatNumber(CntFile, 0) & vbCrLf & "Total Size: " & FormatNumber(SizeFiles, 0) & " bytes"
    On Error Resume Next
    FileCnts.Remove FullPath
    Err.Clear
    On Error GoTo 0
    FileCnts.Add s, FullPath

    If SortCnt <= 0 Then Exit Sub
    BrowseSort SortTable, True

    For i = 0 To SortCnt - 1
        If SortTable(i).Hidden Then
            AddChild BaseNode, KEY_FILE, SortTable(i).Key, IMAGE_FILE_HIDDEN, TAG_FILE
        Else
            AddChild BaseNode, KEY_FILE, SortTable(i).Key, IMAGE_FILE, TAG_FILE
        End If
        If (KeyCnt Mod 100) = 0 Then DoEvents
    Next i
End Sub
Private Function FilterFile(fd As PT_FILE_INFO) As Boolean
    Dim s As String

    s = fd.Name
    FilterFile = False

    If (s = ".") Or (s = "..") Then Exit Function
    If (fd.Attrs And PT_ATTR_DIRECTORY) <> 0 Then Exit Function
    If Not ApiTextMatch(s, mFileFilter) Then Exit Function
    If (fd.Attrs And PT_ATTR_HIDDEN) And (Not mShowHidden) Then Exit Function
    If (fd.Attrs And PT_ATTR_SYSTEM) And (Not mShowSystem) Then Exit Function
    If (fd.Attrs And PT_ATTR_READONLY) And (Not mShowReadOnly) Then Exit Function

    FilterFile = True
End Function
Public Sub About()
Attribute About.VB_Description = "Show information about this control."
Attribute About.VB_UserMemId = -552
'    Dim s As String
'
'    Load frmAbout
'    frmAbout.lblTitle = BROWSER_DESC
'    frmAbout.Caption = "About " & BROWSER_DESC & " v" & BROWSER_VER
'    frmAbout.lblTitle.Caption = BROWSER_DESC
'    s = Replace(BROWSER_COPYRIGHT, "\n", vbCrLf, , , vbTextCompare)
'    frmAbout.lblVersion.Caption = "Version " & BROWSER_VER
'    frmAbout.lblCopyright.Caption = Replace(s, "(c)", "©")
'    frmAbout.txtComments.Text = Replace(BROWSER_COMMENTS, "\n", vbCrLf, , , vbTextCompare)
'    frmAbout.Show vbModal
End Sub
Public Property Get Speedy() As Boolean
    Speedy = mSpeedy
End Property
Public Property Let Speedy(ByVal b As Boolean)
    mSpeedy = b
'    PropertyChanged "Speedy"
End Property
Public Property Get DetailsVisible() As Boolean
    DetailsVisible = mDetailsVisible
End Property
Public Property Let DetailsVisible(ByVal b As Boolean)
    If mDetailsVisible = b Then Exit Property

    mDetailsVisible = b
    fraDetails.Visible = mDetailsVisible
    ControlResize
    PropertyChanged "DetailsVisible"
End Property
Public Property Get NewFolderVisible() As Boolean
Attribute NewFolderVisible.VB_Description = "Determines whether the ""New Folder"" button is visible."
Attribute NewFolderVisible.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    NewFolderVisible = mNewFolderVisible
End Property
Public Property Let NewFolderVisible(ByVal b As Boolean)
    If mNewFolderVisible = b Then Exit Property

    mNewFolderVisible = b
    btnNewFolder.Visible = mNewFolderVisible
    ControlResize
    PropertyChanged "NewFolderVisible"
End Property
Public Property Get ShowFilesVisible() As Boolean
Attribute ShowFilesVisible.VB_Description = "Determines whether the ""Show Files"" checkbox is visible."
Attribute ShowFilesVisible.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowFilesVisible = mShowFilesVisible
End Property
Public Property Let ShowFilesVisible(ByVal b As Boolean)
    If mShowFilesVisible = b Then Exit Property

    mShowFilesVisible = b
    chkShowFiles.Visible = mShowFilesVisible
    ControlResize
    PropertyChanged "ShowFilesVisible"
End Property
Public Property Get ShowNetworkVisible() As Boolean
Attribute ShowNetworkVisible.VB_Description = "Determines whether the ""Show Network"" checkbox is visible."
Attribute ShowNetworkVisible.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowNetworkVisible = mShowNetworkVisible
End Property
Public Property Let ShowNetworkVisible(ByVal b As Boolean)
    If mShowNetworkVisible = b Then Exit Property

    mShowNetworkVisible = b
    chkShowNetwork.Visible = mShowNetworkVisible
    ControlResize
    PropertyChanged "ShownetworkVisible"
End Property
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Sets or Gets Font used by all of Control's elements."
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal NewFont As StdFont)
    Dim Ctrl As Control

    Set UserControl.Font = NewFont
    On Error Resume Next
    For Each Ctrl In UserControl.Controls
        Set Ctrl.Font = NewFont
    Next Ctrl
    PropertyChanged "Font"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Sets or Gets Name of Font  used by all of the control's elements."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal NewName As String)
    Dim Ctrl As Control

    UserControl.FontName = NewName
    On Error Resume Next
    For Each Ctrl In UserControl.Controls
        Set Ctrl.FontName = NewName
    Next Ctrl
    PropertyChanged "FontName"
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Sets or Gets Font Size used by all of the control's elements."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal NewSize As Single)
    Dim Ctrl As Control

    UserControl.FontSize = NewSize
    On Error Resume Next
    For Each Ctrl In UserControl.Controls
        Set Ctrl.FontSize = NewSize
    Next Ctrl
    PropertyChanged "FontSize"
End Property
Public Property Get Height() As Single
Attribute Height.VB_Description = "Sets or Gets Control's Height."
Attribute Height.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Height = UserControl.Height
End Property
Public Property Let Height(ByVal NewHeight As Single)
    If NewHeight < MinHeight Then
        UserControl.Height = MinHeight
    Else
        UserControl.Height = NewHeight
    End If
    mHeight = UserControl.Height
    PropertyChanged "Height"
End Property
Public Property Get Width() As Single
Attribute Width.VB_Description = "Sets or Gets Width of control."
Attribute Width.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Width = UserControl.Width
End Property
Public Property Let Width(ByVal NewWidth As Single)
    If NewWidth > MinWidth Then
        UserControl.Width = NewWidth
    Else
        UserControl.Width = MinWidth
    End If
    mWidth = UserControl.Width
    PropertyChanged "Width"
End Property
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Set Control's Appearance (flat or 3D). Cannot be changed at run-time."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(ByVal NewAppearance As AppearanceConstants)
    If Ambient.UserMode Then
'        Err.Raise 382
        Exit Property
    End If

    UserControl.Appearance = NewAppearance
    mAppearance = UserControl.Appearance
    PropertyChanged "Appearance"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gets or Sets the control's Enabled state."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = mEnabled
End Property
Public Property Let Enabled(ByVal b As Boolean)
    If b = mEnabled Then Exit Property
    mEnabled = b

    UserControl.Enabled = mEnabled
    PropertyChanged "Enabled"
End Property
Public Property Get Comments() As String
    Comments = Replace(BROWSER_COMMENTS, "\n", vbCrLf, , , vbTextCompare)
End Property
Public Property Get Copyright() As String
    Copyright = Replace(BROWSER_COPYRIGHT, "\n", vbCrLf, , , vbTextCompare)
End Property
Public Property Get Version() As String
    Version = BROWSER_VER
End Property
Public Property Get Description() As String
    Description = BROWSER_DESC
End Property
Public Property Get Caption() As String
    Caption = lblBrowse.Caption
End Property
Public Property Let Caption(ByVal NewCaption As String)
    lblBrowse.Caption = NewCaption
    mCaption = NewCaption
    PropertyChanged "Caption"
End Property
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Gets or Sets the control's Visible state."
Attribute Visible.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Visible = mVisible
End Property
Public Property Let Visible(ByVal b As Boolean)
    Dim Ctrl As Control

    If mVisible = b Then Exit Property
    mVisible = b

    On Error Resume Next
    For Each Ctrl In UserControl.ContainedControls
        Ctrl.Visible = b
    Next Ctrl

    If mVisible Then
        chkShowNetwork.Visible = mShowNetworkVisible
        chkShowFiles.Visible = mShowFilesVisible
        btnNewFolder.Visible = mNewFolderVisible
    End If

    PropertyChanged "Visible"
End Property
Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Sets or Gets Control's border style."
Attribute Border.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Border.VB_UserMemId = -504
    If UserControl.BorderStyle = 0 Then
        Border = False
    Else
        Border = True
    End If
End Property
Public Property Let Border(ByVal b As Boolean)
    If b Then
        UserControl.BorderStyle = 1
    Else
        UserControl.BorderStyle = 0
    End If
    mBorder = b
    PropertyChanged "BorderStyle"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property
Public Property Let BackColor(NewColor As OLE_COLOR)
    If mBackColor = NewColor Then Exit Property

    mBackColor = NewColor
    picBorder.BackColor = mBackColor
    lblBrowse.BackColor = mBackColor
    chkShowFiles.BackColor = mBackColor
    chkShowNetwork.BackColor = mBackColor
    fraDetails.BackColor = mBackColor
    txtDetails.BackColor = mBackColor
    PropertyChanged "BackColor"
End Property
Public Property Get BackStyle() As Long
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(NewStyle As Long)
    Err.Raise PROPERTY_READONLY
End Property
Public Property Get ShowLocal() As Boolean
Attribute ShowLocal.VB_Description = "Determines if Local Drives are displayed."
Attribute ShowLocal.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowLocal = mShowLocal
End Property
Public Property Let ShowLocal(ByVal b As Boolean)
    Dim n As MSComctlLib.Node

    If mShowLocal = b Then Exit Property
    mShowLocal = b

    If b Then
        Set n = tvBrowse.Nodes.Add(, , KEY_LOCALROOT, "My Computer", IMAGE_LOCALROOT) ', TAG_LOCALROOT)
        n.Tag = TAG_LOCALROOT
        Set LocalRoot = n
        AddDummy n
        n.EnsureVisible
    Else    ' delete local tree
        Set n = tvBrowse.Nodes(KEY_LOCALROOT)
        tvBrowse.Nodes.Remove n.Index
    End If
    PropertyChanged "ShowLocal"
End Property
Public Property Get ShowNetwork() As Boolean
Attribute ShowNetwork.VB_Description = "Determines whether Network Providers/Domains/Shares are shown."
Attribute ShowNetwork.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowNetwork = mShowNetwork
End Property
Public Property Let ShowNetwork(ByVal b As Boolean)
    Dim n As MSComctlLib.Node

    If mShowNetwork = b Then Exit Property
    mShowNetwork = b

    If b Then
        'Set tvBrowse.ImageList = ilIcons
        Set n = tvBrowse.Nodes.Add(, , KEY_NETROOT, "Network", IMAGE_NETROOT)
        n.Tag = TAG_NETROOT
        AddDummy n
        n.EnsureVisible
        chkShowNetwork.Value = vbChecked
    Else    ' delete network tree
        Set n = tvBrowse.Nodes(KEY_NETROOT)
        tvBrowse.Nodes.Remove n.Index
        chkShowNetwork.Value = vbUnchecked
    End If
    PropertyChanged "ShowNetwork"
End Property
Public Property Get ShowHidden() As Boolean
Attribute ShowHidden.VB_Description = "Determines whether Hidden Folders and Files will be shown."
Attribute ShowHidden.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowHidden = mShowHidden
End Property
Public Property Let ShowHidden(ByVal b As Boolean)
    If b = mShowHidden Then Exit Property
    mShowHidden = b
    PropertyChanged "ShowHidden"

    If mShowFiles Then
        ShowFiles = False
        ShowFiles = True
    End If
    ' should refresh folders too.
End Property
Public Property Get ShowSystem() As Boolean
Attribute ShowSystem.VB_Description = "Determines whether System folders and files are shown."
Attribute ShowSystem.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowSystem = mShowSystem
End Property
Public Property Let ShowSystem(ByVal b As Boolean)
    If b = mShowSystem Then Exit Property
    mShowSystem = b
    PropertyChanged "ShowSystem"

    If mShowFiles Then
        ShowFiles = False
        ShowFiles = True
    End If
    ' should refresh folders too.
End Property
Public Property Get ShowReadOnly() As Boolean
Attribute ShowReadOnly.VB_Description = "Determines whether Read-Only folders and files are shown."
Attribute ShowReadOnly.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowReadOnly = mShowReadOnly
End Property
Public Property Let ShowReadOnly(ByVal b As Boolean)
    If b = mShowReadOnly Then Exit Property
    mShowReadOnly = b
    PropertyChanged "ShowReadOnly"

    If mShowFiles Then
        ShowFiles = False
        ShowFiles = True
    End If
    ' should refresh folders too.
End Property
Public Property Get IsSpecialFolder() As Boolean
    IsSpecialFolder = mIsSpecialFolder
End Property
Public Property Get SpecialFolderDesc() As String
    SpecialFolderDesc = mSpecialFolderDesc
End Property
Public Property Let SpecialFolderDesc(ByVal NewDesc As String)
    Dim n As MSComctlLib.Node

    If LocalRoot Is Nothing Then ' not displaying local folders
        Exit Property
    End If

    If SpecialRoot Is Nothing Then ExpandLocal LocalRoot
    LocalRoot.Expanded = True
    SpecialRoot.Expanded = True

    Set n = SpecialRoot.Child
    Do While Not (n Is Nothing)
        If StrComp(n.Text, NewDesc, vbTextCompare) = 0 Then
            MakeSelection n
            Exit Property
        End If
        Set n = n.Next
    Loop
End Property
Public Property Get FullPath() As String
Attribute FullPath.VB_Description = "Returns Full Path (Folder + Filename) of currently selected item."
Attribute FullPath.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Dim SaveSep As String

    If tvBrowse.SelectedItem Is Nothing Then
        FullPath = ""
        Exit Property
    End If

    SaveSep = tvBrowse.PathSeparator
    tvBrowse.PathSeparator = vbTab

    FullPath = tvBrowse.SelectedItem.FullPath
    tvBrowse.PathSeparator = SaveSep
End Property
Public Property Let FullPath(ByVal NewPath As String)
    Dim sa() As String
    Dim Level As Long
    Dim Parent As MSComctlLib.Node
    Dim Child As MSComctlLib.Node
    Dim Sibling As MSComctlLib.Node

    If (Len(NewPath) <= 0) Or (tvBrowse.Nodes.Count <= 0) Then
        MakeSelection Nothing
        Exit Property
    End If

    sa = Split(NewPath, "\")
    If InStr(1, NewPath, ":") > 0 Then ' drive letter
        Set Parent = tvBrowse.Nodes(KEY_LOCALROOT)
        Parent.Expanded = True
        Set Child = Parent.Child
        If ExtractKeyCode(Child) = KEY_SPECIALROOT Then
            Set Child = Child.Next
        End If
        Level = 0
    ElseIf left(NewPath, 2) = "\\" Then ' network path
        Set Parent = tvBrowse.Nodes(KEY_NETROOT)
        Parent.Expanded = True
        If MWN_PRESENT Then
            Set Parent = tvBrowse.Nodes(KEY_CONTAINER_MWN)
            Parent.Expanded = True
            If Parent.Children > 0 Then
                Set Parent = Parent.Child
                Parent.Expanded = True
            End If
            Set Child = Parent.Child
        Else
            Set Child = Parent.Child
        End If
        Level = 2
        sa(Level) = "\\" & sa(Level)
    Else
        Set Parent = tvBrowse.Nodes(KEY_LOCALROOT).Root
        Parent.Expanded = True
        Set Child = Parent.FirstSibling
        Level = 0
    End If

    For Level = Level To UBound(sa)
        If Len(sa(Level)) > 0 Then
            Set Sibling = Child
            Do While Not (Sibling Is Nothing)
                If StrComp(sa(Level), Sibling.Text, vbTextCompare) = 0 Then Exit Do
                Set Sibling = Sibling.Next
            Loop
            If Sibling Is Nothing Then
                Err.Raise 76 ' path not found
                MakeSelection Nothing
                Exit Property
            End If
            Sibling.Expanded = True
            Set Child = Sibling.Child
        End If
    Next Level

    MakeSelection Sibling
End Property
Public Property Get Folder() As String
Attribute Folder.VB_Description = "Sets or Gets currently-selected Folder."
Attribute Folder.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Folder = mFolder
End Property
Public Property Let Folder(ByVal FolderName As String)
    FullPath = FolderName
End Property
Public Property Get File() As String
Attribute File.VB_Description = "Sets or Gets currently-selected File."
Attribute File.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    File = mFile
End Property
Public Property Let File(ByVal Filename As String)
    Err.Raise PROPERTY_READONLY
End Property
Public Property Get FileFilter() As String
Attribute FileFilter.VB_Description = "Sets or Gets Filter that determines which files will be shown."
Attribute FileFilter.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FileFilter = mFileFilter
End Property
Public Property Let FileFilter(ByVal NewFilter As String)
    If StrComp(mFileFilter, NewFilter, vbTextCompare) = 0 Then
        mFileFilter = NewFilter
        Exit Property
    End If

    mFileFilter = NewFilter
    Set FileCnts = New Collection
    DispFolderInfo tvBrowse.SelectedItem
    If Not mShowFiles Then Exit Property

    ShowFiles = False
    ShowFiles = True
    PropertyChanged "FileFilter"
End Property
Public Property Get FolderFilter() As String
Attribute FolderFilter.VB_Description = "Sets or Gets Filter that determines which folders will be shown."
Attribute FolderFilter.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    FolderFilter = mFolderFilter
End Property
Public Property Let FolderFilter(ByVal NewFilter As String)
    If StrComp(mFolderFilter, NewFilter, vbTextCompare) = 0 Then
        mFolderFilter = NewFilter
        Exit Property
    End If

    mFolderFilter = NewFilter
    ' should clear tvBrowse, and re-add LocalRoot and NetRoot
    PropertyChanged "FolderFilter"
End Property
Public Property Get ShowFiles() As Boolean
Attribute ShowFiles.VB_Description = "Determines whether the files found in a folder are shown on-screen when a folder is selected."
Attribute ShowFiles.VB_ProcData.VB_Invoke_Property = "PropertyPage2"
    ShowFiles = mShowFiles
    PropertyChanged "ShowFiles"
End Property
Public Property Let ShowFiles(ByVal b As Boolean)
    If b = mShowFiles Then Exit Property

    mShowFiles = b
    If mShowFiles Then
        chkShowFiles.Value = vbChecked
        fraDetails.Caption = "Folder/File Details: "
    Else
        chkShowFiles.Value = vbUnchecked
        fraDetails.Caption = "Selected Folder: "
    End If

    DispFiles
    PropertyChanged "ShowFiles"
End Property


