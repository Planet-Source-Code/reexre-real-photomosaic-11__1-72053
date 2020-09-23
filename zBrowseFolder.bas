Attribute VB_Name = "zBrowseFolder"
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BFFM_SETSELECTION = (WM_USER + 102)

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
