VERSION 5.00
Begin VB.UserControl vbalScrollButtonCtl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      ScaleHeight     =   255
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1425
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdButton 
      Height          =   195
      Index           =   0
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkButton 
      Height          =   195
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "vbalScrollButtonCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' API declares
' ---------------------------------------------------------------------------
'ToolTIp Scroll Stuff
' WinAPI declares:
Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_TOOLWINDOW = &H80&

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48&

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const WM_ACTIVATE = &H6
Private Const WM_KEYDOWN = &H100

'----------------------------------------------

' Scroll bar stuff
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Const WM_PAINT = &HF

Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
   Private Const SB_BOTH = 3
   Private Const SB_BOTTOM = 7
   Private Const SB_CTL = 2
   Private Const SB_ENDSCROLL = 8
   Private Const SB_HORZ = 0
   Private Const SB_LEFT = 6
   Private Const SB_LINEDOWN = 1
   Private Const SB_LINELEFT = 0
   Private Const SB_LINERIGHT = 1
   Private Const SB_LINEUP = 0
   Private Const SB_PAGEDOWN = 3
   Private Const SB_PAGELEFT = 2
   Private Const SB_PAGERIGHT = 3
   Private Const SB_PAGEUP = 2
   Private Const SB_RIGHT = 7
   Private Const SB_THUMBPOSITION = 4
   Private Const SB_THUMBTRACK = 5
   Private Const SB_TOP = 6
   Private Const SB_VERT = 1
   
   Private Const SIF_RANGE = &H1
   Private Const SIF_PAGE = &H2
   Private Const SIF_POS = &H4
   Private Const SIF_DISABLENOSCROLL = &H8
   Private Const SIF_TRACKPOS = &H10
   Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
   
   Private Const ESB_DISABLE_BOTH = &H3
   Private Const ESB_ENABLE_BOTH = &H0
   
   Private Const SBS_HORZ = &H0&
   Private Const SBS_VERT = &H1&
   Private Const SBS_TOPALIGN = &H2&
   Private Const SBS_LEFTALIGN = &H2&
   Private Const SBS_BOTTOMALIGN = &H4&
   Private Const SBS_RIGHTALIGN = &H4&
   Private Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
   Private Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
   Private Const SBS_SIZEBOX = &H8&
   Private Const SBS_SIZEGRIP = &H10&
   
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

' Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL = &H1&
Private Const WSB_PROP_CXHSCROLL = &H2&
Private Const WSB_PROP_CYHSCROLL = &H4&
Private Const WSB_PROP_CXVSCROLL = &H8&
Private Const WSB_PROP_CXHTHUMB = &H10&
Private Const WSB_PROP_CYVTHUMB = &H20&
Private Const WSB_PROP_VBKGCOLOR = &H40&
Private Const WSB_PROP_HBKGCOLOR = &H80&
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const WSB_PROP_WINSTYLE = &H400&
Private Const WSB_PROP_PALETTE = &H800&
Private Const WSB_PROP_MASK = &HFFF&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Declare Function FlatSB_EnableScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function FlatSB_GetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal LPINT1 As Long, ByVal LPINT2 As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_GetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long) As Long
Private Declare Function FlatSB_GetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal propIndex As Long, ByVal LPINT As Long) As Long

Private Declare Function FlatSB_SetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal pos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal Min As Long, ByVal Max As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function InitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long


Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114
Private Const WM_LBUTTONUP = &H202

Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' Windows General
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const CW_USEDEFAULT = &H80000000
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Private Const WS_EX_LEFTSCROLLBAR = &H4000&
Private Const WS_EX_RIGHTSCROLLBAR = &H0&
' Window relationship functions:

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Show window styles
Private Const SW_SHOWNORMAL = 1
Private Const SW_ERASE = &H4
Private Const SW_HIDE = 0
Private Const SW_INVALIDATE = &H2
Private Const SW_MAX = 10
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_OTHERUNZOOM = 4
Private Const SW_OTHERZOOM = 2
Private Const SW_PARENTCLOSING = 1
Private Const SW_RESTORE = 9
Private Const SW_PARENTOPENING = 3
Private Const SW_SHOW = 5
Private Const SW_SCROLLCHILDREN = &H1
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4


Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

' Button messages:
Private Const BM_GETCHECK = &HF0&
Private Const BM_SETCHECK = &HF1&
Private Const BM_GETSTATE = &HF2&
Private Const BM_SETSTATE = &HF3&
Private Const BM_SETSTYLE = &HF4&
Private Const BM_CLICK = &HF5&
Private Const BM_GETIMAGE = &HF6&
Private Const BM_SETIMAGE = &HF7&

Private Const BST_UNCHECKED = &H0&
Private Const BST_CHECKED = &H1&
Private Const BST_INDETERMINATE = &H2&
Private Const BST_PUSHED = &H4&
Private Const BST_FOCUS = &H8&

' Button notifications:
Private Const BN_CLICKED = 0&
Private Const BN_PAINT = 1&
Private Const BN_HILITE = 2&
Private Const BN_UNHILITE = 3&
Private Const BN_DISABLE = 4&
Private Const BN_DOUBLECLICKED = 5&
Private Const BN_PUSHED = BN_HILITE
Private Const BN_UNPUSHED = BN_UNHILITE
Private Const BN_DBLCLK = BN_DOUBLECLICKED
Private Const BN_SETFOCUS = 6&
Private Const BN_KILLFOCUS = 7&

' Button Styles:
Private Const BS_3STATE = &H5&
Private Const BS_AUTO3STATE = &H6&
Private Const BS_AUTOCHECKBOX = &H3&
Private Const BS_AUTORADIOBUTTON = &H9&
Private Const BS_CHECKBOX = &H2&
Private Const BS_DEFPUSHBUTTON = &H1&
Private Const BS_GROUPBOX = &H7&
Private Const BS_LEFTTEXT = &H20&
Private Const BS_OWNERDRAW = &HB&
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_RADIOBUTTON = &H4&
Private Const BS_USERBUTTON = &H8&
Private Const BS_ICON = &H40&
Private Const BS_BITMAP = &H80&
Private Const BS_LEFT = &H100&
Private Const BS_RIGHT = &H200&
Private Const BS_CENTER = &H300&
Private Const BS_TOP = &H400&
Private Const BS_BOTTOM = &H800&
Private Const BS_VCENTER = &HC00&
Private Const BS_PUSHLIKE = &H1000&
Private Const BS_MULTILINE = &H2000&
Private Const BS_NOTIFY = &H4000&
Private Const BS_FLAT = &H8000&
Private Const BS_RIGHTBUTTON = BS_LEFTTEXT

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
' Draw an item in an ImageList:
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
' Draw an item in an ImageList with more control over positioning
' and colour:
Private Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" ( _
      ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal rgbBk As Long, _
      ByVal rgbFg As Long, _
      ByVal fStyle As Long _
   ) As Long
' Built in ImageList drawing methods:
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_OVERLAYMASK = 3840
' Use default rgb colour:
Private Const CLR_NONE = -1
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
    ) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

' Standard GDI draw icon function:
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_SIZE = &H5
Private Const WM_CTLCOLORSCROLLBAR = &H137

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, _
   ByVal wParam As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long
Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

' Missing Draw State constants declarations:
'/* Image type */
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

' /* State type */
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Const BF_LEFT = 1
Private Const BF_TOP = 2
Private Const BF_RIGHT = 4
Private Const BF_BOTTOM = 8
Private Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Private Const BDR_RAISEDOUTER = 1
Private Const BDR_SUNKENOUTER = 2
Private Const BDR_RAISEDINNER = 4
Private Const BDR_SUNKENINNER = 8
Private Const BDR_BUTTONPRESSED = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Private Const BDR_BUTTONNORMAL = BDR_RAISEDINNER Or BDR_RAISEDOUTER
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


' XP DrawTheme declares for XP version
Private Declare Function GetVersion Lib "Kernel32" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lhdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, pRect As RECT, _
    ByVal hIml As Long, ByVal iImageIndex As Long) As Long
Private Const S_OK = 0

Implements ISubclass

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type tButtonInfo
   sKey As String
   sHelpText As String
   lIconIndexUp As Long
   lIconIndexDown As Long
   ePosition As ESBCButtonPositionConstants
   bCheck As Boolean
   sCheckGroup As String
   ctlThis As Control
End Type

Private m_hWndControl As Long
Private m_hWndParent As Long
Private m_hWNd As Long
Private m_eScrollType As ESBCScrollTypes
Private m_iButtonCount As Long
Private m_tButtons() As tButtonInfo
Private m_iOptCount As Long
Private m_iCmdCount As Long
Private m_lPos1 As Long
Private m_lPos2 As Long
Private m_hIml As Long
Private m_ptrVb6ImageList As Long
Private m_lIconSizeX As Long
Private m_lIconSizeY As Long
Private m_lSmallChange As Long
Private m_bScrollEnabled As Boolean
Private m_bNoFlatScrollBars As Boolean
Private m_bXPStyleButtons As Boolean

Public Event ButtonClick(ByVal lButton As Long)
Public Event Change()
Public Event Scroll()

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Friend Property Get ButtonKey(ByVal lButton As Long) As String
   If (ButtonIndex(lButton) > 0) Then
      ButtonKey = m_tButtons(lButton).sKey
   End If
End Property

Friend Property Get ButtonToolTipText(ByVal vKey As Variant) As String
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonToolTipText = m_tButtons(iBtnIndex).sHelpText
   End If
End Property

Friend Property Let ButtonToolTipText(ByVal vKey As Variant, ByVal sText As String)
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).sHelpText = sText
      m_tButtons(iBtnIndex).ctlThis.ToolTipText = sText
   End If
End Property

Friend Property Get ButtonVisible(ByVal vKey As Variant) As Boolean
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonVisible = m_tButtons(iBtnIndex).ctlThis.Visible
   End If
End Property

Friend Property Let ButtonVisible(ByVal vKey As Variant, ByVal bState As Boolean)
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).ctlThis.Visible = bState
      UserControl_Resize
   End If
End Property

Friend Property Get ButtonEnabled(ByVal vKey As Variant) As Boolean
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      ButtonEnabled = m_tButtons(iBtnIndex).ctlThis.Enabled
   End If
End Property

Friend Property Let ButtonEnabled(ByVal vKey As Variant, ByVal bState As Boolean)
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      m_tButtons(iBtnIndex).ctlThis.Enabled = bState
   End If
End Property

Friend Property Get ButtonValue(ByVal vKey As Variant) As OLE_TRISTATE
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      If (TypeOf m_tButtons(iBtnIndex).ctlThis Is CommandButton) Then
         ButtonValue = Abs(m_tButtons(iBtnIndex).ctlThis.Value)
      Else
         ButtonValue = m_tButtons(iBtnIndex).ctlThis.Value
      End If
   End If
End Property

Friend Property Let ButtonValue(ByVal vKey As Variant, oValue As OLE_TRISTATE)
   Dim iBtnIndex As Long
   iBtnIndex = ButtonIndex(vKey)
   If (iBtnIndex <> 0) Then
      If (TypeOf m_tButtons(iBtnIndex).ctlThis Is CommandButton) Then
         m_tButtons(iBtnIndex).ctlThis.Value = -1 * oValue
      Else
         m_tButtons(iBtnIndex).ctlThis.Value = oValue
      End If
   End If
End Property

Private Function pTranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, pTranslateColor) Then
        pTranslateColor = CLR_INVALID
    End If
End Function

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property

Private Sub pDrawImage( _
      ByVal ptrVB6ImageList As Long, _
      ByVal hIml As Long, _
      ByVal iIndex As Long, _
      ByVal hdc As Long, _
      ByVal xPixels As Integer, _
      ByVal yPixels As Integer, _
      ByVal lIconSizeX As Long, ByVal lIconSizeY As Long, _
      Optional ByVal bSelected = False, _
      Optional ByVal bCut = False, _
      Optional ByVal bDisabled = False, _
      Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
      Optional ByVal hExternalIml As Long = 0 _
    )
Dim hIcon As Long
Dim lFlags As Long
Dim lhIml As Long
Dim lColor As Long
Dim iImgIndex As Long

   ' Draw the image at 1 based index or key supplied in vKey.
   ' on the hDC at xPixels,yPixels with the supplied options.
   ' You can even draw an ImageList from another ImageList control
   ' if you supply the handle to hExternalIml with this function.
   
   iImgIndex = iIndex
   If (iImgIndex > -1) Then
      If (hExternalIml <> 0) Then
          lhIml = hExternalIml
      Else
          lhIml = hIml
      End If
      
      lFlags = ILD_TRANSPARENT
      If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
      End If
      
      If (bCut) Then
        ' Draw dithered:
        lColor = pTranslateColor(oCutDitherColour)
        If (lColor = -1) Then lColor = pTranslateColor(vbWindowBackground)
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hdc, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, lColor, _
              lFlags
      ElseIf (bDisabled) Then
         If (ptrVB6ImageList <> 0) Then
            Dim o As Object
            On Error Resume Next
            Set o = ObjectFromPtr(ptrVB6ImageList)
            If Not (o Is Nothing) Then
                hIcon = o.ListImages(iImgIndex + 1).ExtractIcon()
            End If
            On Error GoTo 0
         Else
            ' extract a copy of the icon:
            hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
         End If
         If (hIcon <> 0) Then
            ' Draw it disabled at x,y:
            DrawState hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED
            ' Clear up the icon:
            DestroyIcon hIcon
         End If
      Else
         If (ptrVB6ImageList <> 0) Then
             On Error Resume Next
             Set o = ObjectFromPtr(ptrVB6ImageList)
             If Not (o Is Nothing) Then
                 o.ListImages(iImgIndex + 1).Draw hdc, xPixels * Screen.TwipsPerPixelX, yPixels * Screen.TwipsPerPixelY, lFlags
             End If
             On Error GoTo 0
         Else
            ' Standard draw:
            ImageList_Draw _
                lhIml, _
                iImgIndex, _
                hdc, _
                xPixels, _
                yPixels, _
                lFlags
         End If
      End If
   End If
End Sub

Friend Property Let ImageList(vThis As Variant)
    m_hIml = 0
    m_ptrVb6ImageList = 0
    If (VarType(vThis) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vThis
    ElseIf (VarType(vThis) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vThis.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vThis.hImageList
        If (Err.Number = 0) Then
            ' Check for VB6 image list:
            If (TypeName(vThis) = "ImageList") Then
                If (vThis.ListImages.count <> ImageList_GetImageCount(m_hIml)) Then
                    Dim o As Object
                    Set o = vThis
                    m_ptrVb6ImageList = ObjPtr(o)
                End If
            End If
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        If (m_ptrVb6ImageList <> 0) Then
            m_lIconSizeX = vThis.ImageWidth
            m_lIconSizeY = vThis.ImageHeight
        Else
            Dim rc As RECT
            ImageList_GetImageRect m_hIml, 0, rc
            m_lIconSizeX = rc.Right - rc.Left
            m_lIconSizeY = rc.Bottom - rc.Top
        End If
    End If
End Property

Friend Sub AddButton( _
      Optional ByVal sKey As String = "", _
      Optional ByVal sToolTipText As String = "", _
      Optional ByVal lIconIndexUp As Long = -1, _
      Optional ByVal lIconIndexDown As Long = -1, _
      Optional ByVal ePosition As ESBCButtonPositionConstants = esbcButtonPositionDefault, _
      Optional ByVal bCheck As Boolean = False, _
      Optional ByVal sCheckGroup As String = "", _
      Optional ByVal bVisible As Boolean = True, _
      Optional ByVal vKeyBefore As Variant _
   )
Dim lBtnIndex As Long
Dim iBtn As Long
Dim lStyle As Long

   If (m_eScrollType = esbcSizeGripper) Then
      ' No buttons on size grippers.
      Exit Sub
   End If

   ' Check if inserting a button:
   If Not (IsMissing(vKeyBefore)) Then
      ' Get button:
      lBtnIndex = ButtonIndex(vKeyBefore)
      If (lBtnIndex > 0) Then
         m_iButtonCount = m_iButtonCount + 1
         ReDim Preserve m_tButtons(1 To m_iButtonCount) As tButtonInfo
         ' Shift the array:
         For iBtn = m_iButtonCount To lBtnIndex + 1 Step -1
            LSet m_tButtons(iBtn) = m_tButtons(iBtn - 1)
         Next iBtn
      Else
         Exit Sub
      End If
   Else
      m_iButtonCount = m_iButtonCount + 1
      lBtnIndex = m_iButtonCount
      ReDim Preserve m_tButtons(1 To m_iButtonCount) As tButtonInfo
   End If
   
   ' Set the values:
   With m_tButtons(lBtnIndex)
      .sKey = sKey
      .sHelpText = sToolTipText
      .lIconIndexUp = lIconIndexUp
      .lIconIndexDown = lIconIndexDown
      If (ePosition = esbcButtonPositionDefault) Then
         If (m_eScrollType = esbcHorizontal) Then
            .ePosition = esbcButtonPositionLeftTop
         Else
            .ePosition = esbcButtonPositionRightBottom
         End If
      Else
         .ePosition = ePosition
      End If
      .bCheck = bCheck
      .sCheckGroup = sCheckGroup
      If (bCheck) Then
         m_iOptCount = m_iOptCount + 1
         If (m_iOptCount > 1) Then
            Load chkButton(m_iOptCount - 1)
         End If
         Set .ctlThis = chkButton(m_iOptCount - 1)
      Else
         m_iCmdCount = m_iCmdCount + 1
         If (m_iCmdCount > 1) Then
            Load cmdButton(m_iCmdCount - 1)
         End If
         Set .ctlThis = cmdButton(m_iCmdCount - 1)
      End If
      .ctlThis.Visible = bVisible
      .ctlThis.ToolTipText = sToolTipText
   End With
      
   pResizeButtons
   
End Sub

Friend Property Get ButtonCount() As Long
   ButtonCount = m_iButtonCount
End Property

Friend Property Get ButtonIndex(ByVal vKey As Variant) As Long
Dim lBtn As Long
Dim lIndex As Long
   If (IsNumeric(vKey)) Then
      lBtn = CLng(vKey)
      If (lBtn > 0) And (lBtn <= m_iButtonCount) Then
         lIndex = lBtn
      End If
   Else
      For lBtn = 1 To m_iButtonCount
         If (m_tButtons(lBtn).sKey = vKey) Then
            lIndex = lBtn
            Exit For
         End If
      Next lBtn
   End If
   If (lIndex > 0) Then
      ButtonIndex = lIndex
   Else
      Err.Raise 9, "ntFlexGrid.ScrollBar.ButtonIndex", "Button subscript out of range."
   End If
   
End Property

Public Property Get ScrollType() As ESBCScrollTypes
   ScrollType = m_eScrollType
End Property

Public Property Let ScrollType(ByVal eType As ESBCScrollTypes)
   m_eScrollType = eType
   pCreateScrollControl
   PropertyChanged "ScrollType"
   Resize
End Property

Public Property Get XpStyleButtons() As Boolean
   XpStyleButtons = m_bXPStyleButtons
End Property

Public Property Let XpStyleButtons(ByVal bState As Boolean)
   m_bXPStyleButtons = bState
End Property

Public Property Get Visible() As Boolean
   Visible = UserControl.Extender.Visible
End Property

Public Property Let Visible(ByVal bState As Boolean)
   UserControl.Extender.Visible = bState
   Select Case m_eScrollType
   Case esbcVertical
      If (m_hWndParent <> 0) Then
         SetProp m_hWndParent, "vbalScrollButtons:VERT", Abs(bState)
      End If
   Case esbcHorizontal
      If (m_hWndParent <> 0) Then
         SetProp m_hWndParent, "vbalScrollButtons:HORZ", Abs(bState)
      End If
   End Select
End Property

Public Property Get SmallChange() As Long
   SmallChange = m_lSmallChange
End Property

Public Property Let SmallChange(ByVal lSmallChange As Long)
   m_lSmallChange = lSmallChange
End Property

Public Property Get ScrollEnabled() As Boolean
   Enabled = m_bScrollEnabled
End Property

Public Property Let ScrollEnabled(ByVal bEnabled As Boolean)
   Dim lF As Long
        
   If (bEnabled) Then
      lF = ESB_ENABLE_BOTH
   Else
      lF = ESB_DISABLE_BOTH
   End If
   If (m_bNoFlatScrollBars) Then
      EnableScrollBar m_hWNd, SB_CTL, lF
   Else
      FlatSB_EnableScrollBar m_hWNd, SB_CTL, lF
   End If
    
End Property

Private Sub pGetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)
   
   If (m_bNoFlatScrollBars) Then
       GetScrollInfo m_hWNd, SB_CTL, tSI
   Else
       FlatSB_GetScrollInfo m_hWNd, SB_CTL, tSI
   End If

End Sub

Private Sub pLetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
        
   tSI.fMask = fMask
   tSI.cbSize = LenB(tSI)
   If (m_bNoFlatScrollBars) Then
       SetScrollInfo m_hWNd, SB_CTL, tSI, True
   Else
       FlatSB_SetScrollInfo m_hWNd, SB_CTL, tSI, True
   End If
    
End Sub

Public Property Get Min() As Long
   Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_RANGE
    Min = tSI.nMin
End Property
Public Property Get Max() As Long
   Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage
End Property
Public Property Get Value() As Long
   Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_POS
    Value = tSI.nPos
End Property
Public Property Get LargeChange() As Long
   Dim tSI As SCROLLINFO
    pGetSI tSI, SIF_PAGE
    LargeChange = tSI.nPage
End Property
Public Property Let Min(ByVal iMin As Long)
   Dim tSI As SCROLLINFO
   If iMin < 0 Then
      Debug.Print "BADMIN"
   End If
    tSI.nMin = iMin
    tSI.nMax = Max + (LargeChange - 1)
    pLetSI tSI, SIF_RANGE
End Property
Public Property Let Max(ByVal iMax As Long)
   Dim tSI As SCROLLINFO
    tSI.nMax = iMax + LargeChange
    tSI.nMin = Min
    pLetSI tSI, SIF_RANGE
    pRaiseEvent False
    picTip.Width = TextWidth(CStr(iMax)) + 240
    Label1.Move 0, 0, picTip.ScaleWidth, picTip.ScaleHeight
    Rem picTip.Width = TextWidth("ROW " & CStr(iMax)) + 240
    Rem Label1.Width = picTip.Width
End Property

Public Property Let Value(ByVal iValue As Long)
   Dim tSI As SCROLLINFO
   If iValue < 0 Then
      Debug.Print "BADiValue"
   End If
   Dim lPercent As Long
    If (iValue <> Value) Then
        tSI.nPos = iValue
        pLetSI tSI, SIF_POS
        lPercent = iValue * 100 \ Max
        If (m_eScrollType = esbcHorizontal) Then
           UserControl.Extender.ToolTipText = "Col " & CStr(iValue) ' lPercent & "%"
        ElseIf (m_eScrollType = esbcVertical) Then
            UserControl.Extender.ToolTipText = "Row " & CStr(iValue) ' lPercent & "%"
        End If
        pRaiseEvent False
    End If
End Property

Public Property Let LargeChange(ByVal iLargeChange As Long)
   Dim tSI As SCROLLINFO
   Dim lCurMax As Long
   Dim lCurLargeChange As Long
    
   pGetSI tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pLetSI tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Function pRaiseEvent(ByVal bScroll As Boolean)
   Static s_lLastValue As Long
   If (Value <> s_lLastValue) Then
      If (bScroll) Then
         RaiseEvent Scroll
      Else
         RaiseEvent Change
      End If
      s_lLastValue = Value
   End If
   
End Function

Private Sub pCreateScrollControl()
   Dim lStyle As Long
   Dim lWidth As Long
   Dim lHeight As Long
   
   If (m_hWndParent <> 0) Then
      pDestroyScrollControl
      lStyle = WS_CHILD Or WS_VISIBLE
      If (m_eScrollType = esbcHorizontal) Then
         lStyle = lStyle Or SBS_HORZ And Not SBS_VERT
         lWidth = UserControl.Width \ Screen.TwipsPerPixelX
         lHeight = CW_USEDEFAULT
      ElseIf (m_eScrollType = esbcVertical) Then
         lStyle = lStyle Or SBS_VERT And Not SBS_HORZ
         lHeight = UserControl.Height \ Screen.TwipsPerPixelY
         lWidth = CW_USEDEFAULT
      Else
         lStyle = lStyle Or SBS_SIZEBOX Or SBS_SIZEBOXBOTTOMRIGHTALIGN
      End If
      
      m_hWNd = CreateWindowEx(0, "SCROLLBAR", "", lStyle, 0, 0, lWidth, lHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)
      If (m_hWNd <> 0) Then
         ShowScrollBar m_hWNd, SB_CTL, 1
         If (lStyle And SBS_SIZEBOX) <> SBS_SIZEBOX Then
            AttachMessage Me, m_hWndControl, WM_VSCROLL
            AttachMessage Me, m_hWndControl, WM_HSCROLL
            AttachMessage Me, m_hWndControl, WM_LBUTTONUP
            Min = 0
            Max = 255
            SmallChange = 1
            LargeChange = 32
         Else
            UserControl.BackColor = vbButtonFace
         End If
      End If
   End If
End Sub

Private Sub pDestroyScrollControl()
   If (m_hWNd <> 0) Then
      DetachMessage Me, m_hWndControl, WM_VSCROLL
      DetachMessage Me, m_hWndControl, WM_HSCROLL
      DetachMessage Me, m_hWndControl, WM_LBUTTONUP
      
      ShowWindow m_hWNd, SW_HIDE
      SetParent m_hWNd, 0
      DestroyWindow m_hWNd
   End If
End Sub

Private Sub pResizeButtons()
   Dim lPos1 As Long
   Dim lPos2 As Long
   Dim lBtn As Long
   Dim lExtent As Long
   
   On Error Resume Next
   
   If (m_eScrollType = esbcHorizontal) Then
      lExtent = GetSystemMetrics(SM_CYHSCROLL)
      lPos2 = UserControl.Width - lExtent * Screen.TwipsPerPixelX
   ElseIf (m_eScrollType = esbcVertical) Then
      lExtent = GetSystemMetrics(SM_CXVSCROLL)
      lPos2 = UserControl.Height - lExtent * Screen.TwipsPerPixelY
   Else
      Exit Sub
   End If
   
   For lBtn = 1 To m_iButtonCount
      With m_tButtons(lBtn)
         If (.ctlThis.Visible) Then
            If (.ePosition = esbcButtonPositionLeftTop) Then
               If (m_eScrollType = esbcHorizontal) Then
                  .ctlThis.Move lPos1, 0, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                  lPos1 = lPos1 + lExtent * Screen.TwipsPerPixelX
               Else
                  .ctlThis.Move 0, lPos1, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                  lPos1 = lPos1 + lExtent * Screen.TwipsPerPixelY
               End If
            Else
               If (m_eScrollType = esbcHorizontal) Then
                  .ctlThis.Move lPos2, 0, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                  lPos2 = lPos2 - lExtent * Screen.TwipsPerPixelX
               Else
                  .ctlThis.Move 0, lPos2, lExtent * Screen.TwipsPerPixelX, lExtent * Screen.TwipsPerPixelY
                  lPos2 = lPos2 - lExtent * Screen.TwipsPerPixelY
               End If
            End If
         End If
      End With
   Next lBtn
   m_lPos1 = lPos1
   If (m_eScrollType = esbcHorizontal) Then
      m_lPos2 = lPos2 + lExtent * Screen.TwipsPerPixelX
   Else
      m_lPos2 = lPos2 + lExtent * Screen.TwipsPerPixelY
   End If

End Sub
Private Sub pResizeScroll()
   Dim x As Long, y As Long
   Dim cx As Long, cy As Long

   If (m_hWNd <> 0) Then
      If (m_eScrollType = esbcHorizontal) Then
         y = 0
         x = m_lPos1 \ Screen.TwipsPerPixelX
         cx = m_lPos2 \ Screen.TwipsPerPixelX - x
         cy = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      ElseIf (m_eScrollType = esbcVertical) Then
         x = 0
         y = m_lPos1 \ Screen.TwipsPerPixelY
         cy = m_lPos2 \ Screen.TwipsPerPixelY - y
         cx = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
      Else
         x = 0
         y = 0
         cx = UserControl.ScaleWidth \ Screen.TwipsPerPixelY
         cy = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      End If
      MoveWindow m_hWNd, x, y, cx, cy, 1
   End If
End Sub

Public Sub Resize()
   Dim tR As RECT
   Dim bLeftScroll As Boolean
   Dim bVert As Boolean
   Dim bHorz As Boolean
   Dim lStyle As Long
   Dim lSize As Long

   GetClientRect m_hWndParent, tR
   ' Determine what other scroll bars on the parent:
   bVert = (GetProp(m_hWndParent, "vbalScrollButtons:VERT") <> 0)
   bHorz = (GetProp(m_hWndParent, "vbalScrollButtons:HORZ") <> 0)
   ' Determine if scroll bars are on the left or right:
   lStyle = GetWindowLong(m_hWndParent, GWL_EXSTYLE)
   If (lStyle And WS_EX_LEFTSCROLLBAR) Then
      bLeftScroll = True
   End If
   
   Select Case m_eScrollType
   Case esbcSizeGripper
      ' Only visible if both horz and vert.
      If (bVert) And (bHorz) And Not (bLeftScroll) Then
         tR.Left = tR.Right - GetSystemMetrics(SM_CXVSCROLL)
         tR.Top = tR.Bottom - GetSystemMetrics(SM_CYHSCROLL)
         MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
         UserControl_Resize
      End If
   Case esbcHorizontal
      ' We resize to the bottom of form.  Horizontal
      ' extent depends on whether Vertical scroll is
      ' visible
      lSize = GetSystemMetrics(SM_CYHSCROLL)
      tR.Top = tR.Bottom - lSize
      If (bVert) Then
         If (bLeftScroll) Then
            tR.Left = tR.Left + GetSystemMetrics(SM_CXVSCROLL)
         Else
            tR.Right = tR.Right - GetSystemMetrics(SM_CXVSCROLL)
         End If
      End If
      MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
      UserControl_Resize
      
   Case esbcVertical
      ' We resize to the right or left of form.  Horizontal
      ' extent depends on whether Vertical scroll is
      ' visible
      lSize = GetSystemMetrics(SM_CXVSCROLL)
      If (bLeftScroll) Then
         tR.Right = tR.Left + lSize
      Else
         tR.Left = tR.Right - lSize
      End If
      If (bHorz) Then
         tR.Bottom = tR.Bottom - GetSystemMetrics(SM_CYHSCROLL)
      End If
      MoveWindow UserControl.hwnd, tR.Left, tR.Top, (tR.Right - tR.Left), (tR.Bottom - tR.Top), 1
      UserControl_Resize
   End Select
End Sub

Private Sub pDrawButton(tDis As DRAWITEMSTRUCT)
   Dim hBr As Long
   Dim lState As Long
   Dim bPushed As Boolean
   Dim bDisabled As Boolean
   Dim bChecked As Boolean
   Dim iBtn As Long
   Dim iBtnIndex As Long
   Dim lSize As Long
   Dim x As Long, y As Long
   Dim bXpStyle As Boolean
   Dim hTheme As Long
   Dim hR As Long

   lState = SendMessageLong(tDis.hwndItem, BM_GETSTATE, 0, 0)
   
   bPushed = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
      
   For iBtn = 1 To m_iButtonCount
      If (m_tButtons(iBtn).ctlThis.hwnd = tDis.hwndItem) Then
         iBtnIndex = iBtn
         bChecked = (m_tButtons(iBtn).ctlThis.Value = Checked)
         bPushed = bPushed Or bChecked
         bDisabled = Not (m_tButtons(iBtnIndex).ctlThis.Enabled)
         Exit For
      End If
   Next iBtn
      
   If (m_bXPStyleButtons) Then
      On Error Resume Next
      hTheme = OpenThemeData(hwnd, StrPtr("Button"))
      If (Err.Number <> 0) Or (hTheme = 0) Then
         bXpStyle = False
      Else
         bXpStyle = True
      End If
   End If
   
   If bChecked Then
      hBr = GetSysColorBrush(vb3DHighlight And &H1F&)
   Else
      hBr = GetSysColorBrush(vbButtonFace And &H1F&)
   End If
   FillRect tDis.hdc, tDis.rcItem, hBr
   DeleteObject hBr
   
   If (bXpStyle) Then
      If bDisabled Then
         hR = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                4, tDis.rcItem, tDis.rcItem)
      ElseIf bChecked Or bPushed Then
         hR = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                3, tDis.rcItem, tDis.rcItem)
      Else
         hR = DrawThemeBackground(hTheme, tDis.hdc, 1, _
                1, tDis.rcItem, tDis.rcItem)
      End If
   End If
   
   If (iBtnIndex > 0) Then
      If (m_eScrollType = esbcHorizontal) Then
         lSize = GetSystemMetrics(SM_CYHSCROLL) - 4
      Else
         lSize = GetSystemMetrics(SM_CXVSCROLL) - 4
      End If
      x = 2 + (lSize - m_lIconSizeX) \ 2
      y = x
      If (bPushed) Then
         x = x + 1
         y = y + 1
         pDrawImage m_ptrVb6ImageList, m_hIml, m_tButtons(iBtnIndex).lIconIndexDown, tDis.hdc, x, y, m_lIconSizeX, m_lIconSizeY, , , bDisabled
      Else
         pDrawImage m_ptrVb6ImageList, m_hIml, m_tButtons(iBtnIndex).lIconIndexUp, tDis.hdc, x, y, m_lIconSizeX, m_lIconSizeY, , , bDisabled
      End If
   End If
   
   If (bXpStyle) Then
   
   Else
      If (bPushed) Then
         DrawEdge tDis.hdc, tDis.rcItem, BDR_SUNKENOUTER, BF_RECT
      Else
         DrawEdge tDis.hdc, tDis.rcItem, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_RECT
      End If
   End If
   
   If (hTheme) Then
      CloseThemeData hTheme
   End If
   
   
End Sub

Private Sub chkButton_Click(index As Integer)
Dim iB As Long
Dim lBtnIndex As Long
   For iB = 1 To m_iButtonCount
      If (m_tButtons(iB).ctlThis Is chkButton(index)) Then
         lBtnIndex = iB
         Exit For
      End If
   Next iB
   If (lBtnIndex > 0) Then
      RaiseEvent ButtonClick(lBtnIndex)
   End If
End Sub

Private Sub cmdButton_Click(index As Integer)
Dim iB As Long
Dim lBtnIndex As Long
   For iB = 1 To m_iButtonCount
      If (m_tButtons(iB).ctlThis Is cmdButton(index)) Then
         lBtnIndex = iB
         Exit For
      End If
   Next iB
   If (lBtnIndex > 0) Then
      RaiseEvent ButtonClick(lBtnIndex)
   End If

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If (CurrentMessage = WM_DRAWITEM) Or (CurrentMessage = WM_CTLCOLORSCROLLBAR) Then
      ISubclass_MsgResponse = emrConsume
   Else
      ISubclass_MsgResponse = emrPreprocess
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDis As DRAWITEMSTRUCT
Dim lBar As Long
Dim lScrollcode As Long
Dim lV As Long, lSC As Long
Dim tSI As SCROLLINFO
Dim bShowTip As Boolean

   Select Case iMsg
  
   Case WM_DRAWITEM
      CopyMemory tDis, ByVal lParam, Len(tDis)
      pDrawButton tDis
      ISubclass_WindowProc = 1
            
   Case WM_SIZE
      Resize
            
   Case WM_CTLCOLORSCROLLBAR
      'Debug.Print "WM_CTLCOLORSCROLLBAR"
      If (wParam = m_hWndControl) Then
         'Debug.Print "WM_CTLCOLORSCROLLBAR"
         ISubclass_WindowProc = GetSysColorBrush(SystemColorConstants.vbWindowBackground And &H1F)
      End If
         
   Case WM_VSCROLL, WM_HSCROLL
                             
      lBar = SB_CTL
      
      lScrollcode = (wParam And &HFFFF&)
      Select Case lScrollcode
      Case SB_THUMBTRACK
         ' Is vertical/horizontal?
         pGetSI tSI, SIF_TRACKPOS
         If tSI.nTrackPos > Max Then
            Value = Max
         ElseIf tSI.nTrackPos < Min Then
            Value = Min
         Else
            Value = tSI.nTrackPos
         End If
         pRaiseEvent True
         DrawTip m_eScrollType, Value
         
      Case SB_LEFT, SB_BOTTOM
         Value = Min
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_RIGHT, SB_TOP
         Value = Max
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_LINELEFT, SB_LINEUP
         'Debug.Print "Line"
         lV = Value
         lSC = m_lSmallChange
         If (lV - lSC < Min) Then
            Value = Min
         Else
            Value = lV - lSC
         End If
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_LINERIGHT, SB_LINEDOWN
          'Debug.Print "Line"
         lV = Value
         lSC = m_lSmallChange
         If (lV + lSC > Max) Then
            Value = Max
         Else
            Value = lV + lSC
         End If
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_PAGELEFT, SB_PAGEUP
         If Value - LargeChange < Min Then
            Value = Min
         Else
            Value = Value - LargeChange
         End If
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_PAGERIGHT, SB_PAGEDOWN
         If Value + LargeChange > Max Then
            Value = Max
         Else
            Value = Value + LargeChange
         End If
         pRaiseEvent False
         DrawTip m_eScrollType, Value
         
      Case SB_ENDSCROLL
         picTip.Visible = False
         pRaiseEvent False
           
      End Select
                       
   End Select
     
End Function

Public Sub HideTip()
   picTip.Visible = False
End Sub

Private Sub DrawTip(ByRef eType As ESBCScrollTypes, ByVal lValue As Long)
   Dim pt As POINTAPI
   Dim lHeight As Long
   Dim lWidth As Long
   Dim x As Long
   Dim y As Long
   Dim lDiv As Long
      
   lDiv = Max - Min
   Label1.Caption = CStr(lValue)
   
   If lDiv <= 0 Then lDiv = 1
      
   If (eType = esbcHorizontal) Then
            
      'Determine working height
      lWidth = (UserControl.ScaleWidth - 510)
               
      x = UserControl.ScaleX((lWidth * (lValue / lDiv)), vbTwips, vbPixels)
      pt.x = x + UserControl.ScaleX((600 - (600 * (lValue / lDiv))), vbTwips, vbPixels)
      pt.y = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels) + 6
                
   ElseIf (eType = esbcVertical) Then
            
      'Determine working height
      lHeight = (UserControl.ScaleHeight - 510)
               
      y = UserControl.ScaleY((lHeight * (lValue / lDiv)), vbTwips, vbPixels)
      pt.y = y + UserControl.ScaleY((225 - (225 * (lValue / lDiv))), vbTwips, vbPixels)
      pt.x = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels) + 4
                   
   End If
     
   Call ClientToScreen(UserControl.hwnd, pt)
   
   If picTip.Visible Then
      MoveTip pt.x, pt.y
   Else
      ShowTip pt.x, pt.y
      picTip.Visible = True
      MoveTip pt.x, pt.y
   End If
     
End Sub

Private Sub MoveTip(ByVal Left As Long, ByVal Top As Long)
         
   If Not picTip.Visible Then Exit Sub
     
   MoveWindow picTip.hwnd, Left, Top, picTip.Width \ Screen.TwipsPerPixelX, picTip.Height \ Screen.TwipsPerPixelY, -1
       
End Sub

Private Sub ShowTip(ByVal x As Long, ByVal y As Long)
   Dim tP As POINTAPI
   Dim hWndDesktop As Long
   Dim lStyle As Long
   Dim lhWnd As Long
   Dim lParenthWNd As Long
      
   ' Make sure the picture box won't appear in the
   ' task bar by making it into a Tool Window:
   lhWnd = picTip.hwnd
   lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_TOOLWINDOW
   lStyle = lStyle And Not (WS_EX_APPWINDOW)
   
   SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
   
   ' Determine where to show it in Screen coordinates:
   tP.x = x \ Screen.TwipsPerPixelX: tP.y = y \ Screen.TwipsPerPixelY
   lParenthWNd = UserControl.hwnd
            
   ' Make the picture box a child of the desktop (so
   ' it can be fully shown even if it extends beyond
   ' the form boundaries):
   SetParent lhWnd, hWndDesktop
         
   ' Show the form:
   SetWindowPos lhWnd, -1, tP.x, tP.y, picTip.Width \ Screen.TwipsPerPixelX, picTip.Height \ Screen.TwipsPerPixelY, SWP_NOMOVE Or SWP_NOSIZE
          
End Sub

Private Sub UserControl_Initialize()
   m_bNoFlatScrollBars = True
End Sub

Private Sub UserControl_LostFocus()
   Debug.Print "LostFocus"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   If (UserControl.Ambient.UserMode) Then
      m_hWndControl = UserControl.hwnd
      AttachMessage Me, m_hWndControl, WM_DRAWITEM
      AttachMessage Me, m_hWndControl, WM_CTLCOLORSCROLLBAR
      m_hWndParent = UserControl.Extender.Container.hwnd
      AttachMessage Me, m_hWndParent, WM_SIZE
      
   End If
   ScrollType = PropBag.ReadProperty("ScrollType", esbcHorizontal)
   Visible = PropBag.ReadProperty("Visible", True)
   
End Sub

Private Sub UserControl_Resize()
   If (m_hWndControl <> 0) Then
      pResizeButtons
      pResizeScroll
   End If
End Sub

Private Sub UserControl_Terminate()
   If (m_hWndControl <> 0) Then
      DetachMessage Me, m_hWndControl, WM_DRAWITEM
      DetachMessage Me, m_hWndControl, WM_CTLCOLORSCROLLBAR
      DetachMessage Me, m_hWndControl, WM_SIZE
      pDestroyScrollControl
   End If
   SetParent picTip.hwnd, UserControl.hwnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "ScrollType", ScrollType, esbcHorizontal
   PropBag.WriteProperty "Visible", Visible, True
End Sub
