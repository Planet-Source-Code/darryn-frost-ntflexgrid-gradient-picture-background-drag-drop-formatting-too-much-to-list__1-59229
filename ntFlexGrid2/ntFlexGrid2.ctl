VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.UserControl ntFlexGrid2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7005
   ScaleWidth      =   8835
   ToolboxBitmap   =   "ntFlexGrid2.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8370
      Top             =   5910
   End
   Begin ntFxGd2.vbalScrollButtonCtl hScroll 
      Height          =   465
      Left            =   570
      TabIndex        =   11
      Top             =   6240
      Width           =   7875
      _extentx        =   13891
      _extenty        =   820
   End
   Begin ntFxGd2.vbalScrollButtonCtl vScroll 
      Height          =   4965
      Left            =   8040
      TabIndex        =   10
      Top             =   420
      Width           =   345
      _extentx        =   609
      _extenty        =   8758
      scrolltype      =   1
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   4695
      Left            =   5160
      ScaleHeight     =   4635
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   930
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   420
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   5310
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox cmbEdit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4590
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picCheck 
      Height          =   315
      Left            =   1950
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   3630
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picUnCheck 
      Height          =   315
      Left            =   1950
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   3990
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picCheckDis 
      Height          =   315
      Left            =   1950
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   4350
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picUnCheckDis 
      Height          =   315
      Left            =   1950
      ScaleHeight     =   255
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   4710
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox lbltotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Column Totals"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      DragIcon        =   "ntFlexGrid2.ctx":0312
      Height          =   2235
      Left            =   270
      TabIndex        =   1
      Top             =   360
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   3942
      _Version        =   393216
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483643
      HighLight       =   2
      ScrollBars      =   0
      MergeCells      =   2
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuGrid 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuGridFit 
         Caption         =   "Column Best Fit"
      End
      Begin VB.Menu mnuGridBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridSortAsc 
         Caption         =   "Sort Ascending"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGridSortDesc 
         Caption         =   "Sort Descending"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGridRefresh 
         Caption         =   "Remove Sort(s)"
      End
      Begin VB.Menu mnuGridBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridFilterBy 
         Caption         =   "Filter by Selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGridFilterExclude 
         Caption         =   "Filter Excluding Selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGridFilterRemove 
         Caption         =   "Remove Filter(s)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGridBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridEdit 
         Caption         =   "Edit Range"
      End
      Begin VB.Menu mnuGridBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridResetCols 
         Caption         =   "Reset Columns"
      End
      Begin VB.Menu mnuResetBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridReset 
         Caption         =   "Reset All"
      End
      Begin VB.Menu mnuCustBar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCustom 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTotal 
      Caption         =   ""
      Begin VB.Menu mnuTotalShow 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "ntFlexGrid2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2002
Attribute VB_Description = "ntFlexGridControl.ntFlexGrid"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Option Compare Text


Public Enum ntFgBackGroundStyle
       ntFgBsSolidColor = 0
       ntFgBsPicture = 1
       ntFgBsGradient = 2
End Enum

Public Enum ntFgForeGroundDrawMode
   ntFgFGMOpaque = vbSrcCopy
   ntFgFGMTransParent = vbSrcAnd
End Enum

Private m_FgrdDraw As ntFgForeGroundDrawMode

Public Enum ntFgBackGroundPictureMode
       ntFgBPMNormal = 0
       ntFgBPMStretch = 1
       ntFgBPMTile = 2
End Enum

Public Enum ntFgGradientType
   ntFgGTHorizontal = 0
   ntFgGTVertical = 1
End Enum

Private m_BkgdBmp As cBitMap

Private bRedrawFlag As Boolean
Private m_BackPic As StdPicture
Private m_GradClrStart As OLE_COLOR
Private m_GradClrEnd As OLE_COLOR
Private m_BackStyle As ntFgBackGroundStyle
Private m_GradType As ntFgGradientType
Private m_GradTransColor As OLE_COLOR
Private m_BackPicDrawMode As ntFgBackGroundPictureMode

Private pRect As RECT
Private m_bScrolling As Boolean


Private m_iKeyCode As Integer
Private m_iShift As Integer

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private m_bScrollType As Integer

Private prevRow As Long
Private prevRowSel As Long
Private m_blnDoSel As Boolean
Private m_blnDidSort As Boolean
Private m_blnMoving As Boolean

Private m_bln_NeedHorzScroll As Boolean
Private m_bln_NeedVertScroll As Boolean

Private Const ERR_INVCOL = vbObjectError + 5005
Private Const ERR_INVROW = vbObjectError + 5015
Private Const ERR_NORS = vbObjectError + 5025
Private Const ERR_NOHDR = vbObjectError + 5035
Private Const ERR_NOSEL = vbObjectError + 5045

Private Const MSG_INVCOL = "Invalid Column Key or Index"
Private Const MSG_INVROW = "Invalid Row Index"

Private m_lPrevCol As Long
Private m_lPrevColSel As Long
Private m_lPrevRow As Long
Private m_lPrevRowSel As Long
   
Private m_blnDragging As Boolean
Private m_lDragCol As Long
Private m_lLastDragCol As Long
Private m_lLineCol As Long
Private m_lLastLineCol As Long
Private m_arrDragColors() As Long

Private m_blnPaging As Boolean
Private m_blnChanged As Boolean
Private m_UpdateScrollValue As Long
Private m_intResizing As Integer
Private m_blnHasFocus As Boolean
Private m_blnIgnoreRCChange As Boolean

Private m_blnAllowMenu As Boolean

' API Constants
Private Const WM_ERASEBKGND = &H14
Private Const WM_LBUTTONUP = &H202
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_PAINT = &HF
Private Const WM_TIMER = &H113

Private Const VK_NEXT = &H22 'Page Down Key
Private Const VK_PRIOR = &H21 'Page Up key

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

Public Enum nfgGridModeConstants
   nfgGridmodeUnbound = 0
   nfgGridModeBound = 1
End Enum

Public Enum nfgEditType
   ntEditAdd = 0
   ntEditDelete = 1
   ntEditField = 2
End Enum

Public Enum EditKey
   [SpaceBar] = 32
   [Enter] = 13
   [Insert] = 45
   [Ctrl + E] = 69
   [F2] = 113
   [F3] = 114
   [F4] = 115
   [F5] = 116
   [F6] = 117
   [F7] = 118
   [F8] = 119
   [F9] = 120
   [F10] = 121
End Enum

'Keep track of painted area
Private m_CurrRow As Long
Private m_CurrRowSel As Long
Private m_CurrCol As Long
Private m_CurrColSel As Long

Private m_blnShown As Boolean
Private m_blnNeedsRedraw As Boolean
Private m_blnDblClickFilter As Boolean

Private m_HTimer As Long
Private m_VTimer As Long
Private m_hScroll As ntFxGdScrollBar
Private m_vScroll As ntFxGdScrollBar

Private m_blnManualScroll As Boolean
Private m_blnIgnoreSel As Boolean
Private m_blnBigSel As Boolean
Private m_ScrollDown As Boolean
Private m_ScrollLeft As Boolean

'Default Property Values:
Private Const m_def_ColWidth = 940
Private Const m_def_DisabledColor As Long = &HEEEEEE
Private Const m_def_FormatString = ""
Private Const m_def_GRID_PLUSCOLOR = &H80000008
Private Const m_def_GRID_MINUSCOLOR = &HC0&
Private Const m_def_Height = 360
Private Const m_def_MinColWidth = 0
Private Const m_def_MaxColWidth = 0
Private Const m_def_MinUpDownWidth = 150
Private Const m_def_MinComboWidth = 360
Private Const m_def_MinDTWidth = 360
Private Const m_def_MinTextBoxWidth = 360
Private Const m_def_RecordSelWidth = 200      'For RecordSelectors Only
Private Const m_def_Rowheight = 270
Private Const m_def_RowHeightFixed = 285
Private Const m_def_RowHeightMin = 0
Private Const m_def_ShowTotalHeight = 300
Private Const m_def_Text = ""
Private Const m_def_Width = 360
Private Const m_def_FocusRect = 1
Private Const m_def_Highlight = 2
Private Const m_def_TotalFloat = 0

Private m_RSMaster                  As Object
Private m_RsFiltered                As Object
Private m_ntColumns                 As ntColumns
Private m_colFilters                As ntColFilters
Private m_colColpics                As ntColPics
Private m_colRowColors              As ntRowInfo
Private m_arrVisCols()              As Long

'Property variables
Private m_eScrollBars               As nfgScrollBarSettings
Private m_blnAllowColMove           As Boolean
Private m_blnGridMode               As Boolean
Private m_sngRecordSelectorWidth    As Single
Private m_blnAllowFilter            As Boolean
Private m_blnAllowSort              As Boolean
Private m_blnAllowEdit              As Boolean
Private m_blnRedraw                 As Boolean
Private m_blnAutoSizeColumns        As Boolean
Private m_blnColorByRow             As Boolean
Private m_blnEnabled                As Boolean
Private m_sngMaxColWidth            As Single
Private m_sngMinColWidth            As Single
Private m_sngRowHeightFixed         As Single
Private m_sngRowHeight              As Single
Private m_sngRowHeightMin           As Single
Private m_sngRowHeightMax           As Single
Private m_picChecked                As StdPicture
Private m_picUnchecked              As StdPicture
Private m_picCheckedDis             As StdPicture
Private m_picUncheckedDis           As StdPicture
Private m_blnHeaderRow              As Boolean
Private m_blnRecordSelectors        As Boolean
Private m_blnUseFieldNamesAsHeader  As Boolean
Private m_Text                      As String
Private m_intFocusRect              As Integer
Private m_intHighlight              As Integer
Private m_blnTotalRow               As Boolean
Private m_clrGridTotalPlus          As OLE_COLOR
Private m_clrGridTotalMinus         As OLE_COLOR
Private m_clrEnabledBackcolor       As OLE_COLOR
Private m_clrDisBackcolor           As OLE_COLOR
Private m_clrPositive               As OLE_COLOR
Private m_clrNegative               As OLE_COLOR
Private m_intTotalFloat             As Integer
Private m_intEditKey                As Integer

'Working variables
Private m_lonEditCol                As Long  ' Set to the current Col being edited
Private m_lonEditRow                As Long  ' Set to the current Row being edited
Private m_lonEditRowSel             As Long  ' Set to the current RowSel being edited
Private m_lonEditRowIDs()           As Long  ' Array Holding Row(s) being edited
Private m_LastCol                   As Integer     ' Set to Col Index on Col Change, check to see if Col actually changed
Private m_LastRow                   As Long
Private m_blnLoading                As Boolean
Private m_PrevEditVal               As Variant     ' Store previous value for cancel edit operation
Private m_blnCancelEdit             As Boolean
Private m_blnInitEdit               As Boolean
Private m_blnLButtonDown            As Boolean     ' Used for subclassing
Private m_blnResizing               As Boolean     ' Used for subclassing
Private m_blnGridSubclassed         As Boolean
Private m_blnDidMenu                As Boolean
Private m_blnLeftClick              As Boolean
Private m_PrevVertValue             As Long
Private m_PrevHorzValue             As Long
Private m_lonRows                   As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'**
'The AfterEdit() event is raised after the changes have been made to the underlying recordset.
'Public Event AfterEdit(ByVal pEditType As nfgEditType, ByVal vNewValue As Variant, ByVal pColName As String, ByRef arrlonRowID() As Long)
'Event AfterEdit(ByVal pEditType As nfgEditType, ByVal vNewValue As Variant, ByVal pColName As String, arrlonRowID() As Long)
Event AfterEdit(ByVal pEditType As nfgEditType, ByVal vNewValue As Variant, ByVal pColName As String, arrlonRowID() As Long)
Attribute AfterEdit.VB_HelpID = 2340

'**
'The BeforeEdit() event is raised when the edit control is about to lose focus, either by clicking somewhere else in the grid, or pressing Enter to accept the edit, but prior to the changes being made final in the underlying recordset itself.
Public Event BeforeEdit(ByVal pEditType As nfgEditType, ByVal vCurrentValue As Variant, ByRef vNewValue As Variant, ByVal pColName As String, ByRef arrlonRowID() As Long, ByVal bValid As Boolean, ByRef pCancel As Boolean)
Attribute BeforeEdit.VB_HelpID = 2360
'**
'The BeforeShowMenu() event is raised prior to showing the right mouse menu, so any validation or settings applicable to the custom menu can be modified if necessary.
Public Event BeforeShowMenu()
Attribute BeforeShowMenu.VB_HelpID = 2370
'**
'The Click() event is raised whenever the mouse button is pressed and released over any portion of the grid.
Public Event Click() 'MappingInfo=Grid1,Grid1,-1,Click
Attribute Click.VB_HelpID = 2380
Attribute Click.VB_UserMemId = -600
'**
'The ColChange() event is raised whenever the selected column in the grid control changes.
'The ntCol Parameter returns as reference to the ntColumn Object.
Public Event ColChange(ByVal ntCol As ntColumn) 'MappingInfo=UserControl1,UserControl1,-1,ColChange
Attribute ColChange.VB_HelpID = 1390
'**
'The CustomMenuClick() event is raised whenever a custom menu item that has been added programmatically to the grid is clicked by the user.
Public Event CustomMenuClick(ByVal sTag As String, ByVal index As Integer)
Attribute CustomMenuClick.VB_HelpID = 2400
'**
'The DblClick() event is raised whenever the mouse button is pressed and released twice quickly over any portion of the grid.
Public Event DblClick() 'MappingInfo=Grid1,Grid1,-1,DblClick
Attribute DblClick.VB_HelpID = 2410
'**
'The EditControlValidate() event is raised whenever an edit procedure is taking place in the grid and the edit control itself is about to lose focus.
Public Event EditControlValidate(ByRef pText As String, ByVal pColName As String, ByRef Cancel As Boolean)
Attribute EditControlValidate.VB_HelpID = 2420
'**
'The EnterCell() event is raised whenever a cell in the grid gets the focus.
Public Event EnterCell() 'MappingInfo=Grid1,Grid1,-1,EnterCell
Attribute EnterCell.VB_HelpID = 2430
'**
'The KeyDown() event is raised whenever the user presses a key while the grid has the focus.
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Grid1,Grid1,-1,KeyDown
Attribute KeyDown.VB_HelpID = 2440
'**
'The KeyPress() event is raised whenever the user presses a valid Ascii key while the grid has the focus.
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=Grid1,Grid1,-1,KeyPress
Attribute KeyPress.VB_HelpID = 2450
'**
'The KeyUp() event is raised whenever the user release a key while the grid has the focus.
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Grid1,Grid1,-1,KeyUp
Attribute KeyUp.VB_HelpID = 2460
'**
'The LeaveCell() event is raised whenever the focus leaves a cell in the grid.
Public Event LeaveCell() 'MappingInfo=Grid1,Grid1,-1,LeaveCell
Attribute LeaveCell.VB_HelpID = 2470
'**
'The MouseDown() event is raised whenever the user presses a button while the cursor is over the grid.
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Grid1,Grid1,-1,MouseDown
Attribute MouseDown.VB_HelpID = 2480
'**
'The MouseMove() event is raised whenever the user moves the mouse while the cursor is over the grid.
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Grid1,Grid1,-1,MouseMove
Attribute MouseMove.VB_HelpID = 2490
'The MouseMove() event is raised whenever the user moves the mouse while the cursor is over the grid.
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Grid1,Grid1,-1,MouseUp
Attribute MouseUp.VB_HelpID = 2500

'The BeforeFilter() event is raised whenever the user changes the FormattedFilter Property of the grid, either by double-clicking on a cell, or selecting a cell(or range of cells) and using the right-click menu to filter on the cells contents.
Public Event BeforeFilter(ByRef nFilter As ntFilter, ByRef bCancel As Boolean)
Attribute BeforeFilter.VB_HelpID = 2510

Public Event OnFilter(ByVal nFilter As ntFilter)
Attribute OnFilter.VB_HelpID = 2520

'The OnFilterRemove() event is raised whenever the FormattedFilter Property of the grid is reset.
Public Event OnFilterRemove()
Attribute OnFilterRemove.VB_HelpID = 2530

'The Resize() event is raised whenever the grid control changes size.
Public Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_HelpID = 2540

'The RowChange() event is raised whenever the selected row in the grid control changes.
Public Event RowChange(ByVal lRow As Long)  'MappingInfo=UserControl1,UserControl1,-1,RowChange
Attribute RowChange.VB_HelpID = 2550
'The Scroll() event is raised whenever the grid is scrolled.
Public Event Scroll() 'MappingInfo=Grid1,Grid1,-1,Scroll
Attribute Scroll.VB_HelpID = 2560
Public Event BeforeSort(ByVal sBeginCol As String, ByVal sEndCol As String, ByRef bCancel As Boolean)
Attribute BeforeSort.VB_HelpID = 2570
'The Sort() event is raised whenever the grid is sorted.
Public Event OnSort(ByVal sBeginCol As String, ByVal sEndCol As String)
Attribute OnSort.VB_HelpID = 2580
'The Sort() event is raised whenever the grid is sorted.
Public Event OnSortRemove()
Attribute OnSortRemove.VB_HelpID = 2590
'The BeginColDrag() event is raised whenever a column drag is started
Public Event ColBeginDrag(ByVal sDragColName As String, ByRef pCancel As Boolean)
Attribute ColBeginDrag.VB_HelpID = 1600
'The DragOver() event is raised whenever a drag operation is crossing a column
Public Event ColDragOver(ByVal sDragColName As String, ByVal lNewColPos As Long, ByRef pCancel As Boolean)
Attribute ColDragOver.VB_HelpID = 2610
'The DragDrop() event is raised whenever the user lets up on the mouse while a ColDrag operation is pending
Public Event ColDragDrop(ByVal sDragColName As String, ByVal lNewColPos As Long, ByRef pCancel As Boolean)
Attribute ColDragDrop.VB_HelpID = 2620

Public Event ColEndDrag(ByVal sDragColName As String, ByVal lNewColPos As Long)
Attribute ColEndDrag.VB_HelpID = 2630

Public Event OnColReorder()
Attribute OnColReorder.VB_HelpID = 2640

Public Event SelChange()
Attribute SelChange.VB_HelpID = 2650

Public Event hScrollButtonClick(ByVal lButton As Long)
Attribute hScrollButtonClick.VB_HelpID = 2660
Public Event vScrollButtonClick(ByVal lButton As Long)
Attribute vScrollButtonClick.VB_HelpID = 2670

Implements ISubclass

Private Sub Grid1_DragDrop(Source As Control, x As Single, y As Single)
      
   Picture2.Visible = False
   m_blnDragging = False
   Grid1.Drag vbEndDrag
   
   If Not m_blnAllowColMove Then Exit Sub
   If m_lLineCol < Grid1.FixedCols Then Exit Sub
   If m_lLineCol = m_lDragCol Or m_lLineCol = m_lDragCol + 1 Then Exit Sub
   
On Error Resume Next

   Dim bCancel As Boolean
   
   If m_lDragCol > m_lLineCol Then
      RaiseEvent ColDragDrop(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, m_lLineCol - Grid1.FixedCols, bCancel)
      If Not bCancel Then
         Call MoveCol(m_lDragCol - Grid1.FixedCols, m_lLineCol - Grid1.FixedCols)
      End If
   Else
      RaiseEvent ColDragDrop(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, (m_lLineCol - Grid1.FixedCols) - 1, bCancel)
      If Not bCancel Then
         Call MoveCol(m_lDragCol - Grid1.FixedCols, (m_lLineCol - Grid1.FixedCols) - 1)
      End If
   End If
   
End Sub

Private Sub Grid1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   Dim bCancel As Boolean
   Dim bLeft As Boolean
   Dim bCurrCol As Long
   
   SetDragLine x
   
   bCurrCol = Grid1.MouseCol
   
   bLeft = x < (Grid1.ColPos(bCurrCol) + (Grid1.ColWidth(bCurrCol) / 2))
      
   If bCurrCol = m_lDragCol Then
      RaiseEvent ColDragOver(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, bCurrCol - Grid1.FixedCols, bCancel)
   Else
      If (bCurrCol = (m_lDragCol + 1)) And bLeft Or (bCurrCol = (m_lDragCol - 1)) And Not bLeft Then
         RaiseEvent ColDragOver(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, m_lDragCol - Grid1.FixedCols, bCancel)
      Else
         If m_lDragCol > m_lLineCol Then
            RaiseEvent ColDragOver(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, m_lLineCol - Grid1.FixedCols, bCancel)
         Else
            RaiseEvent ColDragOver(m_ntColumns(m_lDragCol - Grid1.FixedCols).Name, (m_lLineCol - Grid1.FixedCols) - 1, bCancel)
         End If
      End If
   End If
      
   If bCancel Then
      Picture2.Visible = False
      m_blnDragging = False
      Grid1.Drag vbEndDrag
   End If
  
End Sub

Private Sub SetDragLine(ByVal x As Single)
   Dim bLeft As Boolean
   
   bLeft = x < (Grid1.ColPos(Grid1.MouseCol) + (Grid1.ColWidth(Grid1.MouseCol) / 2))
   If bLeft Then
      m_lLineCol = Grid1.MouseCol
   Else
      m_lLineCol = Grid1.MouseCol + 1
   End If
   
   If m_lLineCol <> m_lLastLineCol Then
      Picture2.Visible = True
      Picture2.ZOrder 0
      If bLeft Then
         Picture2.Left = Grid1.ColPos(Grid1.MouseCol) - 45
      Else
         Picture2.Left = (Grid1.ColPos(Grid1.MouseCol) + Grid1.ColWidth(Grid1.MouseCol)) - 45
      End If
      Picture2.Width = 90
      Picture2.Top = -30
      Picture2.Height = (Grid1.RowPos(Grid1.Rows - 1) + Grid1.RowHeight(Grid1.Rows - 1) + 60)
   End If
   
   m_lLastLineCol = m_lLineCol

End Sub

Private Sub Grid1_LostFocus()
   m_blnHasFocus = False
End Sub

Private Sub hScroll_ButtonClick(ByVal lButton As Long)
   If Ambient.UserMode And m_blnShown Then
      If bEditing Then Grid1_GotFocus
   End If
   RaiseEvent hScrollButtonClick(lButton)
End Sub

Private Sub mnuGridResetCols_Click()
   Me.ResetColumns
End Sub


Private Sub Timer1_Timer()
   Timer1.Enabled = False
   If m_bScrollType = 0 Then
      Vert_Scroll
   Else
      Horz_Scroll
   End If
End Sub

Private Sub UserControl_GotFocus()
   Grid1.SetFocus
End Sub

Private Sub vScroll_ButtonClick(ByVal lButton As Long)
   If Ambient.UserMode And m_blnShown Then
      If bEditing Then Grid1_GotFocus
   End If
   RaiseEvent vScrollButtonClick(lButton)
End Sub

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim fEat As Boolean
   Dim LOWORD As Long
   Dim HIWORD As Long
   Dim blnSetCols As Boolean
   Dim blnSetRows As Boolean
   Static numtimes As Long
   Dim cGridBmp As cBitMap
   Dim cBufferBmp As cBitMap
   
On Error GoTo SubClass_Error

   blnSetCols = False
   blnSetRows = False

  If hwnd = UserControl.hwnd Then

      Select Case iMsg

         Case WM_TIMER
                      
            Call ScrollTimer(wParam)
      
      End Select

   Else

      Select Case iMsg
         
         Case WM_PAINT
                        
            'Exit here, since paint gets called after scrolling also
            If m_bScrolling Then
               ISubclass_WindowProc = 0
               Exit Function
            End If
               
            'If the background is Solid Color, do nothing, pass it on
            If (m_BackStyle = ntFgBsSolidColor) Then
               ISubclass_WindowProc = MSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
               Exit Function
            End If
                         
            Dim PS As PAINTSTRUCT
            Dim GridPicDC As Long, GridPicBMP As Long
            Dim BufferDC As Long, BufferBMP As Long
            Dim gridDC As Long
            Dim pRect As RECT
              
            GetClientRect UserControl.hwnd, pRect
              
            'Add 2 pixels to account for border area when no scroll bars
            pRect.Right = pRect.Right + 2
            pRect.Bottom = pRect.Bottom + 2
                                             
            'Start painting control ...
            Graphics.BeginPaint hwnd, PS
              
            gridDC = PS.HDC 'Get the Grid DC
            
            Set cGridBmp = Graphics.BitmapFromHDC(gridDC, pRect.Right, pRect.Bottom)
            Set cBufferBmp = Graphics.BitmapFromHDC(gridDC, pRect.Right, pRect.Bottom)
          
            'this is the big thing ! We are sending WM_PAINT to our backbuffer
            MSubclass.CallOldWindowProc hwnd, iMsg, ByVal cGridBmp.HDC, 0&
              
            'Our MsFlexGrid at this point, is painted into cGridBmp
                       
            'Only do this part when necessary
            'bRedraw indicates a need to redraw background due to size change, picture change, etc.
            If ((bRedrawFlag) Or (m_BkgdBmp Is Nothing)) Then
               bRedrawFlag = False
               If m_BackStyle = ntFgBsGradient Then
                  'Draws a Gradient on the m_BkgdBmp HDC
                  CreateGradientBackground pRect.Right, pRect.Bottom
                ElseIf m_BackStyle = ntFgBsPicture Then
                  CreatePicBkgd
               End If
            End If
                            
           
            'This happens every time
            If m_BackStyle = ntFgBsGradient Then
               'Copy the previously created background to our buffer
               Graphics.Blit cBufferBmp.HDC, m_BkgdBmp, vbSrcCopy
            ElseIf m_BackStyle = ntFgBsPicture Then
               'Copy the previously created grid to our buffer
               DrawPicBackground cBufferBmp, cGridBmp
            End If
                                    
            'And now, we can overlay the pic of the grid on our background.
            'We use the BkgdMaskColor property to determine which pixels are transparent
            If m_FgrdDraw = ntFgFGMOpaque Then
               Graphics.BlitTransparent cBufferBmp.HDC, cGridBmp, m_GradTransColor
            Else
               Graphics.Blit cBufferBmp.HDC, cGridBmp, vbSrcAnd
            End If
            
            'We have all the changes into backbuffer. Let's bring in back to MsFlexGrid.hDc
            With PS.rcPaint
               Graphics.BlitHDC gridDC, .Left, .Top, .Right - .Left, .Bottom - .Top, cBufferBmp.HDC, .Left, .Top, vbSrcCopy
            End With
             
           Set cBufferBmp = Nothing
           Set cGridBmp = Nothing
            
           Graphics.EndPaint hwnd, PS
            
           ISubclass_WindowProc = 0 'When a function intercepts WM_PAINT it must return 0
            
         Case WM_LBUTTONUP
           
            If m_HTimer <> 0 Then
               Call KillTimer(UserControl.hwnd, 2)
               m_HTimer = 0
            End If
            If m_VTimer <> 0 Then
               Call KillTimer(UserControl.hwnd, 1)
               m_VTimer = 0
            End If
         
         Case WM_ERASEBKGND
                     
            If Grid1.MouseCol >= Grid1.FixedCols Then
               If Grid1.MouseRow < Grid1.FixedRows And m_blnHeaderRow Then
                  If (AllowUserResizing = nfgResizeBoth) Or (AllowUserResizing = nfgResizeColumns) Then
                     WriteColWidths
                  End If
               End If
            Else
               If m_blnRecordSelectors Then
                  If (AllowUserResizing = nfgResizeBoth) Or (AllowUserResizing = nfgResizeRows) Then
                     WriteRowHeights
                  End If
               End If
            End If
            
       End Select

   End If
         
Exit Function

SubClass_Error:

  Exit Function

End Function

Private Sub CreatePicBkgd()
   If m_BackPic Is Nothing Then Exit Sub
   
   If (Not m_BkgdBmp Is Nothing) Then Set m_BkgdBmp = Nothing
   
   'Create the background bitmap actual grid client size
   Set m_BkgdBmp = Graphics.BitmapFromPicture(m_BackPic)
   
End Sub

Private Sub DrawPicBackground(ByRef cBuffer As cBitMap, ByVal cGridBmp As cBitMap)
   
   Dim x As Long
   Dim y As Long
   Dim lDrawWidth As Long
   Dim lDrawHeight As Long
   Dim bSmaller As Boolean
   
   Dim cMode As ntFgBackGroundPictureMode
   
   cMode = m_BackPicDrawMode
   
   'If the selection is tile, but the picture is bigger than the grid, use normal style
   If ((cMode = ntFgBPMTile) And (m_BkgdBmp.Width >= cBuffer.Width) And (m_BkgdBmp.Height >= cBuffer.Height)) Then
      cMode = ntFgBPMNormal
   End If
      
   Select Case cMode

      'Center the picture
      Case ntFgBPMNormal
         bSmaller = False
         If m_BkgdBmp.Width >= cBuffer.Width Then
            x = 0
            lDrawWidth = cBuffer.Width
         Else
            x = ((cBuffer.Width - m_BkgdBmp.Width) / 2)
            lDrawWidth = m_BkgdBmp.Width
            bSmaller = True
         End If
         If m_BkgdBmp.Height >= cBuffer.Height Then
            y = 0
            lDrawHeight = cBuffer.Height
         Else
            y = ((cBuffer.Height - m_BkgdBmp.Height) / 2)
            lDrawHeight = m_BkgdBmp.Height
            bSmaller = True
         End If
               
         'If the picture is smaller than the grid background, copy the grid to the background first
         'If the picture is bigger, don't waste time as the picture will cover the entire Background
         If bSmaller Then Graphics.Blit cBuffer.HDC, cGridBmp, vbSrcCopy
                     
         'Draw the Picture over the top
         Graphics.BlitHDC cBuffer.HDC, x, y, lDrawWidth, lDrawHeight, m_BkgdBmp.HDC, 0, 0, vbSrcCopy
                        
      Case ntFgBPMStretch
         'Don't need to copy grid since picture will cover background
         Graphics.BlitStretch cBuffer.HDC, cBuffer.Width, cBuffer.Height, m_BkgdBmp, vbSrcCopy
      
      Case ntFgBPMTile

         Graphics.TilePicture cBuffer, m_BackPic
   
   End Select

End Sub

Private Sub CreateGradientBackground(ByVal lwidth As Long, ByVal lheight As Long)
   
   If (Not m_BkgdBmp Is Nothing) Then Set m_BkgdBmp = Nothing
   
   Set m_BkgdBmp = Graphics.BitmapFromHDC(GetDC(Grid1.hwnd), lwidth, lheight)
                
   Graphics.DrawGradientRect m_BkgdBmp, m_GradClrStart, m_GradClrEnd, m_GradType

End Sub

Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function


Private Function GetHiLoWord(lParam As Long, LOWORD As Long, HIWORD As Long) As Long

    ' This is the LOWORD of the lParam:
    LOWORD = lParam And &HFFFF&
    ' LOWORD now equals 65,535 or &HFFFF
    
    ' This is the HIWORD of the lParam:
    HIWORD = lParam \ &H10000 And &HFFFF&
    ' HIWORD now equals 30,583 or &H7777
    
    GetHiLoWord = 1

End Function

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If MSubclass.CurrentMessage = WM_PAINT Then
      ISubclass_MsgResponse = emrConsume
   Else
      ISubclass_MsgResponse = emrPostProcess
   End If
End Property

Private Function IsMouseOutOfGrid() As Boolean
   Dim lRet As Long
   Dim pt As POINTAPI

   IsMouseOutOfGrid = False

   lRet = GetCursorPos(pt)
   lRet = ScreenToClient(Grid1.hwnd, pt)
 
   If (Screen.TwipsPerPixelX * pt.x) > (Grid1.ColPos(m_arrVisCols(UBound(m_arrVisCols)) + Grid1.FixedCols) + Grid1.ColWidth(m_arrVisCols(UBound(m_arrVisCols)) + Grid1.FixedCols)) Or _
      (Screen.TwipsPerPixelY * pt.y) > (Grid1.RowPos(Grid1.Rows - 1) + Grid1.RowHeight(Grid1.Rows - 1)) Then
      IsMouseOutOfGrid = True
   End If

End Function
'***********************************************************************************************************
' PROPERTIES **************************************************************************************
'*********************************************************************************************************

Public Property Get AllowDblClickFilter() As Boolean
Attribute AllowDblClickFilter.VB_HelpID = 1280
  AllowDblClickFilter = m_blnDblClickFilter
End Property

'**
'Returns or sets whether the grid will respond to user-generated events.
'@param        bValue Boolean. Required.
Public Property Let AllowDblClickFilter(ByVal bValue As Boolean)
   m_blnDblClickFilter = bValue
   PropertyChanged "DblClickFilter"
End Property

Public Property Get AllowEdit() As Boolean
Attribute AllowEdit.VB_HelpID = 1290
   AllowEdit = m_blnAllowEdit
End Property

'**
'Returns or sets the value determining whether the grid will be editable or not.
'Operates independently of the Column Enabled property.
'@param        bValue Boolean. Required.
'@rem Note: each column to be edited must be enabled seperately.
Public Property Let AllowEdit(ByVal bValue As Boolean)
   m_blnAllowEdit = bValue
   PropertyChanged "AllowEdit"
   If Ambient.UserMode Then
      If m_blnShown Then
         If Not IsUnbound Then
            If HasRecords Then
               If bValue = False Then
                  If bEditing Then Grid1_GotFocus
               End If
            End If
         End If
      End If
   End If
End Property

Public Property Get AllowFilter() As Boolean
Attribute AllowFilter.VB_HelpID = 1300
  AllowFilter = m_blnAllowFilter
End Property
'**
'Returns or sets the value determining whether the user will be able to
'filter the grid by double-clicking a selection,
'or using the menu to filter on a cell or range of cells.
'@param        bValue Boolean. Required.
Public Property Let AllowFilter(ByVal bValue As Boolean)
   m_blnAllowFilter = bValue
   PropertyChanged "AllowFilter"
   If Ambient.UserMode Then
      If m_blnShown Then
         If Not IsUnbound Then
            If HasRecords Then
               If bEditing Then Grid1_GotFocus
               If bValue = False Then
                  If IsFiltered Then Call RemoveFormattedFilter
               End If
            End If
         End If
      End If
   End If
End Property

Public Property Let AllowMenu(ByVal bValue As Boolean)
   m_blnAllowMenu = bValue
   PropertyChanged "AllowMenu"
End Property

Public Property Get AllowMenu() As Boolean
   AllowMenu = m_blnAllowMenu
End Property

Public Property Let AllowMoveCols(ByVal bAllowMove As Boolean)
Attribute AllowMoveCols.VB_HelpID = 1310
   m_blnAllowColMove = bAllowMove
   PropertyChanged "AllowMoveCols"
End Property

Public Property Get AllowMoveCols() As Boolean
   AllowMoveCols = m_blnAllowColMove
End Property

Public Property Get AllowSort() As Boolean
Attribute AllowSort.VB_HelpID = 1320
  AllowSort = m_blnAllowSort
End Property
'**
'Returns or sets the value determining whether the user will be able to sort
'the grid by double-clicking a column header,
'or using the menu to sort a column or range of columns.
'@param        bValue Boolean. Required.
Public Property Let AllowSort(ByVal bValue As Boolean)
   m_blnAllowSort = bValue
   PropertyChanged "AllowSort"
   If Ambient.UserMode Then
      If m_blnShown Then
         If Not IsUnbound Then
            If HasRecords Then
               If bEditing Then Grid1_GotFocus
               If bValue = False Then
                  If m_RsFiltered.Sort <> m_RSMaster.Sort Then Call RemoveSort
               End If
            End If
         End If
      End If
   End If
End Property

Public Property Get AllowUserResizing() As nfgAllowUserResizeSettings
Attribute AllowUserResizing.VB_HelpID = 1330
  AllowUserResizing = Grid1.AllowUserResizing
End Property
'**
'Returns or sets the value determining whether the user will be able to
'adjust column widths and row heights by dragging with the mouse.
'@param        fgAllowResizing Integer. Required. One of the members of the nfgAllowUserResizeSettings enumeration.
Public Property Let AllowUserResizing(ByVal fgAllowResizing As nfgAllowUserResizeSettings)
   Grid1.AllowUserResizing() = fgAllowResizing
   PropertyChanged "AllowUserResizing"
End Property

Public Property Get Appearance() As nfgAppearanceSettings
Attribute Appearance.VB_HelpID = 1340
   Appearance = Grid1.Appearance
End Property
'**
'Returns or sets the value determining whether the grid control will have a flat or 3-D appearance.
'Available only at design-time.
'@param        vbAppearance Integer. Required. One of the members of the nfgAppearanceSettings enumeration.
Public Property Let Appearance(ByVal new_Appearance As nfgAppearanceSettings)
   'If Ambient.UserMode Then Err.Raise 387
   Grid1.Appearance() = new_Appearance
   PropertyChanged "Appearance"
End Property

Public Property Get AutoSizeColumns() As Boolean
Attribute AutoSizeColumns.VB_HelpID = 1350
   AutoSizeColumns = m_blnAutoSizeColumns
End Property

'**
'Returns or sets the value determining whether the grid will autosize all the columns to best fit the displayed text.
'@param        bValue Boolean. Required.
'@rem Note: if the displayed width of all visible columns is less than the width of the grid
'itself, the columns will be widened proportionately to approximate the width of the grid.
Public Property Let AutoSizeColumns(ByVal bValue As Boolean)
   m_blnAutoSizeColumns = bValue
   PropertyChanged "AutoSizeColumns"
   If Ambient.UserMode Then
      If m_blnShown Then
         If Not IsUnbound Then
            If HasRecords Then
               If bEditing Then Grid1_GotFocus
            End If
            If m_blnAutoSizeColumns = True Then
               GridMousePointer = vbHourglass
               Screen.MousePointer = vbHourglass
               Grid1.Redraw = False
               Call AutosizeGridColumns
               Call RecalcGrid
               GridMousePointer = vbDefault
               Screen.MousePointer = vbDefault
               Grid1.Redraw = m_blnRedraw
            End If
         End If
      End If
   End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_HelpID = 1360
  BackColor = Grid1.BackColor
End Property
'**
'Returns or sets the value for the backcolor of the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let BackColor(ByVal cColor As OLE_COLOR)
  Grid1.BackColor = cColor
  Grid1.Refresh
  m_clrEnabledBackcolor = cColor
  PropertyChanged "BackColor"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_HelpID = 1370
  BackColorBkg = Grid1.BackColorBkg
End Property
'**
'Returns or sets the value for the backcolor of any unpopulated portion of the grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let BackColorBkg(ByVal cColor As OLE_COLOR)
  Grid1.BackColorBkg() = cColor
  PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorDisabled() As OLE_COLOR
Attribute BackColorDisabled.VB_HelpID = 1380
  BackColorDisabled = m_clrDisBackcolor
End Property
'**
'Returns or sets the value for the backcolor of the Grid when it is disabled.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let BackColorDisabled(ByVal cColor As OLE_COLOR)
  m_clrDisBackcolor = cColor
  PropertyChanged "BackColorDisabled"
End Property

Public Property Get BackColorHeader() As OLE_COLOR
Attribute BackColorHeader.VB_HelpID = 1390
  BackColorHeader = Grid1.BackColorFixed
End Property
'**
'Returns or sets the value for the backcolor of the header row in the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
'@rem Note: ShowHeaderRow must be true to see the header row.
Public Property Let BackColorHeader(ByVal cColor As OLE_COLOR)
  Grid1.BackColorFixed() = cColor
  PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_HelpID = 1400
  BackColorSel = Grid1.BackColorSel
End Property
'**
'Returns or sets the value for the Backcolor of any selected portions of the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let BackColorSel(ByVal cColor As OLE_COLOR)
  Grid1.BackColorSel() = cColor
  PropertyChanged "BackColorSel"
End Property

Public Property Get BackStyle() As ntFgBackGroundStyle
   BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal Value As ntFgBackGroundStyle)
   m_BackStyle = Value
   bRedrawFlag = True
   PropertyChanged "BackStyle"
End Property

Public Property Get BkgdMaskColor() As OLE_COLOR
   BkgdMaskColor = m_GradTransColor
End Property

Public Property Let BkgdMaskColor(ByVal Value As OLE_COLOR)
   m_GradTransColor = Value
   bRedrawFlag = True
   PropertyChanged "BkgdMaskColor"
End Property

Public Property Set BkgdPicture(ByVal cPicture As StdPicture)
   Set m_BackPic = cPicture

   bRedrawFlag = True
   PropertyChanged "BkgdPicture"
End Property

Public Property Get BkgdPicture() As StdPicture
   Set BkgdPicture = m_BackPic
End Property

Public Property Let BkgdPictureDrawMode(ByVal Value As ntFgBackGroundPictureMode)
   m_BackPicDrawMode = Value
   bRedrawFlag = True
   PropertyChanged "BkgdPictureDrawMode"
End Property

Public Property Get BkgdPictureDrawMode() As ntFgBackGroundPictureMode
   BkgdPictureDrawMode = m_BackPicDrawMode
   bRedrawFlag = True
End Property

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_HelpID = 1410
Attribute Bookmark.VB_MemberFlags = "400"
   
   If Not Ambient.UserMode Then Err.Raise 387
   
   Bookmark = -1
   
   If IsUnbound Or Not HasRecords Then Exit Property
      
   If m_CurrRow < 0 Or m_CurrRow + 1 > m_RsFiltered.RecordCount Then Exit Property
   
   If Not m_RsFiltered.AbsolutePosition = m_CurrRow + 1 Then
      m_RsFiltered.AbsolutePosition = m_CurrRow + 1
   End If
   Bookmark = m_RsFiltered.Bookmark
   
End Property

'**
'Returns or Sets the Current ADO Recordset Bookmark for the Grid.
'If setting a bookmark to a row other than the current one,
'the bookmarked row will become the currently selected row.
'Available only at run-time.
'@param        vBookMark Variant. Required.
'@rem Note: if setting a bookmark that does not match any records in the Grid, it will be ignored.
Public Property Let Bookmark(ByVal vBookmark As Variant)
   Dim lRow As Long

   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".Bookmark", "Cannot set bookmark in grid with no Recordset or no Rows"
   If IsEmpty(vBookmark) Then GoTo BookMark_Err
   
On Error GoTo BookMark_Err
         
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
   End If
      
   m_RsFiltered.Bookmark = vBookmark

   lRow = m_RsFiltered.AbsolutePosition - 1
         
   m_CurrRow = lRow
   m_CurrRowSel = m_CurrRow
   
   If m_blnShown Then
      Grid1.Redraw = False
   
      If Not IsVisible(lRow) Then
         SetScrollValue lRow
      Else
         Grid1.Row = (lRow - GetScrollValue) + Grid1.FixedRows
         Grid1.RowSel = Grid1.Row
         CalcPaintedArea
      End If
      Grid1.Redraw = m_blnRedraw
   End If
   
   m_LastRow = m_CurrRow
      
   On Error Resume Next
   If m_blnHasFocus And m_blnShown Then Grid1.SetFocus

Exit Property

BookMark_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".Bookmark" & Err.Source, "Invalid Bookmark: " & Err.Description
End Property

Public Property Get BorderStyle() As nfgBorderStyleSettings
Attribute BorderStyle.VB_HelpID = 1420
   BorderStyle = UserControl.BorderStyle
End Property

'**
'Returns or sets the current BorderStyle of the Grid.
'@param        vbBorderStyle Integer. Required. One of the members of the nfgBorderStyleSettings.
Public Property Let BorderStyle(ByVal vbBorderStyle As nfgBorderStyleSettings)
   UserControl.BorderStyle = vbBorderStyle
   UserControl_Resize
   PropertyChanged "BorderStyle"
End Property

Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_HelpID = 1420
   CausesValidation = Grid1.CausesValidation
End Property

'**
'Determines whether the Grid will cause the validation event for another control about to lose focus.
'@param        bValue Boolean. Required.
Public Property Let CausesValidation(ByVal bValue As Boolean)
   Grid1.CausesValidation() = bValue
   PropertyChanged "CausesValidation"
End Property

'**
'Returns the height of the currently selected cell in the grid.
'Read-Only.
Public Property Get CellHeight() As Single
Attribute CellHeight.VB_HelpID = 1440
Attribute CellHeight.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   CellHeight = Grid1.CellHeight
End Property

'**
'Returns the distance, in pixels, from the left edge of the grid to the left edge of the currently selected cell in the grid.
'Read-Only.
Public Property Get CellLeft() As Single
Attribute CellLeft.VB_HelpID = 1450
Attribute CellLeft.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   CellLeft = Grid1.CellLeft
End Property

Public Property Set ColHeaderPicture(ByVal vColIndex As Variant, _
                              Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter, _
                              ByVal sColPicture As StdPicture)
   
   If Not Ambient.UserMode Then Err.Raise 387
   If Not m_blnHeaderRow Then Err.Raise ERR_NOHDR, Ambient.DisplayName & ".ColHeaderPicture", "Cannot set ColHeaderPicture in grid without Header Row"
   If m_ntColumns.Exists(vColIndex) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColHeaderPicture", MSG_INVCOL
      
On Error Resume Next
   
   If Not m_colColpics.Exists(m_ntColumns(vColIndex).Name, "HDR") Then
      m_colColpics.Add m_ntColumns(vColIndex).Name, m_ntColumns(vColIndex).ColID, "HDR"
   End If
   m_colColpics(m_ntColumns(vColIndex).Name, "HDR").PictureAlignment = ePicAlignment
   Set m_colColpics(m_ntColumns(vColIndex).Name, "HDR").CellPicture = sColPicture
      
   If m_blnShown And m_blnRedraw Then
      If bEditing Then Grid1_GotFocus
      Grid1.Redraw = False
      m_blnLoading = True
      m_blnIgnoreRCChange = True
      SaveRestorePrevGrid Grid1
      Grid1.Col = m_ntColumns(vColIndex).index + Grid1.FixedCols
      Grid1.Row = 0
      Set Grid1.CellPicture = sColPicture
      SaveRestorePrevGrid Grid1, False
      m_blnLoading = False
      m_blnIgnoreRCChange = False
      Grid1.Redraw = m_blnRedraw
   End If
   
End Property

Public Property Get ColHeaderPicture(ByVal vColIndex As Variant, _
                              Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter) As StdPicture
Attribute ColHeaderPicture.VB_HelpID = 1460
Attribute ColHeaderPicture.VB_MemberFlags = "400"
   
   If Not Ambient.UserMode Then Err.Raise 387
   If Not m_blnHeaderRow Then Exit Property
   If m_ntColumns.Exists(vColIndex) = False Then Exit Property
   
   If m_colColpics.Exists(m_ntColumns(vColIndex).Name, "HDR") Then
      Set ColHeaderPicture = m_colColpics(m_ntColumns(vColIndex).Name, "HDR").CellPicture
   End If

End Property

Public Property Set RowPicture(ByVal lngRow As Long, _
                              Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter, _
                              ByVal sRowPicture As StdPicture)
   Dim vBkmrk As Variant
      
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowPicture", "Cannot Set RowPicture in grid with no rows"
   If Not m_blnRecordSelectors Then Err.Raise ERR_NOSEL, Ambient.DisplayName & ".RowPicture", "Cannot set RowPicture in grid without RecordSelectors"
   
On Error Resume Next

   If bEditing Then Grid1_GotFocus
   
   vBkmrk = Me.BookmarkFromRow(lngRow)
   If Not m_colColpics.Exists("RS", vBkmrk) Then m_colColpics.Add "RS", -2, vBkmrk
   m_colColpics("RS", vBkmrk).PictureAlignment = ePicAlignment
   Set m_colColpics("RS", vBkmrk).CellPicture = sRowPicture
         
   If IsVisible(lngRow) And m_blnShown And m_blnRedraw Then
      m_blnLoading = True
      m_blnIgnoreRCChange = True
      SaveRestorePrevGrid Grid1
      Grid1.Col = 0
      Grid1.Row = (lngRow - GetScrollValue) + Grid1.FixedRows
      Set Grid1.CellPicture = sRowPicture
      SaveRestorePrevGrid Grid1, False
      m_blnLoading = False
      m_blnIgnoreRCChange = False
   End If
     
End Property

Public Property Get RowPicture(ByVal lngRow As Long, _
                              Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter) As StdPicture
Attribute RowPicture.VB_HelpID = 1470
Attribute RowPicture.VB_MemberFlags = "400"
   
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Exit Property
   If Not m_blnRecordSelectors Then Exit Property
   
   Dim vBkmrk As Variant
   vBkmrk = Me.BookmarkFromRow(lngRow)
   If m_colColpics.Exists("RS", vBkmrk) Then Set RowPicture = m_colColpics("RS", vBkmrk).CellPicture
        
End Property

Public Property Set CellPicture(ByVal vColIndex As Variant, ByVal lngRow As Long, _
                                Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter, _
                                ByVal sCellPicture As StdPicture)
   Dim cp As cPic
   Dim sName As String
   Dim vBkmrk As Variant
      
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".CellPicture", "Cannot Set CellPicture in grid with no rows"
   If m_ntColumns.Exists(vColIndex) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".CellPicture", MSG_INVCOL
   
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
   End If
   
   vBkmrk = Me.BookmarkFromRow(lngRow)
   sName = m_ntColumns(vColIndex).Name

   If Not m_colColpics.Exists(sName, vBkmrk) Then
      Set cp = m_colColpics.Add(sName, m_ntColumns(sName).ColID, vBkmrk)
   Else
      Set cp = m_colColpics(sName, vBkmrk)
   End If
      
   Set cp.CellPicture = sCellPicture
   cp.PictureAlignment = ePicAlignment
     
   If IsVisible(lngRow) And m_blnShown And m_blnRedraw Then
      m_blnLoading = True
      m_blnIgnoreRCChange = True
      SaveRestorePrevGrid Grid1
      Grid1.Col = m_ntColumns(vColIndex).index + Grid1.FixedCols
      Grid1.Row = (lngRow - GetScrollValue) + Grid1.FixedRows
      Set Grid1.CellPicture = sCellPicture
      SaveRestorePrevGrid Grid1, False
      m_blnLoading = False
      m_blnIgnoreRCChange = False
   End If
        
   Set cp = Nothing

End Property

Public Property Get CellPicture(ByVal vColIndex As Variant, ByVal lngRow As Long, _
                                Optional ByVal ePicAlignment As nfgGridAlignment = nfgAlignCenter) As StdPicture
Attribute CellPicture.VB_HelpID = 1480
Attribute CellPicture.VB_MemberFlags = "400"

   Dim vBkmrk As Variant
   
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Exit Property
   If m_ntColumns.Exists(vColIndex) = False Then Exit Property
      
   vBkmrk = Me.BookmarkFromRow(lngRow)

   If m_colColpics.Exists(m_ntColumns(vColIndex).Name, vBkmrk) Then
      Set CellPicture = m_colColpics(m_ntColumns(vColIndex).Name, vBkmrk).CellPicture
   End If

End Property

'**
'Returns the distance, in pixels, from the top edge of the grid to the top edge of the currently selected cell in the grid.
'Read-Only.
Public Property Get CellTop() As Single
Attribute CellTop.VB_HelpID = 1490
Attribute CellTop.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   CellTop = Grid1.CellTop
End Property

'**
'Returns the property of particular cell in the grid.
Public Property Get CellValue(Optional ByVal lRow As Long = -1, Optional ByVal vKey As Variant = -1) As Variant
Attribute CellValue.VB_HelpID = 1500
Attribute CellValue.VB_MemberFlags = "400"

   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Exit Property
   If lRow < -1 Or lRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".CellValue", MSG_INVROW
   
   If vKey = -1 Then vKey = m_CurrCol
   
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".CellValue", MSG_INVCOL
   
   If lRow = -1 Then lRow = m_CurrRow

   m_RsFiltered.AbsolutePosition = lRow + 1

   CellValue = m_RsFiltered.Fields(m_ntColumns(vKey).Name).Value

End Property

'**
'Returns the width, in pixels, of the currently selected cell in the grid.
'Read-Only.
Public Property Let CellValue(Optional ByVal lRow As Long = -1, Optional ByVal vKey As Variant = -1, ByVal vValue As Variant)
   Dim prevValue As Variant
   Dim vBkmrk As Variant
   Dim vFilter As Variant
   Dim bChecked As Boolean
   
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".CellValue", "Cannot Set CellValue in grid with no rows"
   If lRow < -1 Or lRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".CellValue", MSG_INVROW
   
   If vKey = -1 Then vKey = m_CurrCol

   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".CellValue", MSG_INVCOL
   
 On Error GoTo CellVal_Err
   
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
   End If
   
   If lRow = -1 Then lRow = m_CurrRow

   m_blnLoading = True

   'Move to correct record
   m_RsFiltered.AbsolutePosition = lRow + 1
   m_RSMaster.Bookmark = m_RsFiltered.Bookmark

   'Store Current Value to compare new and change total if necessary
   prevValue = m_RsFiltered.Fields(m_ntColumns(vKey).Name).Value

   Dim arrRows(0) As Long
   arrRows(0) = lRow
   Dim pCancel As Boolean
   Dim bValid As Boolean
   bValid = True

   RaiseEvent BeforeEdit(ntEditField, CStr(prevValue & ""), CStr(vValue & ""), m_ntColumns(vKey).Name, arrRows, bValid, pCancel)

   If pCancel Then Exit Property
   
   m_RSMaster.Fields(m_ntColumns(vKey).Name).Value = vValue
   vValue = m_RSMaster.Fields(m_ntColumns(vKey).Name).Value
   m_RsFiltered.Fields(m_ntColumns(vKey).Name).Value = vValue
  
   If m_blnShown Then
      
      If m_RsFiltered.AbsolutePosition <> lRow + 1 Then
         
         Dim bEvent As Boolean
         
         m_blnIgnoreSel = True
         m_blnLoading = True
        
         Grid1.Redraw = False
         FillTextmatrix GetScrollValue()
         bEvent = (m_CurrRowSel <> m_CurrRow)
         m_CurrRowSel = m_CurrRow
         Grid1.RowSel = Grid1.Row
                 
         Grid1.Redraw = m_blnRedraw
                
         If bEvent Then RaiseEvent SelChange
         
         m_blnIgnoreSel = False
         m_blnLoading = False
      
      Else
      
         If IsVisible(lRow) Then
      
            If m_ntColumns(vKey).Visible Then
      
               With Grid1
      
                  .Redraw = False
      
                  m_blnIgnoreSel = True
                  m_blnLoading = True
      
                  SaveRestorePrevGrid Grid1
      
                  .Row = lRow - GetScrollValue() + .FixedRows
                  .Col = m_ntColumns(vKey).index + .FixedCols
      
                  If m_ntColumns(vKey).ColFormat = nfgBooleanCheckBox Then
                     If IsNull(vValue) Or vValue = "" Then
                        bChecked = False
                     Else
                        bChecked = CBool(vValue)
                     End If
                     Call SetPicture(bChecked, m_ntColumns(vKey).Enabled)
                  Else
                     .TextMatrix(.Row, .Col) = m_ntColumns(vKey).FormatValue(vValue)
                  End If
      
                  If m_ntColumns(vKey).UseCriteria Then
                     If m_blnColorByRow = True Then
                        .Col = .FixedCols
                        .colSel = .Cols - 1
                        .FillStyle = flexFillRepeat
                        .CellBackColor = CheckRowColCriteria(lRow)
                        .FillStyle = flexFillSingle
                        .Col = m_ntColumns(vKey).index + .FixedCols
                     Else
                        .CellBackColor = CheckRowColCriteria(lRow, m_ntColumns(vKey).index)
                     End If
                  End If
      
                  SaveRestorePrevGrid Grid1, False
      
                  m_blnIgnoreSel = False
                  m_blnLoading = False
      
                  .Redraw = m_blnRedraw
      
               End With
      
            End If
      
         End If
      End If
   End If
    
   If m_ntColumns(vKey).ShowTotal Then
      lbltotal(m_ntColumns(vKey).index).Text = _
         m_ntColumns(vKey).FormatValue(CCur(lbltotal(m_ntColumns(vKey).index).Text) + (vValue - prevValue))
   End If

   m_blnLoading = False

   RaiseEvent AfterEdit(ntEditField, CStr(vValue & ""), m_ntColumns(vKey).Name, arrRows)

Exit Property

CellVal_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".CellValue" & Err.Source, Err.Description
End Property

'**
'Returns the width, in pixels, of the currently selected cell in the grid.
'Read-Only.
Public Property Get CellValueFormatted(Optional ByVal lRow As Long = -1, Optional ByVal vKey As Variant = -1) As String
Attribute CellValueFormatted.VB_HelpID = 1510
Attribute CellValueFormatted.VB_MemberFlags = "400"

   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or Not HasRecords Then Exit Property
      
   If lRow < -1 Or lRow > m_RsFiltered.RecordCount - 1 Then _
      Err.Raise ERR_INVROW, Ambient.DisplayName & ".CellValueFormatted", MSG_INVROW
   
   If vKey = -1 Then vKey = m_CurrCol

   If m_ntColumns.Exists(vKey) = False Then _
       Err.Raise ERR_INVCOL, Ambient.DisplayName & ".CellValueFormatted", MSG_INVCOL
       
On Error GoTo CellVal_Err
   
   If lRow = -1 Then lRow = m_CurrRow

   m_RsFiltered.AbsolutePosition = lRow + 1

   CellValueFormatted = m_ntColumns(vKey).FormatValue(m_RsFiltered.Fields(m_ntColumns(vKey).Name).Value)

Exit Property

CellVal_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".CellValueFormatted: " & Err.Source, Err.Description
End Property

'**
'Returns the width, in pixels, of the currently selected cell in the grid.
'Read-Only.
Public Property Get CellWidth() As Single
Attribute CellWidth.VB_HelpID = 1520
Attribute CellWidth.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   CellWidth = Grid1.CellWidth
End Property

Public Property Get Col() As Variant
Attribute Col.VB_HelpID = 1530
Attribute Col.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Then
      Col = 0
   Else
      Col = m_CurrCol
   End If
End Property

'**
'Returns a long specifying the currently selected column in the grid.
'Sets the currently selected column in the grid, either by index, or field name.
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
Public Property Let Col(ByVal vKey As Variant)

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".Col", MSG_INVCOL
   If m_ntColumns(vKey).Visible = False Then Err.Raise ERR_INVCOL + 1, Ambient.DisplayName & ".Col", "Cannot set Col property to hidden column"
   
 On Error GoTo Col_Err
      
   m_CurrCol = m_ntColumns(vKey).index
   m_CurrColSel = m_CurrCol
      
   If m_blnShown Then
      
      If bEditing Then Grid1_GotFocus
      
      m_blnManualScroll = True
      
      If Not m_bln_NeedHorzScroll Then
         Grid1.Col = ((m_ntColumns(vKey).index) + Grid1.FixedCols)
         m_CurrCol = Grid1.Col - Grid1.FixedCols
      Else
         If Grid1.ColIsVisible(m_CurrCol + Grid1.FixedCols) = False Or _
            Grid1.ColPos(m_CurrCol + Grid1.FixedCols) + Grid1.ColWidth(m_CurrCol) > Grid1.Width Then
            If (m_ntColumns(vKey).index) > hScroll.Max Then
               hScroll.Value = hScroll.Max
            Else
               hScroll.Value = ((m_ntColumns(vKey).index))
            End If
         End If
      End If
   
      m_blnManualScroll = False
   
      Call CalcPaintedArea
      
      m_LastCol = m_CurrCol
      
      On Error Resume Next
   
      If m_blnHasFocus Then Grid1.SetFocus
   
   End If
   
Exit Property

Col_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".Col: " & Err.Source, Err.Description
End Property

Public Property Get colSel() As Variant
Attribute colSel.VB_HelpID = 1540
Attribute colSel.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Then
      colSel = 0
   Else
      If m_CurrColSel < 0 Then m_CurrColSel = 0
      If m_CurrColSel > (Grid1.Cols - 1) - Grid1.FixedCols Then m_CurrColSel = (Grid1.Cols - 1) - Grid1.FixedCols
      colSel = m_CurrColSel
   End If
End Property

'**
'Returns a long specifying the currently selected column in the grid.
'Sets the currently selected column in the grid, either by index, or field name.
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
Public Property Let colSel(ByVal vKey As Variant)

   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Then Err.Raise 380

   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColSel", MSG_INVCOL
   If m_ntColumns(vKey).Visible = False Then Err.Raise ERR_INVCOL + 1, Ambient.DisplayName & ".ColSel", "Cannot set ColSel property to hidden column"

On Error Resume Next
   
   m_CurrColSel = m_ntColumns(vKey).index
   RaiseEvent SelChange
   
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
      Call CalcPaintedArea
      On Error Resume Next
      If m_blnHasFocus Then Grid1.SetFocus
   End If
   
End Property

Public Property Get ColorByRow() As Boolean
Attribute ColorByRow.VB_HelpID = 1550
   ColorByRow = m_blnColorByRow
End Property

'**
'Returns or sets the value determining whether the grid will color an entire row,
'or just the matching cell, if criteria matching for that column is enabled.
'@param        bValue Boolean. Required.
Public Property Let ColorByRow(ByVal bValue As Boolean)
   m_blnColorByRow = bValue
   PropertyChanged ("ColorByRow")
   If Ambient.UserMode Then
      If m_blnShown Then
         If Not IsUnbound Then
            If HasRecords Then
               If bEditing Then Grid1_GotFocus
               Call SetScrollValue(GetScrollValue())
            End If
         End If
      End If
   End If
End Property

Public Property Get ColHeaderText(ByVal vKey As Variant) As String
Attribute ColHeaderText.VB_HelpID = 1560
Attribute ColHeaderText.VB_MemberFlags = "400"
   
   ColHeaderText = ""
   
   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColHeaderText", MSG_INVCOL
 
   ColHeaderText = m_ntColumns(vKey).HeaderText

End Property

'**
'Returns or sets the text displayed in the Grid header for a particular row
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
'@param        sText String. Required. A valid alpha-numeric string, containing the letters A-Z, the numbers 0-9, or the underscore character (_).
'@rem Note: Only applies if ShowHeaderRow is True
Public Property Let ColHeaderText(ByVal vKey As Variant, ByVal sText As String)

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColHeaderText", MSG_INVCOL
 
On Error GoTo ColHdr_Err
   
   If bEditing Then Grid1_GotFocus
   
   m_ntColumns(vKey).HeaderText = sText

   If m_blnShown Then
      If Grid1.FixedRows > 0 Then
         If m_ntColumns(vKey).Visible Then
            Grid1.TextMatrix(0, m_ntColumns(vKey).index + Grid1.FixedCols) = sText
            If m_blnAutoSizeColumns Then
               If TextWidth(sText) + 300 > Grid1.ColWidth(m_ntColumns(vKey).index + Grid1.FixedCols) Then
                  Grid1.ColWidth(m_ntColumns(vKey).index + Grid1.FixedCols) = TextWidth(sText) + 300
                  Call RecalcGrid
               End If
            End If
         End If
      End If
   End If

Exit Property

ColHdr_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ColHeaderText: " & Err.Source, Err.Description

End Property

Public Property Get ColNegativeColor() As OLE_COLOR
Attribute ColNegativeColor.VB_HelpID = 1570
  ColNegativeColor = m_clrNegative
End Property

Public Property Let ColNegativeColor(ByVal cColor As OLE_COLOR)
   m_clrNegative = cColor
   PropertyChanged "NegativeColor"
   If Ambient.UserMode Then
      If Not IsUnbound Then
         If HasRecords Then
            If m_blnShown Then
               If bEditing Then Grid1_GotFocus
               Call SetScrollValue(GetScrollValue())
            End If
         End If
      End If
   End If
End Property

Public Property Get ColPosition(ByVal vIndexKey As Variant) As Long
Attribute ColPosition.VB_HelpID = 1580
   On Error Resume Next
   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vIndexKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColPosition", MSG_INVCOL
   ColPosition = Grid1.ColPos(m_ntColumns(vIndexKey).index + Grid1.FixedRows)
End Property

Public Property Get ColPositiveColor() As OLE_COLOR
Attribute ColPositiveColor.VB_HelpID = 1590
  ColPositiveColor = m_clrPositive
End Property

Public Property Let ColPositiveColor(ByVal cColor As OLE_COLOR)
   m_clrPositive = cColor
   PropertyChanged "PositiveColor"
   If Ambient.UserMode Then
      If Not IsUnbound Then
         If m_blnShown Then
            If HasRecords Then
               If bEditing Then Grid1_GotFocus
               Call SetScrollValue(GetScrollValue())
            End If
         End If
      End If
   End If
End Property

'**
'Returns the total number of Columns(without recordselector) that exist in the Grid.
Public Property Get Cols() As Long
Attribute Cols.VB_HelpID = 1600
Attribute Cols.VB_MemberFlags = "400"
   If Ambient.UserMode = False Then Err.Raise 387
   Cols = 0
   If Not m_ntColumns Is Nothing Then Cols = m_ntColumns.count
End Property

'**
'Returns the running total for the column specified by index or field name.
'@param        vKey Variant. Required. The index or key of the column to retrieve a total for.
Public Property Get ColumnTotal(ByVal vKey As Variant) As String
Attribute ColumnTotal.VB_HelpID = 1610
Attribute ColumnTotal.VB_MemberFlags = "400"
   If Ambient.UserMode = False Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColumnTotal", MSG_INVCOL
   ColumnTotal = vbNullString
   If m_ntColumns(vKey).ShowTotal Then ColumnTotal = lbltotal(m_ntColumns(vKey).index).Text
End Property

'**
'Returns the layout currently being used by the grid as an ntColLayout object.
Public Property Get Columns() As ntColumns
Attribute Columns.VB_HelpID = 1620
Attribute Columns.VB_MemberFlags = "400"
   If m_blnGridMode Then
      If IsUnbound Then Exit Property
      Set Columns = m_ntColumns
   Else
      If m_ntColumns Is Nothing Then Set m_ntColumns = New ntColumns
      Set Columns = m_ntColumns
   End If
End Property

Public Property Get ColVisible(ByVal vKey As Variant) As Boolean
Attribute ColVisible.VB_HelpID = 1630
Attribute ColVisible.VB_MemberFlags = "400"

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColVisible", MSG_INVCOL
   ColVisible = m_ntColumns(vKey).Visible
 
End Property
'**
'Returns or sets whether a column is visible in the grid or not
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
'@param        bValue Boolean. Required.
Public Property Let ColVisible(ByVal vKey As Variant, ByVal bValue As Boolean)

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColVisible", MSG_INVCOL
     
On Error GoTo ColVisible_Err
      
   If Not m_blnShown Then
      m_ntColumns(vKey).Visible = bValue
   Else
      If bEditing Then Grid1_GotFocus
      ' If it is being changed then reset grid
      If bValue <> m_ntColumns(vKey).Visible Then
         GridMousePointer = vbHourglass
         Screen.MousePointer = vbHourglass
         Grid1.Redraw = False
         Call ResetColumn(vKey, bValue)
         GridMousePointer = vbDefault
         Screen.MousePointer = vbDefault
         Grid1.Redraw = m_blnRedraw
      End If
   End If
   
Exit Property

ColVisible_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ColVisible: " & Err.Source, Err.Description

End Property

Public Property Get ColWidth(ByVal vKey As Variant) As Single
Attribute ColWidth.VB_HelpID = 1640
Attribute ColWidth.VB_MemberFlags = "400"

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColWidth", MSG_INVCOL
 
   ColWidth = m_ntColumns(vKey).Width
   
End Property
'**
'Returns or sets the width of a column in the grid.
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
'@param        sWidth Single. Required.
'@rem Note: If width specified is outside of the range specified by
'the ColWidthMin or ColWidthMax properties, it will be matched to
'the appropriate one - (If ColWidthMin = 0 or ColWidthMax = 0, they have no effect).
Public Property Let ColWidth(ByVal vKey As Variant, ByVal sWidth As Single)

   If Not Ambient.UserMode Then Err.Raise 387
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColWidth", MSG_INVCOL
 
On Error GoTo ColWidth_Err
   
   m_ntColumns(vKey).Width = sWidth
   
   If m_blnShown Then
      
      If bEditing Then Grid1_GotFocus
   
      If (sWidth > m_sngMaxColWidth And m_sngMaxColWidth <> 0) Then sWidth = m_sngMaxColWidth
      If sWidth < m_sngMinColWidth Then sWidth = m_sngMinColWidth
      If sWidth <= 0 Then sWidth = 0
      If sWidth <> 0 Then
         ' If it is Visible then allow change and reset grid
         If m_ntColumns(vKey).Visible Then
            Grid1.ColWidth(m_ntColumns(vKey).index + Grid1.FixedCols) = sWidth
         End If
         Call SetGridRows
         SetTotalPositions
      End If
   
   End If
   
Exit Property

ColWidth_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ColWidth: " & Err.Source, Err.Description
End Property

Public Property Get ColWidthMax() As Single
Attribute ColWidthMax.VB_HelpID = 1650
Attribute ColWidthMax.VB_MemberFlags = "400"
   ColWidthMax = m_sngMaxColWidth
End Property

'**
'Returns or sets the maximum width of a column in the grid.
'@param        lMaxWidth Single. Required.
'@rem Note: a setting of 0 has no effect on the grid.
Public Property Let ColWidthMax(ByVal sMaxWidth As Single)
   Dim loncols As Long

On Error GoTo ColWidthMax_Err

   If sMaxWidth < m_sngMinColWidth Then Err.Raise vbObjectError + 5020, Ambient.DisplayName & ".ColWidthMax", "Cannot set ColWidthMax lower then ColWidthMin"
   If Not (sMaxWidth >= 0 And sMaxWidth <= 32000) Then Err.Raise vbObjectError + 5021, Ambient.DisplayName & ".ColWidthMax", "ColWidthMax must be between 0 and 32,000"
   
   If m_blnShown Then If bEditing Then Grid1_GotFocus
   
   m_sngMaxColWidth = sMaxWidth

   If Not m_sngMaxColWidth = 0 Then

      If Ambient.UserMode = True Then

         If m_ntColumns.count > 0 Then
            For loncols = 0 To m_ntColumns.count - 1
               If m_ntColumns(loncols).Width > m_sngMaxColWidth Then
                  m_ntColumns(loncols).Width = m_sngMaxColWidth
               End If
               If m_blnShown Then
                  If m_ntColumns(loncols).Visible = True Then
                     If Grid1.ColWidth(m_ntColumns(loncols).index + Grid1.FixedCols) > m_sngMaxColWidth Then
                        Grid1.ColWidth(m_ntColumns(loncols).index + Grid1.FixedCols) = m_sngMaxColWidth
                     End If
                  End If
               End If
            Next loncols
         End If

      Else

         For loncols = Grid1.FixedCols To Grid1.Cols - 1
            If Grid1.ColWidth(loncols) > m_sngMaxColWidth Then
               Grid1.ColWidth(loncols) = m_sngMaxColWidth
            End If
         Next loncols

      End If

   End If

   PropertyChanged "MaxColWidth"
   
   If m_blnShown Then
      Call SetGridRows
      SetTotalPositions
   End If
   
Exit Property

ColWidthMax_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ColWidthMax: " & Err.Source, Err.Description
End Property

Public Property Get ColWidthMin() As Single
Attribute ColWidthMin.VB_HelpID = 1660
Attribute ColWidthMin.VB_MemberFlags = "400"
  ColWidthMin = m_sngMinColWidth
End Property

'**
'Returns or sets the minimum width of a column in the grid.
'@param        sMinWidth Single. Required.
'@rem Note: a setting of 0 has no effect on the grid.
Public Property Let ColWidthMin(ByVal sMinWidth As Single)
  Dim loncols As Long

On Error GoTo ColWidthMin_Err

   If (sMinWidth > m_sngMaxColWidth) And (m_sngMaxColWidth > 0) Then Err.Raise vbObjectError + 5022, Ambient.DisplayName & ".ColWidthMin", "Cannot set ColWidthMin lower than ColWidthMax"
   If Not (sMinWidth >= 0 And sMinWidth <= 32000) Then Err.Raise vbObjectError + 5023, Ambient.DisplayName & ".ColWidthMin", "ColWidthMin must be between 0 and 32,000"
   
   If m_blnShown And Ambient.UserMode Then If bEditing Then Grid1_GotFocus
   
   m_sngMinColWidth = sMinWidth

   If Not m_sngMinColWidth = 0 Then

      If Ambient.UserMode = True Then

         If m_ntColumns.count > 0 Then
            For loncols = 0 To m_ntColumns.count - 1
               If m_ntColumns(loncols).Width < m_sngMinColWidth Then
                  m_ntColumns(loncols).Width = m_sngMinColWidth
               End If
               If m_blnShown Then
                  If m_ntColumns(loncols).Visible = True Then
                     If Grid1.ColWidth(m_ntColumns(loncols).index + Grid1.FixedCols) < m_sngMinColWidth Then
                        Grid1.ColWidth(m_ntColumns(loncols).index + Grid1.FixedCols) = m_sngMinColWidth
                     End If
                  End If
               End If
            Next loncols
         End If

      Else

         For loncols = Grid1.FixedCols To Grid1.Cols - 1
            If Grid1.ColWidth(loncols) > m_sngMaxColWidth Then
               Grid1.ColWidth(loncols) = m_sngMaxColWidth
            End If
         Next loncols

      End If

   End If

   PropertyChanged "MinColWidth"

   If m_blnShown Then
      SetGridRows
      SetTotalPositions
   End If
   
Exit Property

ColWidthMin_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ColWidthMin: " & Err.Source, Err.Description
End Property

Public Property Get DataRows() As ntRowInfo
Attribute DataRows.VB_HelpID = 1670
Attribute DataRows.VB_MemberFlags = "400"
   Set DataRows = m_colRowColors
End Property

Public Property Get EditRangeKey() As EditKey
Attribute EditRangeKey.VB_HelpID = 1680
  EditRangeKey = m_intEditKey
End Property

'**
'Returns or sets whether the grid will respond to user-generated events.
'@param        bValue Boolean. Required.
Public Property Let EditRangeKey(ByVal eValue As EditKey)
   m_intEditKey = eValue
   PropertyChanged "EditRangeKey"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_HelpID = 1690
  Enabled = m_blnEnabled
End Property

'**
'Returns or sets whether the grid will respond to user-generated events.
'@param        bValue Boolean. Required.
Public Property Let Enabled(ByVal bValue As Boolean)
   Dim l As Long

On Error Resume Next
   
   m_blnEnabled = bValue
   
   If bValue = False Then
      If m_blnShown Then If bEditing Then Grid1_GotFocus
      Grid1.BackColor = m_clrDisBackcolor
      If lbltotal.UBound > 0 Then
         For l = 0 To lbltotal.UBound
            lbltotal(l).BackColor = m_clrDisBackcolor
         Next l
      End If
   Else
      Grid1.BackColor = m_clrEnabledBackcolor
      If lbltotal.UBound > 0 Then
         For l = 0 To lbltotal.UBound
            lbltotal(l).BackColor = m_clrEnabledBackcolor
         Next l
      End If
   End If
   
   UserControl.Enabled = bValue
   Grid1.Enabled = bValue
   PropertyChanged "Enabled"

End Property

Public Property Get FocusRect() As nfgFocusRectSettings
Attribute FocusRect.VB_HelpID = 1700
  FocusRect = m_intFocusRect
End Property

'**
'Returns or sets the style of focus rectangle displayed by the grid when a cell is in focus.
'@param        fgSetting  Integer. Required. One of the members of the FocusRectSettings enumeration.
Public Property Let FocusRect(ByVal fgSetting As nfgFocusRectSettings)
  m_intFocusRect = fgSetting
  Grid1.FocusRect() = m_intFocusRect
  PropertyChanged "FocusRect"
End Property

Public Property Get Font() As Font
Attribute Font.VB_HelpID = 1710
  Set Font = Grid1.Font
End Property

'**
'Returns or sets the current font used by the grid.
'@param        new_Font SystemFont. Required.
Public Property Set Font(ByVal New_Font As Font)
  Set Grid1.Font = New_Font
  Rem Set txtEdit.Font = New_Font
  Rem Set cmbEdit.Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_HelpID = 1720
  ForeColor = Grid1.ForeColor
End Property

'**
'Returns or sets the current forecolor of the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let ForeColor(ByVal cColor As OLE_COLOR)
  Grid1.ForeColor() = cColor
  PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorHeader() As OLE_COLOR
Attribute ForeColorHeader.VB_HelpID = 1730
  ForeColorHeader = Grid1.ForeColorFixed
End Property

'**
'Returns or sets the current forecolor in the header row of the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
'@rem Note: this only applies if ShowHeaderRow property is true.
Public Property Let ForeColorHeader(ByVal cColor As OLE_COLOR)
  Grid1.ForeColorFixed() = cColor
  PropertyChanged "ForeColorFixed"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_HelpID = 1740
  ForeColorSel = Grid1.ForeColorSel
End Property

'**
'Returns or sets the current forecolor of the selected column or columns in the header row of the Grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
'@rem Note: this only applies if ShowHeaderRow property is true.
Public Property Let ForeColorSel(ByVal cColor As OLE_COLOR)
  Grid1.ForeColorSel() = cColor
  PropertyChanged "ForeColorSel"
End Property

Public Property Let ForeGroundDrawMode(ByVal cMode As ntFgForeGroundDrawMode)
   m_FgrdDraw = cMode
   PropertyChanged "ForeGroundDrawMode"
End Property

Public Property Get ForeGroundDrawMode() As ntFgForeGroundDrawMode
   ForeGroundDrawMode = m_FgrdDraw
End Property

'**
'Returns a string array containing all the formatted filters applied, either programatically,
'or by the user, in the order they were applied.
'@rem Read - only.
Public Property Get FormattedFilters() As ntColFilters
Attribute FormattedFilters.VB_HelpID = 1750
Attribute FormattedFilters.VB_MemberFlags = "400"
   Set FormattedFilters = m_colFilters
End Property

Public Property Get GradientStartColor() As OLE_COLOR
   GradientStartColor = m_GradClrStart
End Property

Public Property Let GradientStartColor(ByVal Value As OLE_COLOR)
   m_GradClrStart = Value
   PropertyChanged "GradientStartColor"
End Property

Public Property Get GradientEndColor() As OLE_COLOR
   GradientEndColor = m_GradClrEnd
End Property

Public Property Let GradientEndColor(ByVal Value As OLE_COLOR)
   m_GradClrEnd = Value
   PropertyChanged "GradientEndColor"
End Property

Public Property Get GradientType() As ntFgGradientType
   GradientType = m_GradType
End Property

Public Property Let GradientType(ByVal Value As ntFgGradientType)
   m_GradType = Value
   PropertyChanged "GradientType"
End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_HelpID = 1760
   GridColor = Grid1.GridColor
End Property

'**
'Returns or sets the color of the lines displayed in the body of the grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let GridColor(ByVal cColor As OLE_COLOR)
   Grid1.GridColor() = cColor
   PropertyChanged "GridColor"
End Property

Public Property Get GridColorHeader() As OLE_COLOR
Attribute GridColorHeader.VB_HelpID = 1770
   GridColorHeader = Grid1.GridColorFixed
End Property

'**
'Returns or sets the color of the lines displayed in the header of the grid.
'@param        cColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
Public Property Let GridColorHeader(ByVal cColor As OLE_COLOR)
   Grid1.GridColorFixed() = cColor
   PropertyChanged "GridColorFixed"
End Property

Public Property Get GridLines() As nfgGridLineSettings
Attribute GridLines.VB_HelpID = 1780
   GridLines = Grid1.GridLines
End Property

'**
'Returns or sets the style of the lines displayed in the body of the grid.
'@param        fgGridSetting Integer. Required. One of the members of the nfgGridLineSettings enumeration.
Public Property Let GridLines(ByVal fgGridSetting As nfgGridLineSettings)
   Grid1.GridLines() = fgGridSetting
   PropertyChanged "GridLines"
End Property

Public Property Get GridLinesHeader() As nfgGridLineSettings
Attribute GridLinesHeader.VB_HelpID = 1790
   GridLinesHeader = Grid1.GridLinesFixed
End Property

'**
'Returns or sets the style of the lines displayed in the header of the grid.
'@param        fgGridSetting Integer. Required. One of the members of the nfgGridLineSettings enumeration.
Public Property Let GridLinesHeader(ByVal fgGridSetting As nfgGridLineSettings)
   Grid1.GridLinesFixed() = fgGridSetting
   PropertyChanged "GridLinesFixed"
End Property

Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_HelpID = 1800
   GridLineWidth = Grid1.GridLineWidth
End Property

'**
'Returns or sets the thickness in pixels of all grid lines.
'@param        iWidth Integer. Required.
Public Property Let GridLineWidth(ByVal iWidth As Integer)
   Grid1.GridLineWidth() = iWidth
   PropertyChanged "GridLineWidth"
End Property

Public Property Get Highlight() As nfgHighlightSettings
Attribute Highlight.VB_HelpID = 1810
   Highlight = m_intHighlight
End Property

'**
'Returns or sets what type of highlighting the Grid will do.
'@param        fgHighLight Integer. Required. One of the members of the nfgHighlightSettings enumeration.
Public Property Let Highlight(ByVal fgHighLight As nfgHighlightSettings)
   m_intHighlight = fgHighLight
   Grid1.Highlight = m_intHighlight
   PropertyChanged "HighLight"
End Property

Public Property Get hScrollBar() As ntFxGdScrollBar
Attribute hScrollBar.VB_HelpID = 1820
Attribute hScrollBar.VB_MemberFlags = "400"
   Set hScrollBar = m_hScroll
End Property
'**
'Returns the handle(Long) to the Grid Control.
Public Property Get hwnd() As Long
Attribute hwnd.VB_HelpID = 1830
Attribute hwnd.VB_MemberFlags = "400"
   hwnd = UserControl.hwnd
End Property

'**
'Returns a boolean value indicating whether the FormattedFilter property has been applied,
'or the user has filtered the grid.
Public Property Get IsFiltered() As Boolean
Attribute IsFiltered.VB_HelpID = 1840
Attribute IsFiltered.VB_MemberFlags = "400"
   If Ambient.UserMode = False Then Exit Property
   If Not IsUnbound Then IsFiltered = m_colFilters.count > 0
End Property

Public Property Get IsReordered() As Boolean
Attribute IsReordered.VB_HelpID = 1850
   Dim pCol As ntColumn
   IsReordered = False
   For Each pCol In m_ntColumns
      If pCol.ColID <> pCol.index Then
         IsReordered = True
         Exit For
      End If
   Next pCol
End Property

Public Property Get IsSorted() As Boolean
Attribute IsSorted.VB_HelpID = 1860
   IsSorted = False
   If m_RSMaster Is Nothing Then Exit Property
   If m_RsFiltered Is Nothing Then Exit Property
   IsSorted = (m_RsFiltered.Sort <> "")
End Property
'**
'Returns a long integer indicating which column in the Grid the mouse is over.
Public Property Get MouseCol() As Long
Attribute MouseCol.VB_HelpID = 1870
Attribute MouseCol.VB_MemberFlags = "400"
   MouseCol = Grid1.MouseCol - Grid1.FixedCols
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_HelpID = 1880
  Set MouseIcon = Grid1.MouseIcon
End Property

'**
'Return or Set the Mouse Icon displayed in the grid.
'@param        mIcon Icon as stdPicture. Required.
'@rem Note: MousePointer property must be set to "99 - custom" to use a custom Mouse Icon.
Public Property Set MouseIcon(ByVal mIcon As Picture)
  Set Grid1.MouseIcon = mIcon
  PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerSettings
Attribute MousePointer.VB_HelpID = 1890
  MousePointer = Grid1.MousePointer
End Property

'**
'Returns or sets the type of mousepointer displayed in the Grid.
'@rem Note: This property must be set to "99 - custom" to use a custom Mouse Icon.
Public Property Let MousePointer(ByVal iPointer As MousePointerSettings)
  Grid1.MousePointer() = iPointer
  PropertyChanged "MousePointer"
End Property

'**
'Returns a long integer indicating which row in the Grid the mouse is over.
Public Property Get MouseRow() As Long
Attribute MouseRow.VB_HelpID = 1900
Attribute MouseRow.VB_MemberFlags = "400"
  MouseRow = Grid1.MouseRow - Grid1.FixedRows
End Property

Public Sub MoveCol(ByVal lOldIndex As Long, ByVal lNewIndex As Long)
Attribute MoveCol.VB_HelpID = 1020
   Dim l As Long
   Dim arrNewindexes() As Long
         
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName, "Cannot move columns in a grid with no recordset or rows"
   If Not m_ntColumns.Exists(lOldIndex) Or Not m_ntColumns.Exists(lNewIndex) Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".MoveCol", MSG_INVCOL
         
On Error Resume Next

   m_blnMoving = True
   
   ReDim arrNewindexes(m_ntColumns.count - 1)
   
   For l = 0 To UBound(arrNewindexes)
      arrNewindexes(l) = l
   Next l
      
   If lOldIndex > lNewIndex Then
      For l = lNewIndex To lOldIndex - 1
         arrNewindexes(l) = arrNewindexes(l) + 1
      Next l
      arrNewindexes(lOldIndex) = lNewIndex
   Else
      For l = lOldIndex + 1 To lNewIndex
         arrNewindexes(l) = arrNewindexes(l) - 1
      Next l
      arrNewindexes(lOldIndex) = lNewIndex
   End If
     
   If m_blnShown Then Grid1.ColPosition(lOldIndex + Grid1.FixedCols) = (lNewIndex + Grid1.FixedCols)
   
   Me.ReorderColumns arrNewindexes, False
   
   m_blnIgnoreSel = True
   m_CurrCol = lNewIndex
   m_CurrColSel = lNewIndex
   Grid1.Col = lNewIndex + Grid1.FixedCols
   Grid1.colSel = lNewIndex + Grid1.FixedCols
   m_blnIgnoreSel = False
   CalcPaintedArea
   RaiseEvent ColEndDrag(m_ntColumns(lNewIndex).Name, lNewIndex)
   
   m_blnMoving = False
   
End Sub

Public Property Get PicChecked() As Picture
Attribute PicChecked.VB_HelpID = 1910
  Set PicChecked = m_picChecked
End Property

'**
'Returns or sets the picture to use for the checkbox displayed
'when a column Format is set to "5 - nfgBooleanCheckBox" and the column is enabled and checked.
'@param        pIcon Icon as stdPicture. Required.
'@rem Setting this property to Nothing will cause the Grid to use its default picture for this column type.
Public Property Set PicChecked(ByVal pIcon As Picture)
On Error GoTo Pic_Err
   If pIcon Is Nothing Then
      Set m_picChecked = LoadResPicture(101, 1)
   Else
      Set m_picChecked = pIcon
   End If
   picCheck.PaintPicture m_picChecked, 0, 0, picCheck.ScaleWidth, picCheck.ScaleHeight
   PropertyChanged "CheckEnabledPic"
Exit Property
Pic_Err:
   Err.Raise Err.Number, "ntFlexGrid.PicChecked: " & Err.Source, Err.Description
End Property

Public Property Get picCheckedDis() As Picture
Attribute picCheckedDis.VB_HelpID = 1920
  Set picCheckedDis = m_picCheckedDis
End Property

'**
'Returns or sets the picture to use for the checkbox displayed
'when a column Format is set to "5 - nfgBooleanCheckBox" and the column is disabled and checked.
'@param        pIcon Icon as stdPicture. Required.
'@rem Setting this property to Nothing will cause the Grid to use its default picture for this column type.
Public Property Set picCheckedDis(ByVal pIcon As Picture)
On Error GoTo Pic_Err
   If pIcon Is Nothing Then
      Set m_picCheckedDis = LoadResPicture(103, 1)
   Else
      Set m_picCheckedDis = pIcon
   End If
   picCheckDis.PaintPicture m_picCheckedDis, 0, 0, picCheckDis.ScaleWidth, picCheckDis.ScaleHeight
   PropertyChanged "CheckDisabledPic"
Exit Property
Pic_Err:
   Err.Raise Err.Number, "ntFlexGrid.PicChecked: " & Err.Source, Err.Description
End Property

Public Property Get picUnChecked() As Picture
Attribute picUnChecked.VB_HelpID = 1930
  Set picUnChecked = m_picUnchecked
End Property

'**
'Returns or sets the picture to use for the checkbox displayed
'when a column Format is set to "nfgBooleanCheckBox" and the column is enabled and unchecked.
'@param        pIcon Icon as stdPicture. Required.
'@rem Setting this property to Nothing will cause the Grid to use its default picture for this column type.
Public Property Set picUnChecked(ByVal pIcon As Picture)
On Error GoTo Pic_Err
   If pIcon Is Nothing Then
      Set m_picUnchecked = LoadResPicture(102, 1)
   Else
      Set m_picUnchecked = pIcon
   End If
   picUnCheck.PaintPicture m_picUnchecked, 0, 0, picUnCheck.ScaleWidth, picUnCheck.ScaleHeight
   PropertyChanged "UnCheckEnabledPic"
Exit Property
Pic_Err:
   Err.Raise Err.Number, "ntFlexGrid.PicChecked: " & Err.Source, Err.Description
End Property

Public Property Get picUnCheckedDis() As Picture
Attribute picUnCheckedDis.VB_HelpID = 1940
  Set picUnCheckedDis = m_picUncheckedDis
End Property

'**
'Returns or sets the picture to use for the checkbox displayed
'when a column Format is set to "5 - nfgBooleanCheckBox" and the column is disabled and unchecked.
'@param        pIcon Icon as stdPicture. Required.
'@rem Setting this property to Nothing will cause the Grid to use its default picture for this column type.
Public Property Set picUnCheckedDis(ByVal pIcon As Picture)
On Error GoTo Pic_Err
   If pIcon Is Nothing Then
      Set m_picUncheckedDis = LoadResPicture(104, 1)
   Else
      Set m_picUncheckedDis = pIcon
   End If
   picUnCheckDis.PaintPicture m_picUncheckedDis, 0, 0, picUnCheckDis.ScaleWidth, picUnCheckDis.ScaleHeight
   PropertyChanged "UnCheckDisabledPic"
Exit Property
Pic_Err:
   Err.Raise Err.Number, "ntFlexGrid.PicChecked: " & Err.Source, Err.Description
End Property

Public Property Get Recordset() As Object
Attribute Recordset.VB_HelpID = 1950
Attribute Recordset.VB_MemberFlags = "400"
   If IsUnbound Then
      Set Recordset = Nothing
   Else
      Set Recordset = m_RSMaster
      Recordset.Filter = m_RSMaster.Filter
   End If
End Property

'**
'Returns or sets the ADO Recordset property of the grid.
'@param        rs ADODB Recordset. Required.
Public Property Set Recordset(ByVal RS As Object)
    
   If Not RS Is Nothing Then
      If RS.State <> 1 Then
         Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".Recordset", "Cannot bind grid to a closed recordset"
      ElseIf Not RS.Activeconnection Is Nothing Then
         Err.Raise ERR_NORS + 2, Ambient.DisplayName & ".Recordset", "Cannot bind 'Recordset' property to a dynamic recordset. Use the 'RecordSource' property instead."
      End If
   End If

On Error GoTo Recordset_Err
   
   If Not IsUnbound Then If bEditing Then Grid1_GotFocus
   
   Call HookRs(RS, False)

   DoEvents

Exit Property

Recordset_Err:

   Err.Raise Err.Number, Ambient.DisplayName & ".Recordset", Err.Description

End Property

'**
'Returns or sets the ADO Recordset property of the grid.
'@param        rs ADODB Recordset. Required.
Public Property Set RecordSource(ByVal RS As Object)
Attribute RecordSource.VB_HelpID = 1960
Attribute RecordSource.VB_MemberFlags = "400"
   
   If Not RS Is Nothing Then
      If RS.State <> 1 Then
         Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".Recordsource", "Cannot bind grid to a closed recordset"
      End If
   End If

On Error GoTo Recordset_Err
   
   If Not IsUnbound Then If bEditing Then Grid1_GotFocus
   
   Call HookRs(RS, True)

   DoEvents

Exit Property

Recordset_Err:

   Err.Raise Err.Number, Ambient.DisplayName & ".Recordsource", Err.Description

End Property

Public Property Let RecordSelectorWidth(ByVal sWidth As Single)
Attribute RecordSelectorWidth.VB_HelpID = 1970
   If sWidth < 0 Then sWidth = 0
   m_sngRecordSelectorWidth = sWidth
   If m_blnRecordSelectors Then
      Grid1.ColWidth(0) = sWidth
   End If
   PropertyChanged "RecordSelWidth"
End Property

Public Property Get RecordSelectorWidth() As Single
   RecordSelectorWidth = m_sngRecordSelectorWidth
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_HelpID = 1980
Attribute Redraw.VB_MemberFlags = "400"
   If Not Ambient.UserMode Then Err.Raise 387
   Redraw = m_blnRedraw
End Property

Public Property Let Redraw(ByVal bRedraw As Boolean)
On Error Resume Next
   If Not Ambient.UserMode Then Err.Raise 387
   m_blnRedraw = bRedraw
   If m_blnShown And m_blnRedraw = True Then
      m_blnLoading = True
      SetScrollValue m_UpdateScrollValue
      m_blnLoading = False
      FillTextmatrix GetScrollValue
      CalcPaintedArea
      Grid1.Redraw = True
      If m_blnHasFocus Then Grid1.SetFocus
   Else
      Grid1.Redraw = m_blnRedraw
   End If
End Property

Public Property Get Row() As Long
Attribute Row.VB_HelpID = 1990
Attribute Row.VB_MemberFlags = "400"
   If IsUnbound Then
      Row = 0
   Else
      If m_CurrRow < 0 Then m_CurrRow = 0
      If m_CurrRow > m_RsFiltered.RecordCount - 1 Then m_CurrRow = m_RsFiltered.RecordCount - 1
      Row = m_CurrRow
   End If
End Property

'**
'Return or set the currently selected row in the Grid.
'@param        lRow Long. Required.
'@rem Note: when setting a row, if the row is not found, setting is ignored.
Public Property Let Row(ByVal lRow As Long)

   If Not Ambient.UserMode Then Err.Raise 387
   If lRow = m_CurrRow Then Exit Property
   If Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".Row", "Cannot set row in grid with no Recordset or No Rows"
   If lRow < 0 Or lRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".Row", MSG_INVROW
   
On Error Resume Next
   
   m_CurrRow = lRow
   m_CurrRowSel = lRow
      
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
      Call CalcPaintedArea
      RaiseEvent RowChange(m_CurrRow)
      On Error Resume Next
      If m_blnHasFocus Then Grid1.SetFocus
   End If
   
   m_LastRow = m_CurrRow
   
End Property

'**
'This sub will update the item data property for a rowInfo object
'@param        RowIndex Long.
'@param        vValue  Variant.
Public Property Let RowItemData(ByVal RowIndex As Long, ByVal vValue As Variant)

   Dim nRowInf As RowInfo
   Dim vBkmrk As Variant
   
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowitemData", "Cannot set RowItemData in Grid with no rows"
   If Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowItemData", "Cannot set RowItemData in grid with no Recordset or No Rows"
   If RowIndex < 0 Or RowIndex > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".RowItemData", MSG_INVROW
   
On Error GoTo RowItemData_Err

   m_RsFiltered.AbsolutePosition = RowIndex + 1
   vBkmrk = m_RsFiltered.Bookmark

   If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
      Set nRowInf = m_colRowColors.RowItem("ID" & CStr(vBkmrk))
   Else
      Set nRowInf = m_colRowColors.NewRowInfo
      nRowInf.BackColor = -1
      nRowInf.ForeColor = -1
      nRowInf.Bookmark = vBkmrk
   End If
   nRowInf.ItemData = vValue
   If Not m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
      m_colRowColors.Insert nRowInf, "ID" & CStr(vBkmrk)
   End If
   Set nRowInf = Nothing

Exit Property

RowItemData_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowItemData: " & Err.Source, Err.Description

End Property

'**
'This sub will return the item data property for a rowInfo object based on the row param
'@param        RowIndex Long.
'@Return        vValue  Variant.
Public Property Get RowItemData(ByVal RowIndex As Long) As Variant
Attribute RowItemData.VB_HelpID = 2000
Attribute RowItemData.VB_MemberFlags = "400"

   Dim nRowInf As RowInfo
   Dim vBkmrk As Variant
         
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowitemData", "Cannot get RowItemData in Grid with no rows"
   If Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowItemData", "Cannot get RowItemData in grid with no Recordset or No Rows"
   If RowIndex < 0 Or RowIndex > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".RowItemData", MSG_INVROW

 On Error Resume Next

   m_RsFiltered.AbsolutePosition = RowIndex + 1
   vBkmrk = m_RsFiltered.Bookmark

   If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
      RowItemData = m_colRowColors.RowItem("ID" & CStr(vBkmrk)).ItemData
   Else
      RowItemData = -1
   End If

End Property

Public Property Get RowHeight() As Single
Attribute RowHeight.VB_HelpID = 2010
   RowHeight = m_sngRowHeight
End Property

'**
'Return or set the height of the currently selected row in the Grid.
'@param        lrow Long. Required.
'@param        sHeight Single. Required.
'@rem Note: if row is not found, property is ignored.
Public Property Let RowHeight(ByVal sHeight As Single)
   Dim i As Integer

   If sHeight = 0 Then Exit Property

On Error GoTo RowHeight_Err
      
   If Not m_sngRowHeightMin = 0 Then
      If sHeight < m_sngRowHeightMin Then sHeight = m_sngRowHeightMin
   End If

   If Not m_sngRowHeightMax = 0 Then
      If sHeight > m_sngRowHeightMax Then sHeight = m_sngRowHeightMax
   End If

   m_sngRowHeight = CInt((sHeight \ Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
   
   If m_blnShown Then
      If Ambient.UserMode Then If bEditing Then Grid1_GotFocus
      Call RecalcGrid
   End If
   
   PropertyChanged "RowHeight"

Exit Property

RowHeight_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowHeight: " & Err.Source, Err.Description
End Property

Public Property Get RowHeightHeader() As Single
Attribute RowHeightHeader.VB_HelpID = 2020
   RowHeightHeader = m_sngRowHeightFixed
End Property

'**
'Return or set the height of the header row in the Grid.
'@param        sHeight Single. Required.
Public Property Let RowHeightHeader(ByVal sHeight As Single)

On Error GoTo RowHeightHeader_Err

   m_sngRowHeightFixed = CInt((sHeight \ Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
   If Grid1.FixedRows > 0 Then
      Grid1.RowHeight(0) = sHeight
      If m_blnShown Then
         If Ambient.UserMode Then If bEditing Then Grid1_GotFocus
         Call SetGridRows
      End If
   End If
   PropertyChanged "RowHeightFixed"
   
Exit Property

RowHeightHeader_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowHeightHeader: " & Err.Source, Err.Description
End Property

Public Property Get RowHeightMin() As Single
Attribute RowHeightMin.VB_HelpID = 2030
   RowHeightMin = m_sngRowHeightMin
End Property

'**
'Return or set the minimum height of any row in the Grid, except the header row.
'@param        sHeight Single. Required.
Public Property Let RowHeightMin(ByVal sHeight As Single)

   If sHeight < 0 Then sHeight = 0

   If sHeight > 0 Then
      If m_sngRowHeight < sHeight Then m_sngRowHeight = sHeight
   End If

 On Error GoTo RowHeightMin_Err
   
   m_sngRowHeightMin = CInt(sHeight \ Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
   Grid1.RowHeightMin = m_sngRowHeightMin
   
   If m_blnShown Then
      If Ambient.UserMode Then If bEditing Then Grid1_GotFocus
      RecalcGrid
   End If
   
   PropertyChanged "RowHeightMin"

Exit Property

RowHeightMin_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowHeightMin: " & Err.Source, Err.Description

End Property

Public Property Get RowHeightMax() As Single
Attribute RowHeightMax.VB_HelpID = 2040
   RowHeightMax = m_sngRowHeightMax
End Property

'**
'Return or set the maximum height of any row in the Grid, except the header row.
'@param        sHeight Single. Required.
Public Property Let RowHeightMax(ByVal sHeight As Single)

   If sHeight < 0 Then sHeight = 0
   
   If sHeight <> 0 Then If sHeight < m_sngRowHeightMin Then sHeight = m_sngRowHeightMin
   
   If sHeight > 0 Then
      If m_sngRowHeight > sHeight Then m_sngRowHeight = sHeight
   End If

 On Error GoTo RowHeightMax_Err
   
   m_sngRowHeightMax = CInt(sHeight \ Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
      
   If m_blnShown Then
      If Ambient.UserMode Then If bEditing Then Grid1_GotFocus
      RecalcGrid
   End If
   
   PropertyChanged "RowHeightMax"

Exit Property

RowHeightMax_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowHeightMax: " & Err.Source, Err.Description

End Property


'**
'Returns a boolean indicating whether a row is visible in the Grid.
'@param        Index Long. Required.
Public Property Get RowIsVisible(ByVal index As Long) As Boolean
Attribute RowIsVisible.VB_HelpID = 2050
Attribute RowIsVisible.VB_MemberFlags = "400"
  On Error Resume Next
  RowIsVisible = IsVisible(index)
End Property

'**
'Return the distance, in pixels, from the top of the grid, to the top of the row indicated by Index.
'@param        Index Long. Required.
Public Property Get RowPosition(ByVal index As Long) As Single
Attribute RowPosition.VB_HelpID = 2060
Attribute RowPosition.VB_MemberFlags = "400"
   On Error Resume Next
   
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowPosition", "Cannot get RowPosition in Grid with no rows"
   If Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowPosition", "Cannot get RowPosition in grid with no Recordset or No Rows"
   If index < 0 Or index > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".RowPosition", MSG_INVROW
 
   If IsVisible(index) Then
      RowPosition = Grid1.RowPos((index - vScroll.Value) + Grid1.FixedRows)
   Else
      Dim hFixed As Long
      If Grid1.FixedRows > 0 Then hFixed = Grid1.RowHeight(0)
      RowPosition = hFixed + ((index - vScroll.Value) * m_sngRowHeight)
   End If
End Property

'**
'Returns the total number of rows(without header) that exist in the Grid.
Public Property Get Rows() As Long
Attribute Rows.VB_HelpID = 2070
Attribute Rows.VB_MemberFlags = "400"
   If m_RsFiltered Is Nothing Then
      Rows = 0
      Exit Property
   End If
   Rows = m_RsFiltered.RecordCount
End Property

Public Property Get RowSel() As Long
Attribute RowSel.VB_HelpID = 2080
Attribute RowSel.VB_MemberFlags = "400"
  RowSel = m_CurrRowSel
End Property

'**
'Return or set the currently selected row in the Grid.
'@param        lRow Long. Required.
'@rem Note: when setting a row, if the row is not found, setting is ignored.
Public Property Let RowSel(ByVal lRow As Long)

   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowSel", "Cannot get RowPosition in Grid with no rows"
   If Not HasRecords Then Err.Raise ERR_NORS, Ambient.DisplayName & ".RowSel", "Cannot set RowSel in grid with no Recordset or No Rows"
   If lRow < 0 Or lRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".RowSel", MSG_INVROW

On Error GoTo RowSel_Err
   
   m_CurrRowSel = lRow
   RaiseEvent SelChange
   
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
      Call CalcPaintedArea
      On Error Resume Next
      If m_blnHasFocus Then Grid1.SetFocus
   End If
   
Exit Property

RowSel_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".RowSel: " & Err.Source, Err.Description
End Property

Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_HelpID = 2090
  ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_HelpID = 2100
  ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_HelpID = 2110
  ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
  UserControl.ScaleMode() = New_ScaleMode
  PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_HelpID = 2120
  ScaleTop = UserControl.ScaleTop
End Property

Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_HelpID = 2130
  ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScrollBars() As nfgScrollBarSettings
Attribute ScrollBars.VB_HelpID = 2140
  ScrollBars = m_eScrollBars
End Property

Public Property Let ScrollBars(ByVal eScroll As nfgScrollBarSettings)
  m_eScrollBars = eScroll
  hScroll.Visible = ((m_eScrollBars = nfgScrollHorizontal) Or (m_eScrollBars = nfgScrollBoth))
  vScroll.Visible = ((m_eScrollBars = nfgScrollVertical) Or (m_eScrollBars = nfgScrollBoth))
  m_bln_NeedHorzScroll = ((m_eScrollBars = nfgScrollHorizontal) Or (m_eScrollBars = nfgScrollBoth))
  m_bln_NeedVertScroll = ((m_eScrollBars = nfgScrollVertical) Or (m_eScrollBars = nfgScrollBoth))
  If m_blnShown Then SetGridRows
  PropertyChanged "ScrollBars"
End Property

Public Property Get SelectionMode() As nfgSelectionModeSettings
Attribute SelectionMode.VB_HelpID = 2150
  SelectionMode = Grid1.SelectionMode
End Property

'**
'Returns or sets the way the Grid allows selections to be made.
'@param        fgSelectionSetting Integer. Required. A valid member of the nfgSelectionModeSettings enumeration.
Public Property Let SelectionMode(ByVal fgSelectionSetting As nfgSelectionModeSettings)
  Grid1.SelectionMode() = fgSelectionSetting
  PropertyChanged "SelectionMode"
End Property

Public Property Get ShowHeaderRow() As Boolean
Attribute ShowHeaderRow.VB_HelpID = 2160
  ShowHeaderRow = m_blnHeaderRow
End Property

'**
'Returns or sets whether the Grid will display a header row.
'@param        bValue Boolean. Required.
Public Property Let ShowHeaderRow(ByVal bValue As Boolean)
   Dim i As Integer
   Dim prevScroll As Long

 On Error GoTo ShowHeaderRow_Err
   
   If m_blnShown Then If bEditing Then Grid1_GotFocus
      
   If m_blnHeaderRow <> bValue Then
      m_blnHeaderRow = bValue
      If m_blnHeaderRow = True Then
         If Ambient.UserMode = True Then
            Grid1.Redraw = False
            m_blnIgnoreSel = True
            m_blnLoading = True
            Grid1.AddItem "", 0
            Grid1.FixedRows = 1
            Grid1.Row = 0
            Grid1.Col = Grid1.FixedCols
            Grid1.colSel = Grid1.Cols - 1
            Grid1.FillStyle = flexFillRepeat
            Grid1.CellFontBold = True
            Grid1.FillStyle = flexFillSingle
            SetGridRows
            UseFieldNamesAsHeader = m_blnUseFieldNamesAsHeader
            Call FillTextmatrix(GetScrollValue())
            m_blnIgnoreSel = False
            m_blnLoading = False
            Call CalcPaintedArea
            Grid1.Redraw = m_blnRedraw
         Else
            Call BuildGridFromColumns(m_ntColumns)
         End If
      Else
         If Grid1.FixedRows > 0 Then
            Grid1.FixedRows = 0
            If Ambient.UserMode Then
               Grid1.Redraw = False
               If m_RsFiltered.RecordCount < Grid1.Rows Then
                  Grid1.RemoveItem 0
                  SetTotalPositions
               Else
                  If vScroll.Value = vScroll.Max Then
                     Grid1.RemoveItem 0
                     SetGridRows
                     FillTextmatrix vScroll.Max
                  Else
                     Grid1.RemoveItem 0
                     SetGridRows
                     FillTextmatrix GetScrollValue()
                  End If
               End If
               Call CalcPaintedArea
               Grid1.Redraw = m_blnRedraw
               On Error Resume Next
               If m_blnHasFocus Then Grid1.SetFocus
               On Error GoTo 0
            Else
               Call BuildGridFromColumns(m_ntColumns)
            End If
         End If
      End If
      PropertyChanged "FixedHeaderRow"
   End If

Exit Property

ShowHeaderRow_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ShowHeaderRow: " & Err.Source, Err.Description
End Property

Public Property Get ShowRecordSelectors() As Boolean
Attribute ShowRecordSelectors.VB_HelpID = 2170
   ShowRecordSelectors = (Grid1.FixedCols = 1)
End Property

'**
'Returns or sets whether the Grid will display recordselectors as the first column.
'@param        bValue Boolean. Required.
Public Property Let ShowRecordSelectors(ByVal bValue As Boolean)
   
   If Ambient.UserMode Then Err.Raise 387

On Error GoTo ShowRecordSelectors_Err
         
   Dim prevwidth As Long
   m_blnRecordSelectors = bValue
   If bValue <> (Grid1.FixedCols = 1) Then
      If bValue = True Then
         Grid1.Cols = Grid1.Cols + 1
         Grid1.FixedCols = 1
         Grid1.ColWidth(0) = m_sngRecordSelectorWidth
      Else
         prevwidth = Grid1.ColWidth(1)
         Grid1.FixedCols = 0
         Grid1.Cols = Grid1.Cols - 1
         Grid1.ColWidth(0) = prevwidth
      End If
   End If
   PropertyChanged "RecordSelectors"

Exit Property

ShowRecordSelectors_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ShowRecordSelectors: " & Err.Source, Err.Description

End Property

Public Property Get ShowTotals() As Boolean
Attribute ShowTotals.VB_HelpID = 2180
  ShowTotals = m_blnTotalRow
End Property

'**
'Returns or sets whether the Grid will display the totals for columns at the bottom of the Grid.
'If enabled, the Grid will keep a running total for any columns with their ShowTotals property
'set to true, based on the currently filtered data.
'@param        bValue Boolean. Required.
'@rem Note: each column must have its ShowTotals property set for the total to be visible.
Public Property Let ShowTotals(ByVal bValue As Boolean)

   On Error GoTo ShowTotals_Err
         
   m_blnTotalRow = bValue
   
   If Ambient.UserMode And m_blnShown Then
      If bEditing Then Grid1_GotFocus
      If (Not IsUnbound) And HasRecords Then
         Grid1.Redraw = False
         Call SetGridRows
         Call FillTextmatrix(GetScrollValue())
         Grid1.Redraw = m_blnRedraw
      End If
   End If
   Call SetTotalPositions
   PropertyChanged ("TotalRow")

Exit Property

ShowTotals_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".ShowTotals: " & Err.Source, Err.Description
End Property

Public Property Get ShowTotalPosition() As nfgTotalPosition
Attribute ShowTotalPosition.VB_HelpID = 2190
  ShowTotalPosition = m_intTotalFloat
End Property

'**
'Returns or sets whether the Grid will display the totals for columns at the bottom of the Grid.
'If enabled, the Grid will keep a running total for any columns with their ShowTotals property
'set to true, based on the currently filtered data.
'@param        bValue Boolean. Required.
'@rem Note: each column must have its ShowTotals property set for the total to be visible.
Public Property Let ShowTotalPosition(ByVal nValue As nfgTotalPosition)
On Error Resume Next
  m_intTotalFloat = nValue
  Call SetTotalPositions
  PropertyChanged ("TotalFloat")
End Property

'**
'Returns the text from the currently selected cell in the Grid as a string.
Public Property Get Text() As String
Attribute Text.VB_HelpID = 2200
Attribute Text.VB_MemberFlags = "400"
   m_RsFiltered.AbsolutePosition = m_CurrRow + 1
   Text = m_ntColumns(m_CurrCol).FormatValue(m_RsFiltered.Fields(m_ntColumns(m_CurrCol).Name).Value)
End Property

Public Property Get TextStyle() As nfgTextStyleSettings
Attribute TextStyle.VB_HelpID = 2210
  TextStyle = Grid1.TextStyle
End Property

'**
'Returns or sets the style of the text displayed in the Grid.
'@param        fgTextStyle Integer. Required. A valid member of the TextStyleSettings enumeration.
Public Property Let TextStyle(ByVal fgTextStyle As nfgTextStyleSettings)
  Grid1.TextStyle = fgTextStyle
  PropertyChanged "TextStyle"
End Property

Public Property Get TextStyleHeader() As nfgTextStyleSettings
Attribute TextStyleHeader.VB_HelpID = 2220
  TextStyleHeader = Grid1.TextStyleFixed
End Property

'**
'Returns or sets the style of the text displayed in the Grid header.
'@param        fgTextStyle Integer. Required. A valid member of the TextStyleSettings enumeration.
'@rem Note: Only applies if ShowHeader property is true.
Public Property Let TextStyleHeader(ByVal fgTextStyle As nfgTextStyleSettings)
  Grid1.TextStyleFixed = fgTextStyle
  PropertyChanged "TextStyleFixed"
End Property

Public Property Get TopRow() As Long
Attribute TopRow.VB_HelpID = 2230
Attribute TopRow.VB_MemberFlags = "400"

   If IsUnbound Or (Not m_bln_NeedVertScroll) Then
      TopRow = 0
   Else
      TopRow = GetScrollValue
   End If

End Property

'**
'Return or set the current top row in the Grid.
'@param        lRow Long. Required.
Public Property Let TopRow(ByVal lRow As Long)

   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Or (Not HasRecords) Then Err.Raise ERR_NORS, Ambient.DisplayName & ".TopRow", "Cannot set TopRow when grid has no records"
   If Not m_bln_NeedVertScroll Then Exit Property
   
On Error GoTo Row_Err
      
   If bEditing Then Grid1_GotFocus
   
   If lRow < 0 Then lRow = 0
   If lRow > vScroll.Max Then lRow = vScroll.Max
   
   SetScrollValue lRow
   vScroll.Value = lRow
   
   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus

Exit Property

Row_Err:
   Err.Raise Err.Number, UserControl.Name & ".TopRow: " & Err.Source, Err.Description

End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_HelpID = 2240
  ToolTipText = Grid1.ToolTipText
End Property

'**
'Returns or sets the text displayed when the user rests their mouse over the grid.
'@param        sText String. Required.
Public Property Let ToolTipText(ByVal sText As String)
  Grid1.ToolTipText() = sText
  PropertyChanged "ToolTipText"
End Property

Public Property Get UnderlyingFilter() As Variant
Attribute UnderlyingFilter.VB_HelpID = 2250
Attribute UnderlyingFilter.VB_MemberFlags = "400"
   
   UnderlyingFilter = 0
   If m_RSMaster Is Nothing Then Exit Property
   If m_RSMaster.State <> 1 Then Exit Property

   UnderlyingFilter = m_RSMaster.Filter

End Property

'**
'Returns or sets an ADO Recordset Filter to the underlying recordset.
'This limits the data the user can filter and sort on to only the
'data that matches the filter.
'@param        vFilter Variant. Required. A valid ADODB Recordset filter. Applies to the underlying, unformatted data.
'@rem Note: This filter and the FormattedFilter property are independent of each other. Any filter
'applied using this property will still be present in the recordset regardless of user changes
'until removed by setting the filter to 0 (adFilterNone).
Public Property Let UnderlyingFilter(ByVal vFilter As Variant)
  Dim vSort As Variant

On Error GoTo Filter_Err

   If Not Ambient.UserMode Then Err.Raise 380
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".TopRow", "Cannot set UnderlyingFilter when grid has no recordset"
      
   If vFilter = "0" Then vFilter = 0

   If vFilter = 0 Or Len(vFilter) = 0 Then
      If m_RSMaster.Filter = 0 Or Len(m_RSMaster.Filter) = 0 Then Exit Property
   End If

   m_RsFiltered.Filter = vFilter
   m_RSMaster.Filter = vFilter

   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
      GridMousePointer = vbHourglass
      Screen.MousePointer = vbHourglass
      Grid1.Redraw = False
      Call SetGridRows
      m_blnLoading = True
      vScroll.Value = 0
      FillTextmatrix 0
      Grid1.colSel = Grid1.Col
      Grid1.Row = Grid1.FixedRows
      Grid1.RowSel = Grid1.FixedRows
      m_CurrColSel = m_CurrCol
      m_CurrRow = 0
      m_CurrRowSel = 0
      Grid1.FocusRect = m_intFocusRect
      Grid1.Highlight = m_intHighlight
      m_blnLoading = False
      Grid1.Redraw = m_blnRedraw
      GridMousePointer = vbDefault
      Screen.MousePointer = vbDefault
      On Error Resume Next
      m_LastRow = m_CurrRow
      If m_blnHasFocus Then Grid1.SetFocus
   End If
   
Exit Property

Filter_Err:
   Err.Raise Err.Number, UserControl.Name & ".UnderlyingFilter: " & Err.Source, Err.Description

End Property

Public Property Get UseFieldNamesAsHeader() As Boolean
Attribute UseFieldNamesAsHeader.VB_HelpID = 2260
   UseFieldNamesAsHeader = m_blnUseFieldNamesAsHeader
End Property

'**
'Returns or sets whether the Grid will use the field names from the recordset
'as the header text for the columns.
'@param        bValue Boolean. Required.
Public Property Let UseFieldNamesAsHeader(ByVal bValue As Boolean)
   Dim i As Integer
  
 On Error GoTo FieldName_Err
    
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
   
      If m_blnHeaderRow Then

      If bValue = True Then

         If Ambient.UserMode = True Then
   
               For i = 0 To m_ntColumns.count - 1
                  If m_ntColumns(i).Visible = True Then
                     Grid1.TextMatrix(0, i + Grid1.FixedCols) = m_ntColumns(i).Name
                  End If
               Next i
   
            Else
               Call BuildGridFromColumns(m_ntColumns)
            End If
   
         Else
   
            If Ambient.UserMode = True Then
               For i = 0 To m_ntColumns.count - 1
                  If m_ntColumns(i).Visible = True Then
                     Grid1.TextMatrix(0, i + Grid1.FixedCols) = m_ntColumns(i).HeaderText
                  End If
               Next i
            Else
               Call BuildGridFromColumns(m_ntColumns)
            End If
   
         End If
       
         m_blnUseFieldNamesAsHeader = bValue
         PropertyChanged "UseFieldNamesAsHeader"
      
      End If
      
   End If

Exit Property

FieldName_Err:
   Err.Raise Err.Number, UserControl.Name & ".UseFieldNamesAsheader: " & Err.Source, Err.Description

End Property

Public Property Get vScrollBar() As ntFxGdScrollBar
Attribute vScrollBar.VB_HelpID = 2270
Attribute vScrollBar.VB_MemberFlags = "400"
   Set vScrollBar = m_vScroll
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_HelpID = 2280
  WhatsThisHelpID = Grid1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
  Grid1.WhatsThisHelpID() = New_WhatsThisHelpID
  PropertyChanged "WhatsThisHelpID"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_HelpID = 2290
  WordWrap = Grid1.WordWrap
End Property

'**
'This property determines if text will wrap to the next line if it is too long for the cell
'@param bValue Boolean. Required.
Public Property Let WordWrap(ByVal bValue As Boolean)
  Grid1.WordWrap() = bValue
  PropertyChanged "WordWrap"
End Property

'BEGIN METHODS ******************************************************************************************

'**
'Applies a filter to the formatted version of the Grid recordset.
'Filters may be stacked, where the application of one filter does not remove the previous
'filter, but instead, filters within the currently displayed data.
'@param        bApplyToExisting Boolean. Optional. Indicate whether to add filter to previous(if any), or replace previous with new filter.
'@param        vFilter Variant. Required. A valid ADO recordset filter, using the formatted data in the grid to match.
'@rem Note: This filter will look for matches in the currently filtered and sorted data,
'if you wish to filter the underlying recordset, use the UnderlyingFilter property instead.
Public Sub ApplyFormattedFilter(ByVal FieldName As String, ByRef arrValues() As String, _
                               ByVal bIncludeRecords As Boolean, _
                               Optional ByVal bApplyToExisting As Boolean = False)
Attribute ApplyFormattedFilter.VB_HelpID = 1030
   Dim vSort As Variant
   Dim arrFilter() As String
   Dim vCurFilter As Variant
   Dim bOK As Boolean
   Dim PrevSort As String
   
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".ApplyFormattedFilter", "You cannot Apply a Filter while the grid has no recordset."
   If Not m_ntColumns.Exists(FieldName) Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ApplyFormattedFilter", MSG_INVCOL
   If m_RSMaster.RecordCount = 0 Then Exit Sub
   
 On Error GoTo Filter_Err
  
   If m_blnShown Then
      If bEditing Then Grid1_GotFocus
      Grid1.Redraw = False
      
      PrevSort = m_RsFiltered.Sort
      
      If bApplyToExisting = False Then
         Call GetUnFilteredRecords(False)
         RaiseEvent OnFilterRemove
      End If
      Call AddFilter(FieldName, arrValues, bIncludeRecords)
      
      m_RsFiltered.Sort = PrevSort
      
      Dim nCol As ntColumn
   
      For Each nCol In m_ntColumns
         If nCol.ShowTotal Then
            Call CalcTotals
            Exit For
         End If
      Next
         
      Grid1.Redraw = m_blnRedraw
   End If
   
Exit Sub

Filter_Err:
   Err.Raise vbObjectError + 3009, Ambient.DisplayName & ".ApplyFormattedFilter", "Invalid Filter"
End Sub

Private Function AddFilter(ByVal FieldName As String, ByRef arrValues() As String, ByVal bInclude As Boolean) As Boolean
On Error Resume Next
   Dim nf As ntFilter
   Dim arrVals() As String
   Dim PrevSort As String
   
   AddFilter = False
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Grid1.Redraw = False
   m_blnLoading = True
   m_blnManualScroll = True
   If ApplyFilter(FieldName, arrValues, bInclude) Then
      AddFilter = True
      Call SetGridRows
      vScroll.Value = 0
      FillTextmatrix 0
      Grid1.Col = m_CurrCol + Grid1.FixedCols
      Grid1.colSel = Grid1.Col
      Grid1.Row = Grid1.FixedRows
      Grid1.RowSel = Grid1.FixedRows
      m_CurrColSel = m_CurrCol
      m_CurrRow = 0
      m_CurrRowSel = 0
      Grid1.FocusRect = m_intFocusRect
      Grid1.Highlight = m_intHighlight
      m_LastRow = m_CurrRow
      Set nf = m_colFilters.NewFilter
      nf.FieldName = FieldName
      nf.Include = True
      arrVals = arrValues
      nf.Values = arrVals
      RaiseEvent OnFilter(nf)
      Set nf = Nothing
   End If
   m_blnManualScroll = False
   m_blnLoading = False
   Grid1.Redraw = m_blnRedraw
   If m_blnHasFocus Then Grid1.SetFocus
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Function

Public Sub AddBlankRows(ByVal lNumRowsToAdd As Long)
Attribute AddBlankRows.VB_HelpID = 1040

   Dim i As Long
   Dim l As Long
   Dim arrFields() As Variant
   Dim arrVals() As Variant
      
   If Not Ambient.UserMode Then Err.Raise 380
   If m_ntColumns.count = 0 Then Err.Raise ERR_INVCOL + 2, Ambient.DisplayName & ".AddBlankRows", "Cannot add Blank Rows to grid with no Columns"
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".AddBlankRows", "No recordset available to add row"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".AddBlankRows", "Display must be called prior to adding rows"
   
On Error GoTo Add_Err
   
   If bEditing Then Grid1_GotFocus
   
   ReDim arrFields(m_RsFiltered.Fields.count - 1)
   ReDim arrVals(m_RsFiltered.Fields.count - 1)
      
   If IsUnbound Then
      For i = 0 To m_RsFiltered.Fields.count - 1
         arrFields(i) = m_RsFiltered.Fields(i).Name
         If m_ntColumns(m_RsFiltered.Fields(i).Name).ColFormat = nfgBooleanCheckBox Or _
            m_ntColumns(m_RsFiltered.Fields(i).Name).ColFormat = nfgBooleanTrueFalse Then
            arrVals(i) = "False"
         Else
            arrVals(i) = vbNullString
         End If
      Next i
   End If
   
   For i = 1 To lNumRowsToAdd
      If IsUnbound Then
         m_RsFiltered.AddNew arrFields, arrVals
         m_RSMaster.AddNew arrFields, arrVals
      Else
         m_RsFiltered.AddNew
         m_RSMaster.AddNew
      End If
   Next i
   
   Grid1.Redraw = False
   
   Call SetGridRows
   Call SetScrollValue(GetScrollValue)

   If m_blnRedraw = True Then
      Grid1.Redraw = True
      On Error Resume Next
      If m_blnHasFocus Then Grid1.SetFocus
   End If

Exit Sub

Add_Err:
   If Err.Number = -2147217887 Then
      Err.Raise Err.Number, Ambient.DisplayName & ".AddBlankRows: " & Err.Source, "Microsoft Ambigous Error. " & vbCrLf & Err.Description
   Else
      Err.Raise Err.Number, Ambient.DisplayName & ".AddBlankRow: " & Err.Source, Err.Description
   End If

End Sub

'**
'Method will add a row to the Grid and the underlying recordset.
'$EOL$
'$EOL$
Public Function AddRow(ByRef vFields() As Variant, ByRef vValues() As Variant, _
                       Optional ByVal bSelectRow As Boolean = False) As Variant
Attribute AddRow.VB_HelpID = 1050
   Dim i As Long
   Dim lonRowData As Long
   Dim vBkmrk As Variant
   Dim lonFRecords As Long
   Dim lonOrgRecords As Long
   Dim lonID As Long
   Dim nCol As ntColumn

   If Not Ambient.UserMode Then Err.Raise 380
   If m_ntColumns.count = 0 Then Err.Raise ERR_INVCOL + 2, Ambient.DisplayName & ".AddRow", "Cannot Add Rows to grid with no Columns"
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".AddRow", "No recordset available to add row"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".AddRow", "Display must be called prior to adding rows"
   If UBound(vFields) <> UBound(vValues) Then Err.Raise vbObjectError + 5055, Ambient.DisplayName & ".AddRow", "Arrays must have same number of elements"
         
On Error GoTo Add_Err
   
   If m_blnRedraw Then If bEditing Then Grid1_GotFocus
      
   Dim lRecords As Long
   lRecords = m_RSMaster.RecordCount
   
   m_RSMaster.AddNew vFields, vValues
   
   If m_RSMaster.RecordCount <> lRecords + 1 Then GoTo Add_Err
   
On Error GoTo Add_Err

   m_RsFiltered.AddNew vFields, vValues
   
   vBkmrk = m_RsFiltered.Bookmark
   lonID = m_RsFiltered.AbsolutePosition
   
   For Each nCol In m_ntColumns
      If nCol.ShowTotal Then
         If IsNumeric(m_RsFiltered.Fields(nCol.Name).Value) Then
            lbltotal(nCol.index).Text = _
               CCur(lbltotal(nCol.index).Text) + CCur(m_RsFiltered.Fields(nCol.Name).Value & "")
         End If
      End If
   Next nCol

   Err.Clear

On Error GoTo Add_Err

   Grid1.Redraw = False
    
   Call SetGridRows
     
   m_blnLoading = True
   
   If Not m_bln_NeedVertScroll Then
       lonID = 0
   Else
      Call SetScrollValue(lonID - 1)
      If lonID - 1 > vScroll.Max Then
         lonID = vScroll.Max
      Else
         lonID = lonID - 1
      End If
      vScroll.Value = lonID
      m_PrevVertValue = lonID
   End If
   
   m_blnLoading = False
   
   FillTextmatrix lonID
     
   If bSelectRow Then
      m_CurrCol = 0
      m_CurrColSel = (Grid1.Cols - 1) - Grid1.FixedCols
   Else
      m_CurrCol = 0
      m_CurrColSel = 0
   End If
      
   m_CurrRow = RowFromBookmark(vBkmrk)
   m_CurrRowSel = m_CurrRow
      
   CalcPaintedArea
   
   Grid1.Redraw = m_blnRedraw
   
   DoEvents
   
   RaiseEvent ColChange(m_ntColumns(m_CurrCol))
   RaiseEvent RowChange(lonID)
     
   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus

Exit Function

Add_Err:
   If Err.Number = -2147217887 Then
      Err.Raise Err.Number, Ambient.DisplayName & ".AddRow: " & Err.Source, "Microsoft Ambigous Error. " & vbCrLf & Err.Description
   Else
      Err.Raise Err.Number, Ambient.DisplayName & ".AddRow: " & Err.Source, Err.Description
   End If
End Function

Private Function ApplyFilter(ByVal FieldName As String, ByRef arrValues() As String, ByVal bInclude As Boolean) As Boolean
   Dim l As Long
   Dim strFilters() As String
   Dim bDelete As Boolean
   Dim bFilterNulls As Boolean
   Dim bCancel As Boolean
      
   ApplyFilter = False
     
   strFilters = arrValues

   Dim nf As ntFilter

   Set nf = m_colFilters.NewFilter

   nf.FieldName = FieldName
   nf.Include = Abs(bInclude)
   nf.Values = strFilters
     
   bCancel = False
   
   RaiseEvent BeforeFilter(nf, bCancel)
   
   If bCancel Then
      Set nf = Nothing
      Exit Function
   End If
     
   For l = 0 To UBound(arrValues)
      If arrValues(l) = "NULL" Or Len(arrValues(l)) = 0 Then
         bFilterNulls = True
         Exit For
      End If
   Next l
   
   If bInclude Then

      m_RsFiltered.moveLast

      Do While Not m_RsFiltered.BOF
         bDelete = True
         If IsNull(m_RsFiltered.Fields(FieldName).Value) Or _
               Len(m_RsFiltered.Fields(FieldName).Value) = 0 Then
               bDelete = Not bFilterNulls
         Else
            For l = 0 To UBound(arrValues)
               If StrComp(arrValues(l), m_RsFiltered.Fields(FieldName).Value, vbTextCompare) = 0 Then
                  bDelete = False
                  Exit For
               End If
            Next l
         End If
         If bDelete Then m_RsFiltered.Delete
         m_RsFiltered.MovePrevious
      Loop

   Else

      m_RsFiltered.moveLast

      Do While Not m_RsFiltered.BOF
         bDelete = False
         If IsNull(m_RsFiltered.Fields(FieldName).Value) Or _
            Len(m_RsFiltered.Fields(FieldName).Value) = 0 Then
            bDelete = bFilterNulls
         Else
            For l = 0 To UBound(arrValues)
               If StrComp(arrValues(l), m_RsFiltered.Fields(FieldName).Value, vbTextCompare) = 0 Then
                  bDelete = True
                  Exit For
               End If
            Next l
         End If
         If bDelete Then m_RsFiltered.Delete
         m_RsFiltered.MovePrevious
      Loop

   End If
      
   m_colFilters.Add nf, "Filter" & m_colFilters.count

   m_RsFiltered.Filter = m_RSMaster.Filter

   Dim nCol As ntColumn

   For Each nCol In m_ntColumns
      If nCol.ShowTotal Then
         Call CalcTotals
         Exit For
      End If
   Next
         
   Set nf = Nothing
   
   ApplyFilter = True
   
End Function

Private Function BuildGridRSFilterValues(ByVal vColKey As Variant, _
                                    ByVal pRow As Long, ByVal pRowSel As Long, _
                                    Optional ByVal pInclude As Boolean = True) As String()

   Dim strFieldName As String
   Dim lonRows As Long
   Dim pFromRow As Long
   Dim pToRow As Long
   Dim pCompare As String
   Dim strFilters() As String
   Dim i As Integer
   Dim fld As Object

  ' Safety valve only, unlikely to happen
  If (m_RsFiltered Is Nothing) Then Exit Function
  If m_RsFiltered.RecordCount = 0 Then Exit Function
  If m_ntColumns.count = 0 Then Exit Function

  If pRowSel >= pRow Then
    pFromRow = pRow
    pToRow = pRowSel
  Else
    pFromRow = pRowSel
    pToRow = pRow
  End If

   ReDim strFilters(pToRow - pFromRow)

   With m_RsFiltered

      Set fld = .Fields(m_ntColumns(vColKey).Name)

      strFieldName = fld.Name

      For lonRows = pFromRow To pToRow

         .AbsolutePosition = lonRows + 1

         If m_ntColumns(vColKey).Visible = True Then

            If IsNull(fld.Value) Then
               strFilters(lonRows - pFromRow) = "NULL"
            Else
               strFilters(lonRows - pFromRow) = CStr(fld.Value)
            End If

         End If

      Next lonRows

   End With

  ' Remove any duplicate rows
  Call RemoveStringArrayDupes(strFilters)
  BuildGridRSFilterValues = strFilters

End Function

'**
'This method will determine the best fit for the text in each column, and size the column appropriately.
'It will then determine if the columns are smaller than the Grid, and size them to the grid.
'@param        vKey Variant. Optional. The index or key of a column.
'If left out, all columns will be resized. Default value is False.
Public Sub AutosizeGridColumns(Optional ByVal vKey As Variant = -1)
Attribute AutosizeGridColumns.VB_HelpID = 1060
   Dim l As Long
   Dim j As Long
   Dim lonRecords As Long
   Dim intStep As Integer
   Dim lonTextWidth As Integer
   Dim lCol As Long
   Dim strFieldVal As String
   Dim arrWidths() As Long
   Dim arrOffsets() As Integer
   Dim arrFields() As Long
   Dim intCols As Integer
   Dim intColModifier As Integer
   Dim lonTextOffset As Integer
   
   If Not Ambient.UserMode Then Err.Raise 380
   If m_ntColumns.count = 0 Then Err.Raise ERR_INVCOL + 2, Ambient.DisplayName & ".AutosizeGridColumns", "Cannot Autosize Columns in grid with no Columns"
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".AutosizeGridColumns", "Invalid recordset"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".AutosizeGridColumns", "Display must be called prior to adding rows"
   
On Error GoTo AutoSize_Err
       
   If m_blnShown Then If bEditing Then Grid1_GotFocus
   
   SaveRestorePrevGrid Grid1
   
   Const lonPadding = 300
   
   If vKey <> -1 Then
      If Not m_ntColumns.Exists(vKey) Then Exit Sub
      If Not m_ntColumns(vKey).Visible Then Exit Sub
   Else
      If m_ntColumns.count = 0 Then Exit Sub
   End If

   With m_RsFiltered

      'If a vKey was passed in we are only doing one column
      If vKey <> -1 Then
         ReDim arrWidths(0)
         ReDim arrFields(0)
         ReDim arrOffsets(0)
         For j = 0 To .Fields.count - 1
            If .Fields(j).Name = m_ntColumns(vKey).Name Then
               arrFields(0) = j
               If m_blnHeaderRow Then
                  arrWidths(0) = TextWidth(m_ntColumns(.Fields(j).Name).HeaderText) + lonPadding
               Else
                  arrWidths(0) = 0
               End If
               Exit For
            End If
         Next j
      Else  ' We are sizing all the Columns - build indexed arrays
         ' Add the fields to the base filtered recordsets
         For j = 0 To .Fields.count - 1
            If m_ntColumns.Exists(.Fields(j).Name) Then
               If m_ntColumns(.Fields(j).Name).Visible Then
                  ReDim Preserve arrWidths(intCols)
                  ReDim Preserve arrFields(intCols)
                  ReDim Preserve arrOffsets(intCols)
                  If m_blnHeaderRow Then
                     arrWidths(intCols) = TextWidth(m_ntColumns(.Fields(j).Name).HeaderText) + lonPadding
                  Else
                     arrWidths(intCols) = 0
                  End If
                  arrFields(intCols) = j
                  intCols = intCols + 1
               End If
            End If
         Next j
      End If
      
      If HasRecords Then
               
         .movefirst

         'If we have less than 1000 records, check them all
         ' Else start with first 500
         If .RecordCount < 1000 Then
            lonRecords = .RecordCount
         Else
            lonRecords = 500
         End If
   
         For l = 1 To lonRecords
   
            For j = 0 To UBound(arrFields)
                         
               ' The lonPadding constant compensates for text insets
               ' You can adjust this value above as desired.
               lonTextWidth = TextWidth(.Fields(arrFields(j)) & "") + lonPadding
                  
               ' Reset lonBiggestWidth to the intMaxColWidth value if necessary
               If lonTextWidth > arrWidths(j) Then arrWidths(j) = lonTextWidth
                                            
               If l < 25 Then
                  lonTextOffset = TextWidth(m_ntColumns(.Fields(arrFields(j)).Name).FormatValue(.Fields(arrFields(j)).Value))
                  lonTextOffset = lonTextOffset - lonTextWidth
                  If lonTextOffset > arrOffsets(j) Then arrOffsets(j) = lonTextOffset
               End If
                  
            Next j
   
            .MoveNext
   
         Next l
   
         If lonRecords < .RecordCount Then
            intColModifier = CInt(((Grid1.Cols - Grid1.FixedCols) \ 10) + 1) * 10
            intStep = (.RecordCount - lonRecords) \ (100000 \ intColModifier)
            If intStep < 1 Then intStep = 1
   
            Do While Not .EOF
   
               For j = 0 To UBound(arrFields)
   
                  ' The lonPadding constant compensates for text insets
                  ' You can adjust this value above as desired.
                  lonTextWidth = TextWidth(.Fields(arrFields(j)) & "") + lonPadding
   
                  ' Reset lonBiggestWidth to the intMaxColWidth value if necessary
                  If lonTextWidth > arrWidths(j) Then arrWidths(j) = lonTextWidth
   
               Next j
   
               .Move intStep
   
            Loop
   
         End If
      
      End If
      
      For j = 0 To UBound(arrWidths)
         arrWidths(j) = arrWidths(j) + arrOffsets(j)
      Next j
      
      For j = 0 To UBound(arrFields)

         lCol = m_ntColumns(.Fields(arrFields(j)).Name).index + Grid1.FixedCols

         ' Set the width of the column to match largest text
         ' Check property settings for Min and Max Col Widths
         If arrWidths(j) < m_sngMinColWidth Then
            Grid1.ColWidth(lCol) = m_sngMinColWidth
         ElseIf arrWidths(j) > m_sngMaxColWidth Then
            If m_sngMaxColWidth <> 0 Then
               Grid1.ColWidth(lCol) = m_sngMaxColWidth
            Else
               Grid1.ColWidth(lCol) = arrWidths(j)
            End If
         Else
            Grid1.ColWidth(lCol) = arrWidths(j)
         End If

         m_ntColumns(.Fields(arrFields(j)).Name).Width = Grid1.ColWidth(lCol)

      Next j

   End With

On Error GoTo FitCol_Err

   If vKey = -1 And m_blnAutoSizeColumns = True Then Call FitColumnsToGrid
     
   Call SetGridRows
   
   SaveRestorePrevGrid Grid1, False
   
Exit Sub

FitCol_Err:
   If Err.Number = -2147217887 Then
      Err.Raise Err.Number, Ambient.DisplayName & ".FitColumnsToGrid: " & Err.Source, "Microsoft Ambigous Error. " & vbCrLf & Err.Description
   Else
      Err.Raise Err.Number, Ambient.DisplayName & ".FitColumnsToGrid: " & Err.Source, Err.Description
   End If

Exit Sub

AutoSize_Err:
   If Err.Number = -2147217887 Then
      Err.Raise Err.Number, Ambient.DisplayName & ".AutosizeGridColumns: " & Err.Source, "Microsoft Ambigous Error. " & vbCrLf & Err.Description
   Else
      Err.Raise Err.Number, Ambient.DisplayName & ".AutosizeGridColumns: " & Err.Source, Err.Description
   End If

End Sub

Public Function BookmarkFromRow(Optional ByVal RowIndex As Long = -1) As Variant
Attribute BookmarkFromRow.VB_HelpID = 1070
   
   BookmarkFromRow = -1
   
   If IsUnbound Then Exit Function
   If Not HasRecords Then Exit Function
   If RowIndex < 0 Or RowIndex > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".BookmarkFromRow", MSG_INVROW
   
On Error GoTo RFB_Err

   If RowIndex = -1 Then RowIndex = Grid1.Row - Grid1.FixedRows
   m_RsFiltered.AbsolutePosition = RowIndex + 1
   BookmarkFromRow = m_RsFiltered.Bookmark

Exit Function

RFB_Err:
   BookmarkFromRow = -1
End Function

Friend Sub BuildGridFromColumns(ByVal mCols As ntColumns)

   Dim nCol As ntColumn
   Dim l As Long

   Call ClearGrid(True)

   If mCols Is Nothing Then Exit Sub

   If mCols.count = 0 Then Exit Sub

   With Grid1

      If .FixedCols > 0 Then .ColWidth(0) = m_sngRecordSelectorWidth

      .Cols = .FixedCols + mCols.count
                  
      For l = 0 To mCols.count - 1
                 
         If l > lbltotal.UBound Then
            Load lbltotal(l)
            lbltotal(l).Text = ""
         End If

         Set nCol = mCols(l)

         .ColAlignment(l + .FixedCols) = nCol.TextAlignment
         .ColAlignmentFixed(l + .FixedCols) = nCol.TextAlignmentHeader

         If m_blnHeaderRow = True Then
            .Row = 0
            .RowHeight(0) = m_sngRowHeightFixed
            .Col = l + .FixedCols
            .CellFontBold = True
            If nCol.Visible = True Then
               If m_colColpics.Exists(mCols(l).Name, "HDR") Then
                  .CellPictureAlignment = m_colColpics(mCols(l).Name, "HDR").PictureAlignment
                  Set .CellPicture = m_colColpics(mCols(l).Name, "HDR").CellPicture
               End If
               .TextMatrix(0, l + .FixedCols) = nCol.HeaderText
            Else
               Set .CellPicture = Nothing
              .TextMatrix(0, l + .FixedCols) = ""
            End If
         End If

          .Row = .FixedRows
          .Col = .FixedCols

         ' Verify we are not exceeding max or min widths
         If Not m_blnAutoSizeColumns = True Then
            If nCol.Width < m_sngMinColWidth Then
               If m_sngMinColWidth > 0 Then
                  nCol.Width = m_sngMinColWidth
               Else
                  nCol.Width = m_def_ColWidth
               End If
            ElseIf nCol.Width > m_sngMaxColWidth Then
               If Not m_sngMaxColWidth = 0 Then nCol.Width = m_sngMaxColWidth
            End If
         End If

         If nCol.Visible = True Then
            .ColWidth(l + .FixedCols) = nCol.Width
         Else
            .ColWidth(l + .FixedCols) = 0
         End If
                          
      Next l

      If Not Ambient.UserMode = True Then

         Dim i As Integer
         i = 0

         Do While i < 50
            .AddItem vbNullString
            .RowHeight(.Rows - 1) = m_sngRowHeight
            i = i + 1
         Loop

         .Col = .FixedCols
         .colSel = .FixedCols
         .Row = .FixedRows
         .RowSel = .FixedRows

         .FocusRect = flexFocusNone
         .Highlight = flexHighlightNever
         .ScrollBars = flexScrollBarHorizontal

      End If

   End With

End Sub

Private Function BuildGridRSSort(ByVal pCol As Integer, ByVal pColSel As Integer, _
                                 Optional ByVal SortAsc As Boolean = True) As String
  Dim lonCounter As Long
  Dim pFromCol As Long
  Dim pToCol As Long
  Dim strsort As String
  Dim strSortType As String
  Dim strFieldName As String

  ' Safety valve only, unlikely to happen
  If (m_RsFiltered Is Nothing) Then Exit Function
  If m_RsFiltered.RecordCount = 0 Then Exit Function
  If m_ntColumns.count = 0 Then Exit Function

  If SortAsc = True Then
    strSortType = " ASC"
    m_ntColumns(pFromCol).Sorted = nfgSortAsc
  Else
    strSortType = " DESC"
    m_ntColumns(pFromCol).Sorted = nfgSortDesc
  End If

   If pColSel >= pCol Then

      For lonCounter = pCol To pColSel
        If m_ntColumns(lonCounter).Visible = True Then
          strFieldName = m_ntColumns(lonCounter).Name
          If Len(strsort) = 0 Then
            strsort = strFieldName & strSortType
          Else
            strsort = strsort & ", " & strFieldName & strSortType
          End If
        End If
      Next lonCounter

   Else

      For lonCounter = pCol To pColSel Step -1
         If m_ntColumns(lonCounter).Visible = True Then
            strFieldName = m_ntColumns(lonCounter).Name
            If Len(strsort) = 0 Then
               strsort = strFieldName & strSortType
            Else
               strsort = strsort & ", " & strFieldName & strSortType
            End If
         End If
      Next lonCounter

   End If

   BuildGridRSSort = strsort

End Function

Private Sub CalcPaintedArea()
   Dim currTop As Long
   Dim currBottom As Long
   Dim blnRowIn As Boolean
   Dim blnRowSelIn As Boolean
   Dim pRow As Long
   Dim pRowSel As Long

   If Not Ambient.UserMode Then Err.Raise 380

   If Grid1.SelectionMode = flexSelectionByRow Then
      m_CurrCol = 0
      m_CurrColSel = (Grid1.Cols - 1) - Grid1.FixedCols
   ElseIf Grid1.SelectionMode = flexSelectionByColumn Then
      m_CurrRow = 0
      m_CurrRowSel = m_RsFiltered.RecordCount - 1
   End If

    'calc the top and bottom records currently visible in grid
    currTop = GetScrollValue
    
    If m_bln_NeedVertScroll Then
       currBottom = GetScrollValue() + vScroll.LargeChange - 1
    Else
       currBottom = (Grid1.Rows - 1) - Grid1.FixedRows
    End If

    blnRowIn = (m_CurrRow >= currTop And m_CurrRow <= currBottom)

    blnRowSelIn = (m_CurrRowSel <= currBottom And m_CurrRowSel >= currTop)

   Select Case True
      'Both are visible
      Case (blnRowIn And blnRowSelIn)
         Grid1.FocusRect = m_intFocusRect
         Grid1.Highlight = m_intHighlight
         pRow = (m_CurrRow - currTop) + Grid1.FixedRows
         pRowSel = (m_CurrRowSel - currTop) + Grid1.FixedRows

      'Row is in, RowSel is not
      Case (blnRowIn And Not blnRowSelIn)
         Grid1.FocusRect = m_intFocusRect
         Grid1.Highlight = m_intHighlight
         pRow = (m_CurrRow - currTop) + Grid1.FixedRows
         If m_CurrRow < m_CurrRowSel Then
            pRowSel = Grid1.Rows - 1
         Else
            pRowSel = Grid1.FixedRows
         End If

      'Row is Out, RowSel is in
      Case (Not blnRowIn And blnRowSelIn)
         Grid1.FocusRect = flexFocusNone
         Grid1.Highlight = m_intHighlight
         pRowSel = (m_CurrRowSel - currTop) + Grid1.FixedRows
         If m_CurrRow < m_CurrRowSel Then
            pRow = Grid1.FixedRows
         Else
            pRow = Grid1.Rows - 1
         End If

      'Neither is visible in the grid
      Case (Not blnRowIn And Not blnRowSelIn)
         If (m_CurrRow < currTop And m_CurrRowSel < currTop) Then
            Grid1.FocusRect = flexFocusNone
            Grid1.Highlight = flexHighlightNever
            pRow = Grid1.FixedRows
            pRowSel = Grid1.FixedRows
         ElseIf (m_CurrRow < currTop And m_CurrRowSel > currBottom) Then
            Grid1.FocusRect = flexFocusNone
            Grid1.Highlight = m_intHighlight
            pRow = Grid1.FixedRows
            pRowSel = Grid1.Rows - 1
         ElseIf (m_CurrRow > currBottom And m_CurrRowSel > currBottom) Then
            Grid1.FocusRect = flexFocusNone
            Grid1.Highlight = flexHighlightNever
            pRow = Grid1.Rows - 1
            pRowSel = Grid1.Rows - 1
         ElseIf (m_CurrRow > currBottom And m_CurrRowSel < currTop) Then
            Grid1.FocusRect = flexFocusNone
            Grid1.Highlight = m_intHighlight
            pRow = Grid1.Rows - 1
            pRowSel = Grid1.FixedRows
         End If

   End Select

   If m_CurrCol < 0 Then m_CurrCol = 0
   If m_CurrColSel < 0 Then m_CurrColSel = 0
   If m_CurrRow < 0 Then m_CurrRow = 0
   If m_CurrRowSel < 0 Then m_CurrRowSel = 0

   m_blnIgnoreSel = True
   Grid1.Row = pRow
   Grid1.Col = m_CurrCol + Grid1.FixedCols
   Grid1.RowSel = pRowSel
   Grid1.colSel = m_CurrColSel + Grid1.FixedCols
      
   m_blnIgnoreSel = False
     
   m_LastRow = m_CurrRow
   m_LastCol = m_CurrCol
     
End Sub

Private Function CalcTotals()
   Dim lonRowCounter As Long
   Dim arrTotalFields() As Long
   Dim arrIndexes() As Integer
   Dim arrtotals() As Currency
   Dim blnDoTotal As Boolean
   Dim l As Long

   If IsUnbound Then Exit Function

   With m_RsFiltered
      
      Dim intcounter As Integer

      'Build array of Field indexes for columns with totals
      For l = 0 To .Fields.count - 1
         If m_ntColumns(.Fields(l).Name).ShowTotal Then
            blnDoTotal = True
            ReDim Preserve arrTotalFields(intcounter)
            ReDim Preserve arrIndexes(intcounter)
            ReDim Preserve arrtotals(intcounter)
            arrtotals(intcounter) = 0
            arrTotalFields(intcounter) = l
            arrIndexes(intcounter) = m_ntColumns(.Fields(l).Name).index
            intcounter = intcounter + 1
         End If
      Next l

      If blnDoTotal Then

         If .RecordCount <> 0 Then

            .movefirst

            For lonRowCounter = 0 To .RecordCount - 1
               For l = 0 To UBound(arrTotalFields)
                  arrtotals(l) = arrtotals(l) + .Fields(arrTotalFields(l))
               Next l
              .MoveNext
            Next lonRowCounter

         Else
            For l = 0 To UBound(arrTotalFields)
               arrtotals(l) = 0
            Next l
         End If

         If lbltotal.UBound > 0 Then

            For l = 0 To UBound(arrIndexes)
               lbltotal(arrIndexes(l)).Text = m_ntColumns(arrIndexes(l)).FormatValue(arrtotals(l))
               If m_ntColumns(arrIndexes(l)).UseColoredNumbers Then
                  If lbltotal(arrIndexes(l)).Text >= 0 Then
                     lbltotal(arrIndexes(l)).ForeColor = m_clrPositive
                  Else
                     lbltotal(arrIndexes(l)).ForeColor = m_clrNegative
                  End If
               Else
                  lbltotal(arrIndexes(l)).ForeColor = Grid1.ForeColor
               End If
            Next l

         End If

      End If

   End With

End Function

' Takes a actual grid row (long), and returns the color for the row, or the cell depending on m_blnColorByRow
Private Function CheckRowColCriteria(ByVal lRow As Long, Optional ByVal lCol As Long = -1) As OLE_COLOR
   Dim pCol As ntColumn
   Dim lonFromCol As Long
   Dim lonToCol As Long
   Dim loncols As Long
   Dim lonColor As OLE_COLOR
   Dim vBkmrk As Variant
   Dim strFilter As String

On Error Resume Next

   If Not Ambient.UserMode Then Exit Function
   If m_RsFiltered Is Nothing Then Exit Function
   If m_RsFiltered.RecordCount = 0 Then Exit Function
   If m_ntColumns.count = 0 Then Exit Function

   ' Set the Filter to match RowID Passed in
   m_RsFiltered.AbsolutePosition = lRow + 1
   vBkmrk = m_RsFiltered.Bookmark

   ' If no column was passed in use all columns, else use pCol
   If lCol = -1 Then
     lonFromCol = 0
     lonToCol = (Grid1.Cols - 1) - Grid1.FixedCols
   Else
      lonFromCol = lCol
      lonToCol = lCol
      If m_ntColumns(loncols).BackColor <> -1 Then
         CheckRowColCriteria = m_ntColumns(loncols).BackColor
         Exit Function
      End If
   End If

   If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
      CheckRowColCriteria = m_colRowColors.RowItem("ID" & CStr(vBkmrk)).BackColor
      Exit Function
   End If

   'Set the default row color to be the same as Grid Backcolor
   lonColor = Grid1.BackColor

   'Loop through the affected columns
   For loncols = lonFromCol To lonToCol

       ' Get the column Object(0 based index starting after fixed column if there is one)
       Set pCol = m_ntColumns(loncols)

       If pCol.Visible = True Then

           ' If this column is set to check criteria
           If pCol.UseCriteria = True Then

               ' If the criteria matches, then get the color
               ' Only Return the first match, since a row can't contain multiple colors
               If pCol.RowCriteriaMatch(m_RsFiltered.Fields(pCol.Name).Value) Then
                  lonColor = pCol.RowCriteriaColor
                  Set pCol = Nothing
                  Exit For
               End If

           End If

       End If

   Next loncols

   CheckRowColCriteria = lonColor

End Function

Private Function HasRecords() As Boolean

On Error Resume Next

  HasRecords = False
  If m_RsFiltered Is Nothing Then Exit Function
  If m_RsFiltered.State = 0 Then Exit Function
  HasRecords = (m_RsFiltered.RecordCount > 0)

End Function

'**
'Method to clear all data, and reset the grid back to default state.
'If bClearColumns is set to true, all column structures will be cleared also.
'@param        bClearColumns Boolean. Optional. Default is False.
Public Sub ClearGrid(Optional ByVal bClearColumns As Boolean = False)
Attribute ClearGrid.VB_HelpID = 1080
   Dim i As Integer

On Error Resume Next
   
   If m_blnShown Then If bEditing Then Grid1_GotFocus
   
      
   With Grid1
      If bClearColumns Then
         If lbltotal.UBound > 0 Then
            For i = 1 To lbltotal.UBound
               Unload lbltotal(i)
            Next i
         End If
         lbltotal(0).Text = ""
         .Cols = 1 + .FixedCols
         If Ambient.UserMode Then
            hScroll.Visible = False
            m_bln_NeedHorzScroll = False
         End If
         .Rows = 1
         .FixedRows = 0
         .Rows = 0
      End If

      If m_blnHeaderRow Then
         .Rows = 1
         .Rows = 2
         .FixedRows = 1
         If m_sngRowHeightFixed Then .RowHeight(0) = m_sngRowHeightFixed
         .RowHeight(1) = m_sngRowHeight
      Else
         .Rows = 1
         .FixedRows = 0
         .Rows = 0
         .Rows = 1
      End If

      If Ambient.UserMode Then
         vScroll.Visible = False
         m_bln_NeedVertScroll = False
      End If
      
      .FocusRect = m_intFocusRect

   End With

End Sub

' Returns a perfect clone, without any records that were filtered in the Original
' Caveat: Can not deallocate memory used to copy RS to Stream until RS is closed
Private Function Clone(ByVal oRs As Object, _
   Optional ByVal LockType As Integer = 4) As Object

   Dim oStream As Object
   Dim oRsClone As Object

   If LockType < 1 Or LockType > 4 Then LockType = 4

   'save the recordset to the stream object
   Set oStream = CreateObject("ADODB.Stream")
   oRs.Save oStream

   'and now open the stream object into a new recordset
   Set oRsClone = CreateObject("ADODB.Recordset")
   oRsClone.Open oStream, , , LockType

   'return the cloned recordset
   Set Clone = oRsClone

   'release the references
   Set oStream = Nothing
   Set oRsClone = Nothing

End Function

'**
'This sub will apply a backcolor and optionally a forecolor to an entire Column in the grid
'@param        ColIndex Long.
'@param        cBackColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
'@param        cForeColor OLE_COLOR. Optional. A long or hexidecimal value indicating the color.
Public Sub ColorGridCol(ByVal vColKey As Variant, Optional ByVal cBackColor As OLE_COLOR = -1, _
                        Optional ByVal cForeColor As OLE_COLOR = -1)
Attribute ColorGridCol.VB_HelpID = 1090

   If m_ntColumns.Exists(vColKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".ColorGridCol", MSG_INVCOL
         
On Error Resume Next

   If m_blnShown Then If bEditing Then Grid1_GotFocus
   
   If cBackColor = -1 Then cBackColor = m_ntColumns(vColKey).BackColor
   If cForeColor = -1 Then cForeColor = m_ntColumns(vColKey).ForeColor

   m_ntColumns(vColKey).BackColor = cBackColor
   m_ntColumns(vColKey).ForeColor = cForeColor

   If m_blnShown Then Call ColorGridColumn(m_ntColumns(vColKey).index + Grid1.FixedCols, cBackColor, cForeColor)

End Sub

Private Sub ColorGridColumn(ByVal ColIndex As Long, ByVal cBackColor As OLE_COLOR, ByVal cForeColor As OLE_COLOR)
   
   m_blnIgnoreSel = True

On Error Resume Next

   If cBackColor = -1 Then cBackColor = Grid1.BackColor
   If cBackColor = -1 Then cBackColor = Grid1.BackColor
   
   SaveRestorePrevGrid Grid1
   
   With Grid1
      .Redraw = False
      .Col = ColIndex                                          ' Set Grid Column to ColIndex
      .colSel = ColIndex                                       ' Set Grid SelCol to ColIndex
      .Row = .FixedRows
      .RowSel = .Rows - 1
      .FillStyle = flexFillRepeat                              ' Set Grid Fillstyle to Repeat
      .CellBackColor = cBackColor
      .CellForeColor = cForeColor
      .FillStyle = flexFillSingle                              ' Set Grid Fillstyle Back to non-repeat
   End With                                                    ' FillStyle is a non-exposed property
   
   SaveRestorePrevGrid Grid1, False
   Grid1.Redraw = m_blnRedraw
      
   m_blnIgnoreSel = False

   DoEvents

End Sub

'**
'This sub will apply a backcolor and optionally a forecolor to an entire row in the grid
'@param        RowIndex Long.
'@param        cBackColor OLE_COLOR. Required. A long or hexidecimal value indicating the color.
'@param        cForeColor OLE_COLOR. Optional. A long or hexidecimal value indicating the color.
Public Function ColorGridRow(ByVal RowIndex As Long, _
                        Optional ByVal cBackColor As OLE_COLOR = -1, _
                        Optional ByVal cForeColor As OLE_COLOR = -1) As Variant
Attribute ColorGridRow.VB_HelpID = 1100

   Dim nRowInf As RowInfo
   Dim vBkmrk As Variant
   Dim strFilter As String
   Dim prevRedraw As Boolean
   
   ColorGridRow = -1
   
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".ColorGridRow", "Invalid recordset"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".ColorGridRow", "Display must be called prior to calling ColorGridRow"
   If RowIndex < 0 Or RowIndex > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".ColorGridRow", MSG_INVROW
  
On Error Resume Next
  
   If bEditing Then Grid1_GotFocus
   
   m_RsFiltered.AbsolutePosition = RowIndex + 1
   vBkmrk = m_RsFiltered.Bookmark

   'If the colors passed in were both -1, then just remove Row Coloring
   If cBackColor = -1 And cForeColor = -1 Then
      If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
         If m_colRowColors.RowItem("ID" & CStr(vBkmrk)).ItemData = -1 Then
            m_colRowColors.Remove "ID" & CStr(vBkmrk)
         Else
            m_colRowColors.RowItem("ID" & CStr(vBkmrk)).BackColor = cBackColor
            m_colRowColors.RowItem("ID" & CStr(vBkmrk)).ForeColor = cForeColor
         End If
      End If
   Else
      If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
         Set nRowInf = m_colRowColors.RowItem("ID" & CStr(vBkmrk))
      Else
         Set nRowInf = m_colRowColors.NewRowInfo
      End If
      nRowInf.BackColor = cBackColor
      nRowInf.ForeColor = cForeColor
      nRowInf.Bookmark = vBkmrk
      'nRowInf.CurrRow = m_rsFiltered.AbsolutePosition - 1
      If Not m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
         m_colRowColors.Insert nRowInf, "ID" & CStr(vBkmrk)
      End If
      Set nRowInf = Nothing
   End If

   If Not IsVisible(RowIndex) Then Exit Function

   prevRedraw = Grid1.Redraw

   Grid1.Redraw = False

   m_blnIgnoreSel = True

   If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then
      If cForeColor = -1 Then cForeColor = Grid1.ForeColor
      If cBackColor = -1 Then cBackColor = Grid1.BackColor
   Else
      cForeColor = Grid1.ForeColor
      cBackColor = Grid1.BackColor
   End If

   ' Adjust index for fixed row, if it exists(.FixedRows = 0 if not)
   RowIndex = ((RowIndex - GetScrollValue()) + Grid1.FixedRows)

   SaveRestorePrevGrid Grid1

   With Grid1
      .Row = RowIndex
      .RowSel = RowIndex
      .FillStyle = flexFillRepeat
      .Col = .FixedCols
      .colSel = .Cols - 1
      .CellForeColor = cForeColor
      .CellBackColor = cBackColor
      .FillStyle = flexFillSingle
   End With

   SaveRestorePrevGrid Grid1, False

   'Call FillRow(RowIndex)

   If m_blnRedraw Then Call CalcPaintedArea

   m_blnIgnoreSel = False

   Grid1.Redraw = m_blnRedraw

   ColorGridRow = vBkmrk

End Function
 
'**
'Copies entire recordset to the Clipboard.
'@param        bFieldNames Boolean. Optional. If not specified, does not include field names as first row.
'@param        pDelimiter String. Optional. If not specified, default is (,).
'@param        bFilteredRecords Boolean. Optional. If not specified, returns all records, otherwise filtered subset only.
Public Sub CopyAll(Optional ByVal bFieldNames As Boolean = False, _
                   Optional ByVal pDelimiter As String = vbTab, _
                   Optional ByVal bFiltered As Boolean = False)
Attribute CopyAll.VB_HelpID = 1110

   Dim txt As String
   Dim i As Long
   Dim RS As Object

On Error Resume Next
   
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub
   
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass

   ' If they wanted fieldnames at the top, concatenate the field name text
   If bFieldNames = True Then
      For i = 0 To m_RSMaster.Fields.count - 1
         txt = txt & m_RSMaster.Fields(i).Name & pDelimiter
      Next i
      ' Remove delimiter at end, add Carriage return
      txt = Left(txt, Len(txt) - 1) & vbCrLf
   End If

   ' If Filtered records only
   If bFiltered = True Then
      Set RS = m_RsFiltered.Clone
      RS.Filter = m_RsFiltered.Filter
   Else
      Set RS = m_RSMaster.Clone
      RS.Filter = m_RSMaster.Filter
   End If

   RS.movefirst

   ' Call the getstring method with no row count(gets all rows as default)
   txt = txt & RS.GetString(, , pDelimiter)

   Set RS = Nothing

   ' Add text to clipboard
   If Len(txt) > 0 Then
      Clipboard.Clear
      Clipboard.SetText txt, vbCFText
   End If

   Set RS = Nothing

   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault

End Sub

'**
'Copies Grid selection to the Clipboard.
'@param        bFieldNames Boolean. Optional. If not specified, does not include field names as first row.
'@param        pDelimiter String. Optional. If not specified, default is (,).
Public Sub CopySelectedCells(Optional ByVal pFieldNames As Boolean = False, _
                             Optional ByVal pDelimiter As String = vbTab)
Attribute CopySelectedCells.VB_HelpID = 1120

   Dim lonFromRow As Long
   Dim lonToRow As Long
   Dim lonFromCol As Long
   Dim lonToCol As Long
   Dim lonColCounter As Long
   Dim lonRowCounter As Long
   Dim intFields As Integer
   Dim txt As String
   Dim txtRow As String
   Dim arrFields() As String
   Dim blnAllFields As Boolean

On Error Resume Next

   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub

   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass

   ' Dimension the string field array, to hold field names for looping
   ' Much faster than calling the field.name property every time
   ' Use ABS since Col might be less than .ColSel if the user highlighted backwards
   ReDim arrFields(Abs(m_CurrCol - m_CurrColSel))

   If m_CurrRow > m_CurrRowSel Then
      lonFromRow = m_CurrRowSel
      lonToRow = m_CurrRow
   Else
      lonFromRow = m_CurrRow
      lonToRow = m_CurrRowSel
   End If

   If UBound(arrFields) = m_RsFiltered.Fields.count - 1 Then
      blnAllFields = True
   Else
      If m_CurrCol > m_CurrColSel Then
         lonFromCol = m_CurrColSel
         lonToCol = m_CurrCol
      Else
         lonFromCol = m_CurrCol
         lonToCol = m_CurrColSel
      End If

      ' Fill the array, and concatenate field name string
      For lonColCounter = lonFromCol To lonToCol
         ' Only use columns that are visible
         If m_ntColumns(lonColCounter).Visible = True Then
            arrFields(intFields) = m_ntColumns(lonColCounter).Name
            If pFieldNames = True Then txtRow = txtRow & arrFields(intFields) & pDelimiter
            intFields = intFields + 1
         End If
      Next lonColCounter

      ' Knock the array back down to actual field count
      ' Redim might have included hidden columns in the count
      ReDim Preserve arrFields(intFields - 1)

   End If

   ' If the header row was created, remove delimiter from end, add Carriage return
   If Len(txtRow) > 0 Then txt = Left(txtRow, Len(txtRow) - 1) & vbCrLf

   m_RsFiltered.AbsolutePosition = lonFromRow + 1

   ' Loop through and concatenate all the fields in each row, with delimiter
   For lonRowCounter = lonFromRow To lonToRow
      If blnAllFields Then
         ' Call the getstring method with no row count(gets all rows as default)
         txt = txt & m_RsFiltered.GetString(, 1, pDelimiter, vbCrLf)
      Else
         txtRow = vbNullString
         For lonColCounter = 0 To UBound(arrFields)
            txtRow = txtRow & m_RsFiltered.Fields(arrFields(lonColCounter)).Value & "" & pDelimiter
         Next lonColCounter
         ' Remove delimiter at end, add Carriage Return
         If Len(txtRow) > 0 Then txt = txt & Left(txtRow, Len(txtRow) - 1) & vbCrLf
         m_RsFiltered.MoveNext
      End If
      If m_RsFiltered.EOF Then Exit For
   Next lonRowCounter

   ' Add text to the clipboard
   If Len(txt) > 0 Then
      Clipboard.Clear
      Clipboard.SetText txt
   End If

   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault

End Sub

'$END-CODE$
'The key will be returned in the CustomMenuClick() event
'@param        sCaption String. Required. The caption shown on the menu.
'@param        sKey String. Required. The key used to reference the menu programatically.
'@param        bDefault Boolean. Optional. Make menu the default. Default value is false.
Public Function CustomMenuAddItem(ByVal sCaption As String, ByVal sKey As String, Optional ByVal bDefault As Boolean = False) As Long
Attribute CustomMenuAddItem.VB_HelpID = 1130

   Dim iIndex As Long

   CustomMenuAddItem = -1

   If mnuCustom.UBound = 0 And mnuCustom(0).Caption = "" Then
      iIndex = 0
   Else
      iIndex = mnuCustom.UBound + 1
      Load mnuCustom(iIndex)
   End If

   mnuCustom(iIndex).Visible = True
   mnuCustom(iIndex).Tag = sKey
   mnuCustom(iIndex).Caption = sCaption

   CustomMenuAddItem = iIndex

   mnuCustBar.Visible = True

End Function

'$END-CODE$
'@param        vKey Variant. Required. The index or key of the Custom Menu item.
'@param        bEnabled Boolean. Required. True enables the Custom Menu, false disables it.
Public Property Let CustomMenuItemEnabled(ByVal vKey As Variant, ByVal bEnabled As Boolean)
Attribute CustomMenuItemEnabled.VB_HelpID = 2300
Attribute CustomMenuItemEnabled.VB_MemberFlags = "400"

   Dim i As Integer

   If IsNumeric(vKey) Then
      If CustomMenuValidIndex(CInt(vKey)) Then mnuCustom(CInt(vKey)).Enabled = bEnabled
   Else
      For i = 0 To mnuCustom.UBound
         If mnuCustom(i).Tag = CStr(vKey) Then
            mnuCustom(i).Enabled = bEnabled
            Exit For
         End If
      Next i
   End If

End Property


'$END-CODE$
'@param        vKey Variant. Required. The index or key of the Custom Menu item.
'@param        bVisible Boolean.
Public Property Let CustomMenuItemVisible(ByVal vKey As Variant, ByVal bVisible As Boolean)
Attribute CustomMenuItemVisible.VB_HelpID = 2310
Attribute CustomMenuItemVisible.VB_MemberFlags = "400"

   Dim i As Integer

   If IsNumeric(vKey) Then
      If CustomMenuValidIndex(CInt(vKey)) Then mnuCustom(CInt(vKey)).Visible = bVisible
   Else
      For i = 0 To mnuCustom.UBound
         If mnuCustom(i).Tag = CStr(vKey) Then
            mnuCustom(i).Visible = bVisible
            Exit For
         End If
      Next i
   End If

End Property

'$END-CODE$
'@param        vKey Variant. Required. The index or key of the Custom Menu item to remove.
Public Sub CustomMenuRemoveItem(ByVal vKey As Variant)
Attribute CustomMenuRemoveItem.VB_HelpID = 1140
   Dim i As Long
   Dim iIndex As Long
   Dim bItems As Boolean

   If IsNumeric(vKey) Then
      iIndex = CLng(vKey)
   Else
      For i = 0 To mnuCustom.UBound
         If mnuCustom(i).Tag = CStr(vKey) Then
            iIndex = i
            Exit For
         End If
      Next i
   End If

   If CustomMenuValidIndex(iIndex) Then
      If mnuCustom.UBound <= 1 Then
         If (iIndex = 0) Then
            mnuCustom(0).Caption = ""
            mnuCustom(0).Tag = ""
            mnuCustom(0).Visible = False
         End If
      Else
         ' remove the item:
         For i = iIndex + 1 To mnuCustom.UBound
            mnuCustom(iIndex - 1).Caption = mnuCustom(iIndex).Caption
            mnuCustom(iIndex - 1).Tag = mnuCustom(iIndex).Tag
         Next i
         Unload mnuCustom(mnuCustom.UBound)
        End If
    End If

    bItems = False

    For i = 1 To mnuCustom.UBound
      If mnuCustom(i).Visible = True Then
         bItems = True
      End If
   Next i

   If bItems = False Then bItems = (Len(mnuCustom(0).Caption) > 0)

   mnuCustBar.Visible = bItems

End Sub

Private Function CustomMenuValidIndex(ByVal lIndex As Long) As Boolean
    CustomMenuValidIndex = (lIndex >= mnuCustom.LBound And lIndex <= mnuCustom.UBound)
End Function

Private Sub DrawNewRow(ByVal bTop As Boolean)
   Dim i As Integer

On Error Resume Next

   If bTop Then
      Grid1.RemoveItem Grid1.Rows - 1
      Grid1.AddItem "", Grid1.FixedRows
      Call FillRow(0)
      Grid1.RowHeight(Grid1.FixedRows) = m_sngRowHeight
   Else
      Grid1.RemoveItem Grid1.FixedRows
      Grid1.Rows = Grid1.Rows + 1
      Call FillRow((Grid1.Rows - 1) - Grid1.FixedRows)
      Grid1.RowHeight(Grid1.Rows - 1) = m_sngRowHeight
   End If

End Sub

'**
'Method will delete an entire row from the Grid and the underlying recordset.
'$EOL$
'$EOL$
'If Row is not specified, this will delete the current row.
'@param        lonRow Long. Optional. Default value is the currently selected row in the grid.
Public Sub DeleteRow(Optional ByVal lonRow As Long = -1)
Attribute DeleteRow.VB_HelpID = 1150
   Dim i As Long
   Dim lonRowData As Long
   Dim vBkmrk As Variant
   Dim vFilter As Variant
   Dim lonID As Long
   Dim arrPrevValues() As Variant
     
   ' Check the optional parameter, use current row if not passed in
   If lonRow = -1 Then lonRow = m_CurrRow
   
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".DeleteRow", "Invalid recordset"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".DeleteRow", "Display must be called prior to deleting a row"
   If Not HasRecords Or lonRow < 0 Or lonRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".DeleteRow", MSG_INVROW
  
On Error Resume Next
  
   If bEditing Then Grid1_GotFocus
   
   m_RsFiltered.AbsolutePosition = lonRow + 1
   vBkmrk = m_RsFiltered.Bookmark
   
   ReDim arrPrevValues(0 To m_RsFiltered.Fields.count - 1)

   For i = 0 To m_RsFiltered.Fields.count - 1
      If m_colColpics.Exists(m_RsFiltered.Fields(i).Name, vBkmrk) Then
         m_colColpics.Remove m_RsFiltered.Fields(i).Name, vBkmrk
      End If
      If m_colColpics.Exists("RS", vBkmrk) Then
         m_colColpics.Remove "RS", vBkmrk
      End If
      arrPrevValues(i) = m_RsFiltered.Fields(i)
   Next i
   
   m_RsFiltered.Delete
   
   m_RSMaster.Bookmark = vBkmrk
   m_RSMaster.Delete

   If m_colRowColors.RowItemExists("ID" & CStr(vBkmrk)) Then m_colRowColors.Remove "ID" & CStr(vBkmrk)

   Grid1.Redraw = False

   Call SetGridRows

   If IsVisible(lonRow) Then
      ' If there are no rows left, reset the grid
      If m_RsFiltered.RecordCount = 0 Then
         Call ClearGrid
         m_CurrRow = 0
         m_CurrRowSel = 0
      Else
         Call FillTextmatrix(GetScrollValue())
      End If
   End If

   If Not m_RsFiltered.RecordCount = 0 Then
      If m_CurrRow > m_RsFiltered.RecordCount - 1 Then m_CurrRow = m_CurrRow - 1
      If m_CurrRowSel > m_RsFiltered.RecordCount - 1 Then m_CurrRowSel = m_CurrRowSel - 1
   End If

   Call CalcPaintedArea

   Dim nCol As ntColumn

   For i = 0 To m_RsFiltered.Fields.count - 1
      Set nCol = m_ntColumns(m_RsFiltered.Fields(i).Name)
      If nCol.ShowTotal Then
         lbltotal(nCol.index).Text = _
            nCol.FormatValue(CDbl(lbltotal(nCol.index).Text) - CDbl(arrPrevValues(nCol.index)))
      End If
   Next i

   Set nCol = Nothing
   
   m_LastRow = m_CurrRow
   m_LastCol = m_CurrCol
   
   Grid1.Redraw = m_blnRedraw
   
   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus

End Sub

Private Sub DestroyRefs()

On Error Resume Next
   Set m_RSMaster = Nothing
   Set m_RsFiltered = Nothing
   Set m_ntColumns = Nothing
   Set m_colColpics = Nothing
   Set m_colRowColors = Nothing
   Set m_colFilters = Nothing
   Set m_hScroll = Nothing
   Set m_vScroll = Nothing
   
      
   DetachMessage Me, Grid1.hwnd, WM_LBUTTONUP
   DetachMessage Me, Grid1.hwnd, WM_ERASEBKGND
   DetachMessage Me, Grid1.hwnd, WM_LBUTTONDOWN
   DetachMessage Me, Grid1.hwnd, WM_PAINT
         
   DetachMessage Me, UserControl.hwnd, WM_TIMER

   m_blnGridSubclassed = False

End Sub

Public Sub Display()
Attribute Display.VB_HelpID = 1160

   If IsUnbound And m_blnGridMode Then Err.Raise ERR_NORS, Ambient.DisplayName & ".Display", "Cannot Display grid while unbound"
   
On Error Resume Next
   
   If m_blnShown Then If bEditing Then Grid1_GotFocus
   
   'Initalize Previous Scroll value so they are different first time
   m_PrevVertValue = -1

   Grid1.Redraw = False

   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass

   'If we are operating in bound mode
   If Not m_blnGridMode Then
      If m_ntColumns.count = 0 Then
         Err.Raise vbObjectError + 2009, Ambient.DisplayName & ".Display", "When in Unbound mode, you must add columns to the grid prior to display. Use the Cols property or the Columns.New and Columns.Insert to create the columns"
         Exit Sub
      End If
      m_blnLoading = True
      m_ntColumns.SetIndexOrder
      Call BuildRecordsetFromColumns
      Call SubclassGrid
      Call GetUnFilteredRecords(False)
      m_blnLoading = False
   End If
      
   If Not m_blnGridMode Then m_ntColumns.SetIndexOrder
   
   ' Initialize the grid
   Call BuildGridFromColumns(m_ntColumns)

   DoEvents

   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault

   m_blnShown = True

   hScroll.Value = 0
   vScroll.Value = 0

   If m_blnAutoSizeColumns Then Call AutosizeGridColumns

   Grid1.Col = m_CurrCol
   Grid1.Row = m_CurrRow

   SetGridRows

   Dim nCol As ntColumn

   For Each nCol In m_ntColumns
      If nCol.ShowTotal Then
         Call CalcTotals
         Exit For
      End If
   Next nCol

   FillTextmatrix 0

   DoEvents

   m_lonEditRow = 0
   m_lonEditRowSel = 0
   m_lonEditCol = 0
   m_CurrRow = 0
   m_CurrRowSel = 0
   m_CurrCol = 0
   m_CurrColSel = 0

   CalcPaintedArea
   
   m_LastRow = m_CurrRow
   m_LastCol = m_CurrCol
   
   Grid1.Redraw = m_blnRedraw

   DoEvents

End Sub

Private Sub BuildRecordsetFromColumns()

   Dim pCol As ntColumn
   Dim i As Integer

   Set m_RSMaster = Nothing
   Set m_RSMaster = CreateObject("ADODB.Recordset")

   With m_RSMaster
      .CursorType = 3
      .CursorLocation = 3
      .LockType = 4

      For i = 0 To m_ntColumns.count - 1
         .Fields.Append m_ntColumns(i).Name, 202, 255, 32
      Next i

     .Open
   End With

End Sub


' Row is 0 based index, with 0 being first non-fixed rows
'Assumes Cscroll.Value and Grid.Rows are already set
Private Sub FillRow(ByVal pRow As Long)
   Dim l As Long
   Dim i As Integer
   Dim pCol As ntColumn
   Dim pCols() As ntColumn
   Dim blnColored As Boolean
   Dim lonForeColor As Long
   Dim lonBackColor As Long
   Dim vBkmrk As Variant
   Dim rowId As String
   Dim vValue As Variant
   Dim blnChecked As Boolean

   If Not Ambient.UserMode Then Exit Sub

On Error Resume Next

   'Go to the correct record for the row
   m_RsFiltered.AbsolutePosition = GetScrollValue() + pRow + 1
   vBkmrk = m_RsFiltered.Bookmark

   ReDim pCols(m_ntColumns.count - 1)

   i = 0

   ' Build the array of ntColumns
   For l = 0 To m_ntColumns.count - 1
      Set pCols(i) = m_ntColumns(l)
      i = i + 1
   Next l

   With Grid1

      blnColored = False

      rowId = "ID" & CStr(vBkmrk)

      If m_colRowColors.RowItemExists(rowId) Then
         lonForeColor = m_colRowColors.RowItem(rowId).ForeColor
         lonBackColor = m_colRowColors.RowItem(rowId).BackColor
         If lonForeColor = -1 Then lonForeColor = Grid1.ForeColor
         If lonBackColor = -1 Then lonBackColor = Grid1.BackColor
         blnColored = lonBackColor <> Grid1.BackColor
      Else
         lonForeColor = Grid1.ForeColor
         lonBackColor = Grid1.BackColor
      End If

      .Row = pRow + .FixedRows
      .RowSel = pRow + .FixedRows
      .FillStyle = flexFillRepeat
      .Col = .FixedCols
      .colSel = .Cols - 1
      .CellForeColor = lonForeColor
      .CellBackColor = lonBackColor
      .FillStyle = flexFillSingle

      If .FixedCols > 0 Then
         If m_colColpics.Exists("RS", vBkmrk) Then
            If m_colColpics("RS", vBkmrk).HasPic Then
               .Col = 0
               .CellPictureAlignment = m_colColpics("RS", vBkmrk).PictureAlignment
               Set Grid1.CellPicture = m_colColpics("RS", vBkmrk).CellPicture
            End If
         End If
      End If

      For i = 0 To UBound(pCols)
         
         If pCols(i).Visible Then
         
            Set pCol = pCols(i)
   
            vValue = m_RsFiltered.Fields(pCol.Name).Value
   
            If pCol.UseCriteria = True Then
               If m_blnColorByRow = True Then
                  If Not blnColored Then
                     If pCol.RowCriteriaMatch(vValue) = True Then
                        .FillStyle = flexFillRepeat
                        .Col = .FixedCols
                        .colSel = .Cols - 1
                        .CellBackColor = pCol.RowCriteriaColor
                        .FillStyle = flexFillSingle
                        blnColored = True
                     End If
                  End If
               Else
                  If Not blnColored Then
                     If pCol.RowCriteriaMatch(vValue) = True Then
                        .Col = i + .FixedCols
                        .CellBackColor = pCol.RowCriteriaColor
                     End If
                  End If
               End If
            End If
   
            If IsNull(vValue) Or vValue = "" Then
               blnChecked = False
            Else
               blnChecked = CBool(vValue)
            End If
                 
            If pCol.UseColoredNumbers Then
                If IsNumeric(vValue) Then
                  .Col = i + .FixedCols
                  If vValue >= 0 Then
                     .CellForeColor = m_clrPositive
                  Else
                     .CellForeColor = m_clrNegative
                  End If
               End If
            End If
   
            Select Case pCol.ColFormat
               Case nfgBooleanCheckBox
                  If IsNull(vValue) Or vValue = "" Then
                     blnChecked = False
                  Else
                     blnChecked = CBool(vValue)
                  End If
                  .Col = i + .FixedCols
                  .CellPictureAlignment = 4
                  Call SetPicture(blnChecked, pCol.Enabled)
               Case nfgPicture
                  If m_colColpics.Exists(pCol.Name, vBkmrk) Then
                     If m_colColpics(pCol.Name, vBkmrk).HasPic Then
                        .Col = i + .FixedCols
                        .CellPictureAlignment = m_colColpics(pCol.Name, vBkmrk).PictureAlignment
                        Set .CellPicture = m_colColpics(pCol.Name, vBkmrk).CellPicture
                     End If
                  End If
               Case nfgPictureText, nfgPictureCustFormatText
                  If m_colColpics.Exists(pCol.Name, vBkmrk) Then
                     If m_colColpics(pCol.Name, vBkmrk).HasPic Then
                        .Col = i + .FixedCols
                        .CellPictureAlignment = m_colColpics(pCol.Name, vBkmrk).PictureAlignment
                        Set .CellPicture = m_colColpics(pCol.Name, vBkmrk).CellPicture
                     End If
                  End If
                  .TextMatrix(.Row, i + .FixedCols) = pCol.FormatValue(vValue)
               Case Else
                  .TextMatrix(.Row, i + .FixedCols) = pCol.FormatValue(vValue)
            End Select
         
         End If
         
      Next i

      For i = 0 To UBound(pCols)
         If pCols(i).Visible Then
            If pCols(i).ForeColor <> -1 Or pCols(i).BackColor <> -1 Then
               .Col = i + .FixedCols
               .CellBackColor = pCols(i).BackColor
               .CellForeColor = pCols(i).ForeColor
            End If
         End If
      Next i

   End With

   Erase pCols

   Set pCol = Nothing

   m_blnLoading = False

End Sub

Private Sub FillTextmatrix(ByVal pFirstRecord As Long)
   Dim l As Long
   Dim i As Integer
   Dim pCol As ntColumn
   Dim pCols() As ntColumn
   Dim blnColored As Boolean
   Dim lRows As Long
   Dim lonBackColor As Long
   Dim lonForeColor As Long
   Dim rowId As String
   Dim vBkmrk As Variant
   Dim prevRedraw As Boolean
   Dim vValue As Variant
   Dim blnChecked As Boolean
   Dim sName As String

On Error Resume Next

   If m_RsFiltered.RecordCount = 0 Then
      With Grid1
         .Redraw = False
         .Rows = 1 + Grid1.FixedRows
         .Row = .FixedRows
         .Col = 0
         .RowSel = .Row
         .colSel = .Cols - 1
         .FillStyle = flexFillRepeat
         Set .CellPicture = Nothing
         .Text = ""
         .FillStyle = flexFillSingle
         .Redraw = True
      End With
      Exit Sub
   End If

On Error GoTo Fill_Err
  
   'Go to the correct record for the row
   m_RsFiltered.AbsolutePosition = pFirstRecord + 1

   m_blnLoading = True
   m_blnIgnoreSel = True
      
   ReDim pCols(m_ntColumns.count - 1)

   i = 0

   ' Build the array of ntColumns
   For l = 0 To m_ntColumns.count - 1
      Set pCols(l) = m_ntColumns(l)
      i = i + 1
   Next l
   
   With Grid1

      prevRedraw = .Redraw
      .Redraw = False

      If .FixedCols > 0 Then
         .Row = .FixedRows
         .Col = 0
         .RowSel = .Rows - 1
         .colSel = 0
         .FillStyle = flexFillRepeat
         Set .CellPicture = Nothing
         .FillStyle = flexFillSingle
      End If

      For l = .FixedRows To .Rows - 1

         blnColored = False

         vBkmrk = m_RsFiltered.Bookmark
         rowId = "ID" & CStr(vBkmrk)

         If m_colRowColors.RowItemExists(rowId) Then
            lonForeColor = m_colRowColors.RowItem(rowId).ForeColor
            lonBackColor = m_colRowColors.RowItem(rowId).BackColor
            If lonForeColor = -1 Then lonForeColor = Grid1.ForeColor
            If lonBackColor = -1 Then lonBackColor = Grid1.BackColor
            blnColored = lonBackColor <> Grid1.BackColor
         Else
            lonForeColor = Grid1.ForeColor
            lonBackColor = Grid1.BackColor
         End If

         .Row = l
         .RowSel = l
         .FillStyle = flexFillRepeat
         .Col = .FixedCols
         .colSel = (.Cols - 1)
         .CellForeColor = lonForeColor
         .CellBackColor = lonBackColor
         .FillStyle = flexFillSingle

         If .FixedCols > 0 Then
            If m_colColpics.Exists("RS", vBkmrk) Then
               If m_colColpics("RS", vBkmrk).HasPic Then
                  .Col = 0
                  .CellPictureAlignment = m_colColpics("RS", vBkmrk).PictureAlignment
                  Set Grid1.CellPicture = m_colColpics("RS", vBkmrk).CellPicture
               End If
            End If
         End If

         For i = 0 To UBound(pCols)
            
            If pCols(i).Visible Then
            
               Set pCol = pCols(i)
                           
               sName = pCol.Name
               vValue = m_RsFiltered.Fields(sName).Value
   
               If pCol.UseCriteria = True Then
                  If m_blnColorByRow = True Then
                     If Not blnColored Then
                        If m_ntColumns(sName).RowCriteriaMatch(vValue) = True Then
                           .FillStyle = flexFillRepeat
                           .Col = .FixedCols
                           .colSel = (.Cols - 1)
                           .CellBackColor = pCol.RowCriteriaColor
                           .FillStyle = flexFillSingle
                           blnColored = True
                        End If
                     End If
                  Else
                     If Not blnColored Then
                        If pCol.RowCriteriaMatch(vValue) = True Then
                           .Col = i + .FixedCols
                           .CellBackColor = pCol.RowCriteriaColor
                        End If
                     End If
                  End If
               End If
               
               If pCol.UseColoredNumbers Then
                  If IsNumeric(vValue) Then
                     .Col = i + .FixedCols
                     If vValue >= 0 Then
                        .CellForeColor = m_clrPositive
                     Else
                        .CellForeColor = m_clrNegative
                     End If
                  End If
               End If
   
               Select Case pCol.ColFormat
                  Case nfgBooleanCheckBox
                     If IsNull(vValue) Or vValue = "" Then
                        blnChecked = False
                     Else
                        blnChecked = CBool(vValue)
                     End If
                     .Col = i + .FixedCols
                     .CellPictureAlignment = 4
                     Call SetPicture(blnChecked, pCol.Enabled)
                  Case nfgPicture
                     If m_colColpics.Exists(sName, vBkmrk) Then
                        If m_colColpics(sName, vBkmrk).HasPic Then
                           .Col = i + .FixedCols
                           .CellPictureAlignment = m_colColpics(sName, vBkmrk).PictureAlignment
                           Set .CellPicture = m_colColpics(sName, vBkmrk).CellPicture
                        End If
                     End If
                  Case nfgPictureText, nfgPictureCustFormatText
                     If m_colColpics.Exists(sName, vBkmrk) Then
                        If m_colColpics(sName, vBkmrk).HasPic Then
                           .Col = i + .FixedCols
                           .CellPictureAlignment = m_colColpics(sName, vBkmrk).PictureAlignment
                           Set .CellPicture = m_colColpics(sName, vBkmrk).CellPicture
                        End If
                     End If
                     .TextMatrix(l, i + .FixedCols) = pCol.FormatValue(vValue)
                  Case Else
                     .TextMatrix(l, i + .FixedCols) = pCol.FormatValue(vValue)
               End Select
            
            End If
            
         Next i

         m_RsFiltered.MoveNext

      Next l

      For i = 0 To UBound(pCols)
         If pCols(i).Visible Then
            If pCols(i).ForeColor <> -1 Or pCols(i).BackColor <> -1 Then
               Call ColorGridColumn(i + .FixedCols, pCols(i).BackColor, pCols(i).ForeColor)
            End If
         End If
      Next i

      .Redraw = prevRedraw

   End With

   Erase pCols

   Set pCol = Nothing

   m_blnChanged = False
   m_blnLoading = False
   m_blnIgnoreSel = False

Exit Sub

Fill_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".FillTextmatrix; " & Err.Source, Err.Description
   Resume Next
End Sub

'**
'Method to find a particular value in the Grid.
'$EOL$
'Returns -1 if not found, Row number if found.
'@param        vValue Variant. Required. The value to find in the recordset.
'@param        vColKey Variant. Required. The index or key of the column to search.
'@param        bStartFirst Boolean. Optional. Start search at the first record. Default is false, start at the first record after the current one.
'@param        bHighlightRow Boolean. Optional. A value of True will Highlight the row in the Grid where the value was found, if it was found.
Public Function Find(ByVal vValue As Variant, ByVal vColKey As Variant, _
                      Optional ByVal StartBookmark As Variant = -1, _
                      Optional ByVal bForwards As Boolean = True, _
                      Optional ByVal bHighlightRow As Boolean = False) As Variant
Attribute Find.VB_HelpID = 1170

   Dim strFind As String
   Dim vStartRow As Variant
   Dim lonIndex As Long
   Dim blnFound As Boolean
   Dim lonColID As Long
   Dim vBkmrk As Variant
   Dim strFilter As String
   Dim iDirection As Integer
   
   Find = -1
   
   If IsUnbound Then Exit Function
   If Not m_blnShown Then Exit Function
   If Not HasRecords Then Exit Function
    
On Error GoTo Find_Err
  
   lonIndex = -1
      
   If bForwards Then
      iDirection = 1
   Else
      iDirection = -1
   End If

   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass

   ' If the value to find has an asterisk, use 'Like', otherwise use '='
   If InStr(1, vValue, "*", vbTextCompare) > 0 Then
      strFind = m_ntColumns(vColKey).Name & " Like '" & CStr(vValue) & "'"
   Else
      strFind = m_ntColumns(vColKey).Name & " = '" & CStr(vValue) & "'"
   End If

   ' Figure out which column to search
   lonColID = m_ntColumns(vColKey).index

   ' If we are starting at the beginning, move to first record
   If StartBookmark = -1 Then
      m_RsFiltered.movefirst
      StartBookmark = m_RsFiltered.Bookmark
   Else
      m_RsFiltered.Bookmark = StartBookmark
      m_RsFiltered.MoveNext
      If m_RsFiltered.EOF Then Exit Function
      StartBookmark = m_RsFiltered.Bookmark
   End If

   ' Call the ADO recordset find method
   m_RsFiltered.Find strFind, , iDirection, StartBookmark

   ' If the Recordset went to .EOF then it didn't find anything
   blnFound = Not (m_RsFiltered.EOF)

   If blnFound Then

      Find = m_RsFiltered.Bookmark

      lonIndex = m_RsFiltered.AbsolutePosition - 1

      If bHighlightRow Then
         Call HighlightRow(lonIndex)
      Else
         m_blnManualScroll = True
         If Not IsVisible(lonIndex) Then
            If m_bln_NeedVertScroll Then
               If lonIndex > vScroll.Max Then
                  vScroll.Value = vScroll.Max
               Else
                  vScroll.Value = lonIndex
               End If
            Else
               FillTextmatrix 0
            End If
         End If
         m_blnManualScroll = False
         m_CurrRow = lonIndex
         m_CurrRowSel = m_CurrRow
         m_CurrCol = m_ntColumns(vColKey).index
         m_CurrColSel = m_CurrCol
         m_LastRow = m_CurrRow
         m_LastCol = m_CurrCol
         Call CalcPaintedArea
         On Error Resume Next
         If m_blnHasFocus Then Grid1.SetFocus
      End If

   End If

   Grid1.Redraw = m_blnRedraw

   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault

Exit Function

Find_Err:
   m_blnLoading = False
   Grid1.Redraw = m_blnRedraw
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
   Find = -1
End Function

Private Sub FitColumnsToGrid()
  Dim i As Integer
  Dim lonColTotalWidth As Long
  Dim lonTempWidth As Long
  Dim lonLastColAdj As Long

On Error Resume Next

  If m_ntColumns.count = 0 Then Exit Sub

  With Grid1

    For i = .FixedCols To .Cols - 1

      If m_ntColumns(i - .FixedCols).Visible = True Then

         ' Add the column widths together for use later
         lonColTotalWidth = lonColTotalWidth + .ColWidth(i)

      End If

    Next i

    If m_blnRecordSelectors Then lonColTotalWidth = lonColTotalWidth + m_sngRecordSelectorWidth

    ' If the total width of all columns is less then the grid width
    If lonColTotalWidth < .Width Then

        ' Calc the space to add to each Column
        lonTempWidth = Fix(((.Width - lonColTotalWidth)) / ((.Cols - m_ntColumns.HiddenColumns) - .FixedCols))
        lonLastColAdj = (.Width - lonColTotalWidth) - (((.Cols - m_ntColumns.HiddenColumns) - .FixedCols) * lonTempWidth)

        Dim lonNewWidth As Long

        ' Add the necessary width to each column
        For i = .FixedCols To .Cols - 1

            If m_ntColumns(i - .FixedCols).Visible = True Then

               'Get the proposed new width
               If i = .Cols - 1 Then
                  lonNewWidth = .ColWidth(i) + lonTempWidth + lonLastColAdj
               Else
                 lonNewWidth = .ColWidth(i) + lonTempWidth
               End If

                ' Set the width of the column to match largest text
                ' Check property settings for Min and Max Col Widths
                If lonNewWidth < m_sngMinColWidth Then
                    .ColWidth(i) = m_sngMinColWidth
                ElseIf lonNewWidth > m_sngMaxColWidth Then
                    If m_sngMaxColWidth <> 0 Then
                      .ColWidth(i) = m_sngMaxColWidth
                    Else
                      .ColWidth(i) = lonNewWidth
                    End If
                Else
                    .ColWidth(i) = lonNewWidth
                End If

                m_ntColumns(i - .FixedCols).Width = .ColWidth(i)

            End If

        Next i

    End If

  End With

End Sub

Private Sub GetUnFilteredRecords(Optional ByVal bCalc As Boolean = True)
   Dim l As Long
   Dim prevfilter As Variant

 On Error GoTo GF_Err

   If m_RSMaster Is Nothing Then Exit Sub
   If m_RSMaster.State <> 1 Then Exit Sub

   prevfilter = m_RSMaster.Filter
   m_RSMaster.Filter = 0

   Set m_RsFiltered = Nothing
   Set m_RsFiltered = Clone(m_RSMaster)

   m_RsFiltered.Sort = m_RSMaster.Sort
   m_RsFiltered.Filter = prevfilter
   m_RSMaster.Filter = prevfilter

   For l = 0 To lbltotal.UBound
     lbltotal(l).Text = ""
   Next l

   Set m_colFilters = New ntColFilters

   If bCalc Then

      Dim nCol As ntColumn

      For Each nCol In m_ntColumns
         If nCol.ShowTotal Then
            Call CalcTotals
            Exit For
         End If
      Next

   End If

Exit Sub

GF_Err:
   Err.Raise Err.Number, "GetUnfilteredRecords; " & Err.Source, Err.Description
End Sub

'**
'Returns an ADO recordset with a filter applied to match the current records displayed in the Grid.
Public Function GetFilteredRecordset() As Object
Attribute GetFilteredRecordset.VB_HelpID = 1180
   Dim RS As Object
   Dim arrBkmrk() As Variant
   Dim lonRecord As Long

On Error Resume Next

   If m_RsFiltered Is Nothing Or m_RSMaster Is Nothing Then
      Set GetFilteredRecordset = Nothing
      Exit Function
   End If

   Set RS = m_RsFiltered.Clone
   RS.Filter = m_RsFiltered.Filter

   Set GetFilteredRecordset = RS

   Set RS = Nothing

End Function

'**
'Method to highlight a selected row, using the ForeColorSel and BackColorSel properties.
'$EOL$
'$EOL$
'If Row does not exist, setting will be ignored.
'@param        lonRow Long. Required.
'@rem Note: if Row is not visible, then the Grid will scroll to it.
Public Sub HighlightRow(ByVal lRow As Long)
Attribute HighlightRow.VB_HelpID = 1190

On Error Resume Next

   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".HighlightRow", "Invalid recordset"
   If Not m_blnShown Then Err.Raise ERR_NORS + 1, Ambient.DisplayName & ".HighlightRow", "Display must be called prior to deleting a row"
   If Not HasRecords Or lRow < 0 Or lRow > m_RsFiltered.RecordCount - 1 Then Err.Raise ERR_INVROW, Ambient.DisplayName & ".HighlightRow", MSG_INVROW
   
   If bEditing Then Grid1_GotFocus
   
   m_CurrRow = lRow
   m_CurrRowSel = lRow
   m_CurrCol = 0
   m_CurrColSel = (Grid1.Cols - 1) - Grid1.FixedCols
     
   Grid1.Redraw = False

   If Not IsVisible(lRow) Then
      If lRow > vScroll.Max Then
         vScroll.Value = vScroll.Max
      Else
         vScroll.Value = lRow
      End If
   Else
      Grid1.Row = lRow - GetScrollValue() + Grid1.FixedRows
      Grid1.RowSel = Grid1.Row
      Grid1.Col = Grid1.FixedCols
      Grid1.colSel = Grid1.Cols - 1
   End If

   Grid1.Redraw = m_blnRedraw
   
   m_LastRow = lRow
   m_LastCol = 0
    
   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus

   DoEvents

End Sub

Private Sub HookRs(ByVal RS As Object, ByVal bDisconnect As Boolean)

   Dim nCol As ntColumn
   Dim prevfilter As Variant
   Dim lCount As Long
   Dim l As Long

On Error GoTo Recordset_Err
  
   m_lonEditRow = 0
   m_lonEditRowSel = 0
   m_lonEditCol = 0
   m_CurrRow = 0
   m_CurrRowSel = 0
   m_CurrCol = 0
   m_CurrColSel = 0
   m_LastRow = m_CurrRow
   m_LastCol = m_CurrCol
   m_blnShown = False
   
   hScroll.Value = 0
   vScroll.Value = 0
      
   If Not (RS Is Nothing) Then

      m_blnGridMode = True

      m_blnLoading = True

      Set m_RSMaster = Nothing
      Set m_RsFiltered = Nothing
      Set m_colRowColors = New ntRowInfo
      Set m_ntColumns = New ntColumns
      Set m_colFilters = New ntColFilters
      Set m_colColpics = New ntColPics

      Call SubclassGrid

      If Not bDisconnect Then
         Set m_RSMaster = RS
         m_RSMaster.Filter = RS.Filter
      Else
         ' Alternate method to disconnect
         prevfilter = RS.Filter
         RS.Filter = 0
         'Disconnect the recordset - all records
         Set m_RSMaster = Clone(RS)
         RS.Filter = prevfilter
         m_RSMaster.Filter = prevfilter
      End If

      Call InitDefaultColumns

      Call GetUnFilteredRecords(False)
     
      m_blnLoading = False

   Else

      m_blnLoading = True

      hScroll.Value = 0
      vScroll.Value = 0

      Call ClearGrid(True)

      m_blnLoading = False

      Call DestroyRefs

   End If

Exit Sub

Recordset_Err:

   m_blnLoading = True

   If Not hScroll Is Nothing Then hScroll.Value = 0
   If Not vScroll Is Nothing Then vScroll.Value = 0

   Call ClearGrid(True)

   Call DestroyRefs

   m_blnLoading = False

   Err.Raise vbObjectError + Err.Number, UserControl.Name & ".Recordset", Err.Description

End Sub

Private Sub InitDefaultColumns()

   Dim i As Integer
   Dim loncols As Long
   Dim lonRows As Long
   Dim pCol As ntColumn
   Dim lonwidth As Long

On Error Resume Next

   ' If no recordset, can't continue
   If m_RSMaster Is Nothing Then Exit Sub
   
   lonwidth = m_sngMinColWidth
   
   If m_def_ColWidth > m_sngMinColWidth Then
      If m_def_ColWidth < m_sngMaxColWidth Or m_sngMaxColWidth = 0 Then
         lonwidth = m_def_ColWidth
      Else
         lonwidth = m_sngMaxColWidth
      End If
   End If
      
   For loncols = 0 To m_RSMaster.Fields.count - 1

      Set pCol = m_ntColumns.NewColumn

      pCol.Width = lonwidth

      ' Use field name for col name
      pCol.Name = m_RSMaster.Fields(loncols).Name
      ' Store the datatype for reference
      pCol.DataType = m_RSMaster.Fields(loncols).Type
      pCol.Visible = True
      ' If UseFieldNamesAsHeader, then write header Text
      If m_blnUseFieldNamesAsHeader Then
         'Default header text to field names, can be edited through Columns Collection
         pCol.HeaderText = pCol.Name
      End If
      
      ' Add Column to collection
      m_ntColumns.Insert pCol, pCol.Name

   Next loncols

End Sub

Private Function IsUnbound() As Boolean
   IsUnbound = (m_RSMaster Is Nothing)
End Function

Private Function IsVisible(ByVal pRowNum As Long) As Boolean

   If Not Ambient.UserMode Then Err.Raise 380

On Error Resume Next

   If pRowNum >= GetScrollValue() And _
      pRowNum <= (GetScrollValue() + ((Grid1.Rows - 1) - Grid1.FixedRows)) Then
         IsVisible = True
   End If

End Function

Public Property Get LeftCol() As Variant
Attribute LeftCol.VB_HelpID = 2320
Attribute LeftCol.VB_MemberFlags = "400"
   
   LeftCol = 0
   
   If Not Ambient.UserMode Then Err.Raise 387
   If IsUnbound Then Exit Property
   If m_bln_NeedHorzScroll = False Then Exit Property

   LeftCol = Grid1.LeftCol - Grid1.FixedCols

End Property

'**
'Returns a long specifying the currently selected column in the grid.
'Sets the currently selected column in the grid, either by index, or field name.
'@param        vKey Variant. Required. A valid index or key specifying the column in the Grid.
Public Property Let LeftCol(ByVal vKey As Variant)

   Dim i As Integer
   Dim iNewScrollVal As Integer
   Dim iNewLCol As Integer

   If Not Ambient.UserMode Then Err.Raise 387
   If m_bln_NeedHorzScroll = False Then Exit Property
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".HighlightRow", "Invalid recordset"
   If m_ntColumns.Exists(vKey) = False Then Err.Raise ERR_INVCOL, Ambient.DisplayName & ".LeftCol", MSG_INVCOL
   If m_ntColumns(vKey).Visible = False Then Err.Raise ERR_INVCOL + 2, Ambient.DisplayName & ".LeftCol", "Cannot set LeftCol property to hidden column"
   If Not m_blnShown Then Err.Raise ERR_NORS, Ambient.DisplayName & ".LeftCol", "Cannot set LeftCol prior to calling display"
    
 On Error GoTo Col_Err

   If bEditing Then Grid1_GotFocus
   
   iNewLCol = m_ntColumns(vKey).index

   If iNewLCol < m_arrVisCols(LBound(m_arrVisCols)) Then iNewLCol = m_arrVisCols(LBound(m_arrVisCols))
   If iNewLCol > m_arrVisCols(UBound(m_arrVisCols)) Then iNewLCol = m_arrVisCols(UBound(m_arrVisCols))

   iNewScrollVal = -1

   For i = 0 To UBound(m_arrVisCols)
      If m_arrVisCols(i) = iNewLCol Then
         iNewScrollVal = i
         Exit For
      End If
   Next i

   m_blnLoading = True

   hScroll.Value = iNewScrollVal
   Grid1.Height = UserControl.Height
   Grid1.ScrollBars = flexScrollBarHorizontal
   Grid1.LeftCol = iNewScrollVal + Grid1.FixedCols
   Grid1.ScrollBars = flexScrollBarNone
   Grid1.Height = UserControl.ScaleHeight

   If m_blnHeaderRow Then SetTotalPositions

   m_PrevHorzValue = iNewScrollVal

   m_blnLoading = False

   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus

Exit Property

Col_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".LeftCol: " & Err.Source, Err.Description
End Property

'Add the following procedure to initialize the list/combo box and pass the focus from the Hierarchical FlexGrid to the TextBox control:
Private Sub MSFlexGridCombo(Edt As Control, KeyAscii As Integer)
   Dim i As Integer
   Dim intIndex As Integer

On Error Resume Next

   m_blnInitEdit = True
   m_blnCancelEdit = False

   m_RsFiltered.AbsolutePosition = m_lonEditRow + 1

   m_PrevEditVal = CStr(m_RsFiltered.Fields(m_ntColumns(m_lonEditCol).Name).Value & "")

   ' Use the character that was typed.
   Select Case KeyAscii

   ' A space means edit the current text.
   Case 0 To 32

      'If you can't set the text
      If Edt.Style = 2 Then

          intIndex = -1

          For i = 0 To Edt.ListCount - 1
            If m_PrevEditVal = Edt.List(i) Then
              intIndex = i
              Exit For
            End If
          Next i

          m_blnLoading = True
          Edt.ListIndex = intIndex
          m_blnLoading = False

      Else

          Edt.Text = m_PrevEditVal
          Edt.SelStart = 0
          Edt.SelLength = Len(m_PrevEditVal)

      End If

  End Select

  ' Show Edt at the right place.
  Edt.Move Grid1.Left + Grid1.CellLeft, _
   Grid1.Top + Grid1.CellTop - 45, _
   Grid1.CellWidth + 30

  Edt.Visible = True
  Edt.ZOrder 0

  ' And make it work.
  Edt.SetFocus

  m_blnInitEdit = False

End Sub

 'Add the following procedure to initialize the text box and pass the focus from the Hierarchical FlexGrid to the TextBox control:
Private Sub MSFlexGridEdit(Edt As Control, KeyAscii As Integer)

  Dim intIndex As Integer

  m_blnInitEdit = True
  m_blnCancelEdit = False

On Error Resume Next

  m_RsFiltered.AbsolutePosition = m_lonEditRow + 1

  m_PrevEditVal = CStr(m_RsFiltered.Fields(m_ntColumns(m_lonEditCol).Name).Value & "")

  ' Use the character that was typed.
  Select Case KeyAscii

    ' A space means edit the current text.
    Case 0 To 32
      Edt.Text = m_PrevEditVal
      Edt.SelStart = 0
      Edt.SelLength = Len(m_PrevEditVal) + 1
    ' Anything else means replace the current text.
    Case Else
      Edt.Text = Chr(KeyAscii)
      Edt.SelStart = 1
  End Select

  ' Show Edt at the right place.
  Edt.Move Grid1.Left + Grid1.CellLeft, _
           Grid1.Top + Grid1.CellTop, _
           Grid1.CellWidth, Grid1.CellHeight

  Edt.Visible = True
  Edt.ZOrder 0

   ' And make it work.
  Edt.SetFocus

  m_blnInitEdit = False

End Sub

Private Sub RecalcGrid()
   Dim i As Integer
   Dim l As Long
   Dim lRows As Long
   Dim hHeight As Integer
   Dim lPrevScroll As Long
   
On Error Resume Next

   If Not m_blnShown Or IsUnbound Then Exit Sub
       
   lPrevScroll = GetScrollValue
   lRows = Grid1.Rows
   Grid1.Redraw = False
   SetGridRows
   If lPrevScroll > vScroll.Max Then
      vScroll.Value = vScroll.Max
   End If
   If Grid1.Rows > lRows Then
      FillTextmatrix GetScrollValue
   End If
   CalcPaintedArea
   Grid1.Redraw = m_blnRedraw
   Call SetTotalPositions

End Sub

'Will Reorder columns by index
Public Sub ReorderColumns(ByRef arrLonNewIndexes() As Long, Optional ByVal bRedraw As Boolean = True)
Attribute ReorderColumns.VB_HelpID = 1200
  Dim l As Long
  Dim lIndex As Long
  Dim blnMissing As Boolean

   If UBound(arrLonNewIndexes) + 1 <> m_ntColumns.count Then Err.Raise vbObjectError + 2056, Ambient.DisplayName & ".ReorderColumns", "Indexes must have the same number of elements as number of columns."
   If m_ntColumns.count = 0 Then Err.Raise ERR_INVCOL + 10, Ambient.DisplayName & ".ReorderColumns", "Cannot ReorderColumns in a grid with no columns"
   If IsUnbound Then Err.Raise ERR_NORS, Ambient.DisplayName & ".ReorderColumns", "Invalid recordset"
   If Not m_blnShown Then Err.Raise ERR_NORS, Ambient.DisplayName & ".ReorderColumns", "Cannot ReorderColumns prior to calling display"
          
On Error Resume Next

   If bEditing Then Grid1_GotFocus
   
   lIndex = 0

   For l = 0 To UBound(arrLonNewIndexes)
      blnMissing = True
      For lIndex = 0 To UBound(arrLonNewIndexes)
         If arrLonNewIndexes(lIndex) = l Then
            blnMissing = False
            Exit For
         End If
      Next lIndex
      If blnMissing Then
         Err.Raise vbObjectError + 2055, Ambient.DisplayName & ".ReorderColumns", "Index is missing elements."
         Exit For
         Exit Sub
      End If
   Next l
   
   Dim cc As ntColumn
   
    For Each cc In m_ntColumns
      For l = 0 To UBound(arrLonNewIndexes)
         If cc.index = l Then
            cc.index = arrLonNewIndexes(l)
            Exit For
         End If
      Next l
    Next cc
         
   m_ntColumns.SetIndexOrder

   If bRedraw And m_blnShown Then Call Me.Refresh

End Sub

'This resets a column to the desired state
Private Sub ResetColumn(vKey As Variant, pVisible As Boolean)
  Dim lonRows As Long
  Dim blnPicture As Boolean
  Dim vFilter As Variant
  Dim blnEnabled As Boolean
  Dim strColName As String
  Dim lonIndex As Long

On Error Resume Next

   If Not Ambient.UserMode Then Exit Sub
   If Not m_ntColumns.Exists(vKey) Then Exit Sub
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub

   m_blnLoading = True

   m_ntColumns(vKey).Visible = pVisible
   blnPicture = (m_ntColumns(vKey).ColFormat = nfgBooleanCheckBox)
   blnEnabled = m_ntColumns(vKey).Enabled
   strColName = m_ntColumns(vKey).Name
   lonIndex = m_ntColumns(vKey).index + Grid1.FixedCols

   m_RsFiltered.movefirst

   With Grid1

      .Redraw = False

      If pVisible = True Then

         ' Add Header Text Back to grid
         If .FixedRows > 0 Then

            If m_blnUseFieldNamesAsHeader = True Then
               .TextMatrix(0, lonIndex) = m_ntColumns(vKey).Name
            Else
               .TextMatrix(0, lonIndex) = m_ntColumns(vKey).HeaderText
            End If

         End If

         .ColWidth(lonIndex) = m_ntColumns(vKey).Width

         m_RsFiltered.AbsolutePosition = GetScrollValue() + 1

         If blnPicture Then

            .Col = lonIndex

            For lonRows = .FixedRows To .Rows - 1
               .Row = lonRows
               .CellPictureAlignment = 4
               Call SetPicture(StrComp(m_RsFiltered.Fields(strColName).Value, "True", vbTextCompare) = 0, blnEnabled)
               m_RsFiltered.MoveNext
            Next lonRows

         Else

            m_RsFiltered.movefirst

            For lonRows = .FixedRows To .Rows - 1
               .TextMatrix(lonRows, lonIndex) = m_RsFiltered.Fields(strColName).Value
               m_RsFiltered.MoveNext
            Next lonRows

         End If

      Else

         .Col = lonIndex

         ' Get rid of all text in row
         .FillStyle = flexFillRepeat
         .Row = 0
         .RowSel = .Rows - 1
         .Text = ""
         If blnPicture Then Set .CellPicture = Nothing
         m_ntColumns(vKey).Width = .ColWidth(lonIndex)
         .ColWidth(lonIndex) = 0

      End If

      .Redraw = m_blnRedraw

      .FillStyle = flexFillSingle

   End With

   m_RsFiltered.Filter = 0

   If m_blnAutoSizeColumns = True Then Call AutosizeGridColumns

   SetGridRows

   m_blnLoading = False

End Sub

'**
'Refreshes the Grid and the text values, and redraws entire Grid.
Public Sub Refresh()
Attribute Refresh.VB_HelpID = 1210
   
   If IsUnbound Then Exit Sub
   If Not m_blnShown Then Exit Sub

On Error Resume Next
      
   If bEditing Then Grid1_GotFocus
    
   m_blnMoving = True
   
   SaveRestorePrevGrid Grid1
   
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Grid1.Redraw = False
   
   If Ambient.UserMode And m_blnShown Then
      ' Initialize the grid
      m_ntColumns.SetIndexOrder
      Call BuildGridFromColumns(m_ntColumns)
      GridMousePointer = vbDefault
      Screen.MousePointer = vbDefault
      m_blnShown = True
      If m_blnAutoSizeColumns Then Call AutosizeGridColumns
      SetGridRows
      Call SetTotalPositions
      Dim nCol As ntColumn
      For Each nCol In m_ntColumns
         If nCol.ShowTotal Then
            Call CalcTotals
            Exit For
         End If
      Next nCol
      Call FillTextmatrix(vScroll.Value)
   End If
     
   SaveRestorePrevGrid Grid1, False
   CalcPaintedArea
   
   Grid1.Redraw = m_blnRedraw
   
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
   
   m_blnMoving = False
   
   If m_blnHasFocus Then
      On Error Resume Next
      Grid1.SetFocus
   End If
     
   DoEvents
   
End Sub

'Sub to remove duplicates in a sorted array
Private Sub RemoveStringArrayDupes(ByRef strArray() As String, Optional ByVal pCaseSensitive As Boolean = False)
  Dim strCompare As String
  Dim lonNewPos As Long
  Dim lonCounter As Long
  Dim intCompare As Integer

  If pCaseSensitive = True Then intCompare = vbTextCompare

On Error Resume Next

    'Loop through the array
  For lonCounter = 0 To UBound(strArray)
    'If the variable is not the same as the array element
    If StrComp(strCompare, strArray(lonCounter), intCompare) <> 0 Then
      'Set the array element to the current element
      strArray(lonNewPos) = strArray(lonCounter)
      'Increment the counter for the next array position
      lonNewPos = lonNewPos + 1
      'Set the variable
      strCompare = strArray(lonCounter)
    End If
  Next lonCounter

  'Knock the array down to the actual number of non-duplicate items
  If lonNewPos > 0 Then ReDim Preserve strArray(lonNewPos - 1)

End Sub

'**
'Removes any sorts applied to the Grid, and restores default sorting
Public Sub RemoveSort()
Attribute RemoveSort.VB_HelpID = 1220
   Dim i As Integer
   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
On Error Resume Next
   If bEditing Then Grid1_GotFocus
   m_RsFiltered.Sort = m_RSMaster.Sort
   For i = 0 To m_ntColumns.count - 1
      m_ntColumns(i).Sorted = nfgSortNone
   Next i
   Grid1.Redraw = False
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   m_blnLoading = True
   SetScrollValue 0
   FillTextmatrix 0
   m_blnLoading = False
   Grid1.Redraw = m_blnRedraw
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub

'**
'Restores the Grid.ColumnsLayouts back to their default Design-Time state. Any changes made by the user
'will be discarded.
Public Sub ResetColumns()
Attribute ResetColumns.VB_HelpID = 1230

   If m_ntColumns.count > 0 Then
      Grid1.Redraw = False
      m_ntColumns.Reset
      Refresh
      RaiseEvent OnColReorder
   End If

End Sub

Private Sub ResetText()
  Dim lonFromRow As Long
  Dim lonToRow As Long
  Dim lonRows As Long
  Dim lonStartRow As Long
  Dim pCol As ntColumn

On Error Resume Next

  lonStartRow = GetScrollValue()

  With Grid1

    Set pCol = m_ntColumns(m_lonEditCol)

    If m_lonEditRow > m_lonEditRowSel Then
      lonFromRow = m_lonEditRowSel
      lonToRow = m_lonEditRow
    Else
      lonFromRow = m_lonEditRow
      lonToRow = m_lonEditRowSel
    End If

   For lonRows = lonFromRow To lonToRow
      m_RsFiltered.AbsolutePosition = lonRows + 1
      .TextMatrix(lonRows - lonStartRow + .FixedRows, pCol.index + .FixedCols) = pCol.FormatValue(m_RsFiltered.Fields(pCol.Name).Value)
   Next lonRows

  End With

  Set pCol = Nothing

End Sub

Public Function RowFromBookmark(ByVal vBookmark As Variant) As Long
Attribute RowFromBookmark.VB_HelpID = 1240
     
   RowFromBookmark = -1
      
   If IsUnbound Then Exit Function
   If Not HasRecords Then Exit Function
   
On Error GoTo RFB_Err
   m_RsFiltered.Bookmark = vBookmark
   RowFromBookmark = m_RsFiltered.AbsolutePosition - 1

Exit Function

RFB_Err:
   RowFromBookmark = -1
End Function

'Sub to calculate rows, cols, scrollbars in grid based on rowheight,colwidth, gridsize
Private Sub SetGridRows()
   Dim i As Integer
   Dim j As Integer
   Dim FixedRowHeight As Long
   Dim intLastLeftCol As Integer
   Dim intFixedColWidth As Long
   Dim intVScrollWidth As Integer
   Dim intHScrollheight As Integer
   Dim intTotalHeight As Integer
   Dim intRowsInGrid As Integer
   Dim intMaxRows As Integer
   Dim intNewMaxRows As Integer
   Dim intLastCol As Integer
   Dim GridWkHeight As Long
   Dim GridWkWidth As Long
   Dim tempWidth As Long
   Dim bOldVertVisible As Boolean
   Dim bOldHorzVisible As Boolean
   
   bOldVertVisible = m_bln_NeedVertScroll
   bOldHorzVisible = m_bln_NeedHorzScroll
   
   If m_RSMaster Is Nothing Or m_RsFiltered Is Nothing Then
      m_bln_NeedHorzScroll = False
      hScroll.Value = 0
      m_bln_NeedHorzScroll = False
      vScroll.Visible = False
      vScroll.Value = 0
      m_bln_NeedVertScroll = False
      bRedrawFlag = ((bOldVertVisible <> m_bln_NeedVertScroll) Or (bOldHorzVisible <> m_bln_NeedHorzScroll))
      Exit Sub
   End If

On Error Resume Next

   'Get Total width of all cols in Grid
   For j = 0 To Grid1.Cols - 1
      tempWidth = tempWidth + Grid1.ColWidth(j)
   Next j

   'Create array holding visible columns to use for scrolling and setting scroll bars
   ReDim m_arrVisCols((m_ntColumns.count - m_ntColumns.HiddenColumns) - 1)

   Dim pCol As ntColumn

   j = 0

   For i = 0 To m_ntColumns.count - 1
      If m_ntColumns(i).Visible Then
         m_arrVisCols(j) = i
         j = j + 1
      End If
   Next i

   If m_blnTotalRow Then intTotalHeight = m_def_ShowTotalHeight
   If m_blnHeaderRow Then FixedRowHeight = m_sngRowHeightFixed

   With Grid1

      ' If there are more records than rows, need Vertical Scroll bar
      m_bln_NeedVertScroll = (m_RsFiltered.RecordCount > MaxNonFixedRows(False))
      
      'If we need the vert scroll, see if property setting overrides it
      If m_bln_NeedVertScroll Then m_bln_NeedVertScroll = ((m_eScrollBars = nfgScrollBoth) Or (m_eScrollBars = nfgScrollVertical))
      
      'If the last col is past the grid width, need horz scroll
      m_bln_NeedHorzScroll = (tempWidth > UserControl.ScaleWidth - (Abs(m_bln_NeedVertScroll) * 225)) Or _
                           (Grid1.LeftCol > Grid1.FixedCols)
      
      'If we need the vert scroll, see if property setting overrides it
      If m_bln_NeedHorzScroll Then m_bln_NeedHorzScroll = ((m_eScrollBars = nfgScrollBoth) Or (m_eScrollBars = nfgScrollHorizontal))
            
      
      Dim lNumRows As Long

      ' Need to check again, as adding either scroll bar changes the grid dimensions
      If m_RsFiltered.RecordCount > MaxNonFixedRows(m_bln_NeedHorzScroll) Then
         m_bln_NeedVertScroll = ((m_eScrollBars = nfgScrollBoth) Or (m_eScrollBars = nfgScrollVertical))
         lNumRows = MaxNonFixedRows(m_bln_NeedHorzScroll)
      Else
         If m_RsFiltered.RecordCount = 0 Then
            lNumRows = 1
         Else
            lNumRows = m_RsFiltered.RecordCount
         End If
      End If

      .Rows = lNumRows + .FixedRows
                 
      If tempWidth > UserControl.ScaleWidth - (Abs(m_bln_NeedVertScroll) * 225) Or _
               (Grid1.LeftCol > Grid1.FixedCols) Then
         m_bln_NeedHorzScroll = ((m_eScrollBars = nfgScrollBoth) Or (m_eScrollBars = nfgScrollHorizontal))
      End If

      .Width = UserControl.ScaleWidth - (Abs(m_bln_NeedVertScroll) * 225)
      .Height = UserControl.ScaleHeight - (Abs(m_bln_NeedHorzScroll) * 225)

      If m_blnHeaderRow Then Grid1.RowHeight(0) = m_sngRowHeightFixed
            
      For i = .FixedRows To .Rows - 1
         .RowHeight(i) = m_sngRowHeight
      Next i

      If m_bln_NeedHorzScroll Then

         'Now determine which col would be the last one to scroll to to enable all the
         'columns to be visible in the grid
         intFixedColWidth = 0
         tempWidth = 0

         If .FixedCols > 0 Then intFixedColWidth = m_sngRecordSelectorWidth

         For i = UBound(m_arrVisCols) To 0 Step -1
            tempWidth = tempWidth + .ColWidth(m_arrVisCols(i) + .FixedCols)
            'If we are wider than the grid, go back one, set variable and exit
            If tempWidth > (.Width - intFixedColWidth) Then
               If i < UBound(m_arrVisCols) Then
                  intLastLeftCol = i + 1
               Else
                  intLastLeftCol = i
               End If
               Exit For
            End If
         Next i

      End If

   End With

   hScroll.Visible = m_bln_NeedHorzScroll
   vScroll.Visible = m_bln_NeedVertScroll
      
   If m_bln_NeedHorzScroll Then
      If Grid1.Cols - intLastLeftCol < 10 Then
         hScroll.LargeChange = 1
      ElseIf Grid1.Cols - intLastLeftCol < 20 Then
         hScroll.LargeChange = 2
      Else
         hScroll.LargeChange = 3
      End If
      'The max val is set so that the scroll max + grid rows will end the same as recordset
      hScroll.Max = intLastLeftCol + (Grid1.LeftCol - Grid1.FixedCols)
   End If

   If m_bln_NeedVertScroll Then
      'large change is for clicking in scroll bar, or page down and page up
       vScroll.LargeChange = (Grid1.Rows - Grid1.FixedRows)
       'The max val is set so that the scroll max + grid rows will end the same as recordset
       vScroll.Max = m_RsFiltered.RecordCount - ((Grid1.Rows - Grid1.FixedRows))
   End If
     
   Call SetTotalPositions
   bRedrawFlag = ((bOldVertVisible <> m_bln_NeedVertScroll) Or (bOldHorzVisible <> m_bln_NeedHorzScroll))
   
End Sub

Private Function MaxNonFixedRows(ByVal bHorzScrollVisible As Boolean) As Long
   Dim lonTotalRowheight As Long
   Dim lonFixedRowHeight As Long

On Error Resume Next

   If m_blnTotalRow Then lonTotalRowheight = m_def_ShowTotalHeight
   If m_blnHeaderRow Then lonFixedRowHeight = m_sngRowHeightFixed

   'Calc the max rows that would fit in grid WITHOUT Horz scroll bar
   MaxNonFixedRows = _
     (CLng((((UserControl.ScaleHeight - (Abs(bHorzScrollVisible) * 255)) - lonTotalRowheight) - lonFixedRowHeight) \ m_sngRowHeight))
      
End Function

Private Sub SetPicture(ByVal pChecked As Boolean, Optional ByVal pEnabled As Boolean = True)
     
   Grid1.CellPictureAlignment = flexAlignCenterCenter
         
  If pChecked = True Then
    If pEnabled Then
      Set Grid1.CellPicture = picCheck.Image
    Else
      Set Grid1.CellPicture = picCheckDis.Image
    End If
  Else
    If pEnabled Then
      Set Grid1.CellPicture = picUnCheck.Image
    Else
      Set Grid1.CellPicture = picUnCheckDis.Image
    End If
  End If

End Sub

Private Sub SetRowColor(ByVal pRow As Long, ByVal pCol As Long, ByVal pColor As OLE_COLOR)

   With Grid1

      .Row = pRow

      If m_blnColorByRow = True Then
         .Col = .FixedCols
         .colSel = .Cols - 1
         .FillStyle = flexFillRepeat
         .CellBackColor = pColor
         .FillStyle = flexFillSingle
      Else
         .FillStyle = flexFillSingle
         .Col = pCol
         .CellBackColor = pColor
      End If

   End With

End Sub

Private Sub SetScrollValue(ByVal lVal As Long)

On Error Resume Next

   If m_bln_NeedVertScroll Then
      m_blnChanged = True
      If lVal > vScroll.Max Then
         m_UpdateScrollValue = vScroll.Max
      Else
         m_UpdateScrollValue = lVal
      End If
   Else
      m_blnChanged = True
      m_UpdateScrollValue = 0
   End If

   If Not m_blnRedraw Then Exit Sub
   If m_blnLoading Then Exit Sub

   If m_bln_NeedVertScroll Then
      If lVal > vScroll.Max Then
         vScroll.Value = vScroll.Max
      Else
         vScroll.Value = lVal
      End If
   End If

   If m_blnChanged Then
      FillTextmatrix m_UpdateScrollValue
      CalcPaintedArea
   End If

End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long

   Dim Red As Long
   Dim Green As Long
   Dim Blue As Long
   
On Error Resume Next

   Const ShiftAmt As Long = &H100
     
   Blue = ((Abs(Color) \ &H10000) Mod ShiftAmt) + Value
   Green = ((Abs(Color) \ &H100) Mod ShiftAmt) + Value
   Red = (Abs(Color) And &HFF) + Value
   
   If Value > 0 Then
      If Red > 255 Then Red = 255
      If Green > 255 Then Green = 255
      If Blue > 255 Then Blue = 255
   ElseIf Value < 0 Then
      If Red < 0 Then Red = 0
      If Green < 0 Then Green = 0
      If Blue < 0 Then Blue = 0
   End If
   
   ShiftColor = Red + 256& * Green + 65536 * Blue

End Function

'**
'This will show the Grids context menu at the current mouse pointer location,
'or the coordinates (x,y) passed in.
'@param        x Single. Optional.
'@param        y Single. Optional.
Public Function ShowMenu(Optional ByVal x As Single = -1, Optional ByVal y As Single = -1)
Attribute ShowMenu.VB_HelpID = 1250
   RaiseEvent BeforeShowMenu
   If x = -1 Or y = -1 Then
      PopupMenu mnuGrid
   Else
      PopupMenu mnuGrid, , x, y
   End If
End Function

Private Sub hScroll_Change()
   m_bScrollType = 1
   Timer1.Enabled = True
End Sub

Private Sub Horz_Scroll()
   Dim bRedraw As Boolean
   Static blnScrolling As Boolean
   Dim hVal As Long
   Dim tempWidth As Long
   Dim i As Integer
   Dim intFixedWidth As Long

On Error Resume Next
     
   If m_blnLoading Then Exit Sub
   If blnScrolling Then Exit Sub
   If Not (m_bln_NeedHorzScroll) Then Exit Sub
   
   If bEditing Then Grid1_GotFocus
   
   blnScrolling = True

   hVal = hScroll.Value

   If Grid1.FixedCols > 0 Then intFixedWidth = m_sngRecordSelectorWidth

   If Abs(hVal - m_PrevHorzValue) <> 1 Then

      If hVal > m_PrevHorzValue Then

         For i = m_PrevHorzValue To UBound(m_arrVisCols)
            hVal = i
            If Grid1.ColPos(m_arrVisCols(i) + Grid1.FixedCols) + Grid1.ColWidth(m_arrVisCols(i) + Grid1.FixedCols) > (Grid1.Width - intFixedWidth) Then Exit For
         Next i

         If hVal > hScroll.Max Then hVal = hScroll.Max

      ElseIf hVal < m_PrevHorzValue Then

         tempWidth = 0

         For i = (m_PrevHorzValue - 1) To 0 Step -1
            tempWidth = tempWidth + Grid1.ColWidth(m_arrVisCols(i) + Grid1.FixedCols)
            If tempWidth > (Grid1.Width - intFixedWidth) Then Exit For
            hVal = i
         Next i

         If hVal < 0 Then hVal = 0

      End If

   End If

   If Not hVal > m_arrVisCols(UBound(m_arrVisCols)) + Grid1.FixedCols Then
      hScroll.Value = hVal
      Grid1.Height = UserControl.Height
      Grid1.ScrollBars = flexScrollBarHorizontal
      Grid1.LeftCol = m_arrVisCols(hVal) + Grid1.FixedCols
      Grid1.ScrollBars = flexScrollBarNone
      Grid1.Height = UserControl.ScaleHeight
   End If

   If m_blnHeaderRow Then SetTotalPositions

   m_PrevHorzValue = hVal

   blnScrolling = False

End Sub

Private Sub vScroll_Change()
   m_bScrollType = 0
   Timer1.Enabled = True
End Sub

Private Sub Vert_Scroll()
   Dim bRedraw As Boolean
   Dim hVal As Long
   Dim tempWidth As Long
   Dim i As Integer
   Dim intFixedWidth As Long

On Error Resume Next
      
   If m_blnLoading Then Exit Sub
   If m_bScrolling Then Exit Sub
   If Not (m_bln_NeedVertScroll) Then Exit Sub
            
   If bEditing Then Grid1_GotFocus
   
   m_bScrolling = True

   If GetScrollValue() = m_PrevVertValue Then
      m_bScrolling = False
      Exit Sub
   End If

   If Not m_blnRedraw Then
      m_UpdateScrollValue = GetScrollValue()
      vScroll.Value = m_PrevVertValue
      m_bScrolling = False
      Exit Sub
   End If
            
   Grid1.Redraw = False

   m_blnIgnoreSel = True

   If Abs(GetScrollValue() - m_PrevVertValue) = 1 Then
      Call DrawNewRow(vScroll.Value < m_PrevVertValue)
   Else
      FillTextmatrix GetScrollValue()
   End If

   If Not m_blnManualScroll Then CalcPaintedArea

   m_blnIgnoreSel = False

   m_PrevVertValue = GetScrollValue()

   Grid1.Redraw = m_blnRedraw
   
   m_bScrolling = False

End Sub

Private Sub HScroll_Scroll()
   m_bScrollType = 1
   Timer1.Enabled = True
   'hScroll_Change
End Sub

Private Sub vScroll_Scroll()
   m_bScrollType = 0
   Timer1.Enabled = True
   'vScroll_Change
End Sub

Private Sub SaveRestorePrevGrid(ByRef sGrid As MSHFlexGrid, _
                                Optional ByVal bSave As Boolean = True, _
                                Optional ByVal bSaveRestoreRows As Boolean = True, _
                                Optional ByVal bSaveRestoreCols As Boolean = True)
  If bSave Then
      If bSaveRestoreCols Then
         m_lPrevCol = sGrid.Col
         m_lPrevColSel = sGrid.colSel
      End If
      If bSaveRestoreRows Then
         m_lPrevRow = sGrid.Row
         m_lPrevRowSel = sGrid.RowSel
      End If
   Else
      If bSaveRestoreCols Then
         sGrid.Col = m_lPrevCol
         sGrid.colSel = m_lPrevColSel
      End If
      If bSaveRestoreRows Then
         sGrid.Row = m_lPrevRow
         sGrid.RowSel = m_lPrevRowSel
      End If
   End If

End Sub

Private Sub ScrollTimer(ByVal ScrollType As Integer)

   If ScrollType = 1 Then
      ' If mouse is below the grid - means we are scrolling down
      If m_ScrollDown Then
         'If the current scroll value is less than the max
         If GetScrollValue() < vScroll.Max Then
            'Set the Selected Row and Column
            m_CurrRowSel = GetScrollValue() + vScroll.LargeChange
            m_CurrColSel = Grid1.MouseCol - Grid1.FixedCols
            If m_CurrColSel < 0 Then m_CurrColSel = 0
            ' Add one to the scroll bar position
            vScroll.Value = GetScrollValue() + 1
         End If
      ' If mouse is above the Grid - means they want to scroll UP
      Else
         'If the current scroll value is greater than the min
         If GetScrollValue() > vScroll.Min Then
            'Set the Selected Row and Column
            m_CurrRowSel = GetScrollValue() - 1
            m_CurrColSel = Grid1.MouseCol - Grid1.FixedCols
            If m_CurrColSel < 0 Then m_CurrColSel = 0
            'Subtract one from scroll bar position
            vScroll.Value = GetScrollValue() - 1
         End If
      End If
   Else
      If m_ScrollLeft Then
         If hScroll.Value > hScroll.Min Then
            hScroll.Value = hScroll.Value - 1
         End If
      Else
         If hScroll.Value < hScroll.Max Then
            hScroll.Value = hScroll.Value + 1
         End If
      End If
   End If

End Sub

Private Sub SetRecordsetsOnEdit(ByVal pRow As Long, ByVal pRowSel As Long, ByVal vValue As Variant)
  Dim strText As String
  Dim lonCol As Long
  Dim lonRowCounter As Long
  Dim lonColCounter As Long
  Dim lonFromRow As Long
  Dim lonToRow As Long
  Dim pCol As ntColumn
  Dim lonPrevHighlight As Long
  Dim dblnewTotal As Double
  Dim dblOldTotal As Double
  Dim bValid As Boolean
  Dim pCancel As Boolean
  Dim vBkmrk As Variant
  Dim strFilter As String
  Dim vPrevVal As Variant
  Dim blnChanged As Boolean
  Dim prevRow As Long
  Dim prevRowSel As Long
  Dim prevCol As Long
  Dim prevColSel As Long
   Dim bChecked As Boolean
   
On Error Resume Next

   m_blnLoading = True
      
   Set pCol = m_ntColumns(m_lonEditCol)

   If pCol.ColFormat <> nfgBooleanCheckBox Then
      Select Case pCol.EditType
         Case nfgTextBox
            txtEdit.Text = vbNullString
         Case nfgComboBox
            bValid = True
         Case Else

           Exit Sub

      End Select
   End If

   vPrevVal = vValue

   RaiseEvent BeforeEdit(ntEditField, m_PrevEditVal, vValue, pCol.Name, m_lonEditRowIDs, bValid, pCancel)

   If pCancel Then
      m_blnCancelEdit = True
      Call ResetText
      On Error Resume Next
      If m_blnHasFocus Then Grid1.SetFocus
      m_blnLoading = False
      m_blnCancelEdit = False
      Exit Sub
   End If

   blnChanged = (vValue <> vPrevVal) Or (pRow <> pRowSel)

   If m_lonEditRowSel < m_lonEditRow Then
      lonFromRow = m_lonEditRowSel
      lonToRow = m_lonEditRow
   Else
      lonFromRow = m_lonEditRow
      lonToRow = m_lonEditRowSel
   End If

   m_blnIgnoreSel = True

   With Grid1

      .Redraw = False

      For lonRowCounter = lonFromRow To lonToRow

         ' Set the recordset
         m_RsFiltered.AbsolutePosition = lonRowCounter + 1
         vBkmrk = m_RsFiltered.Bookmark
         m_RSMaster.Bookmark = vBkmrk
        
         If pCol.ShowTotal Then
            If IsNumeric(m_RsFiltered.Fields(pCol.Name).Value) Then
               dblOldTotal = dblOldTotal + m_RsFiltered.Fields(pCol.Name).Value
            End If
         End If

         m_RSMaster.Fields(pCol.Name).Value = vValue
         vValue = m_RSMaster.Fields(pCol.Name).Value
         m_RsFiltered.Fields(pCol.Name).Value = vValue
                 
         If pCol.ShowTotal Then
            If IsNumeric(vValue) Then dblnewTotal = dblnewTotal + CDbl(vValue)
         End If
                 
         If blnChanged Then

            .Row = lonRowCounter + .FixedRows
            If pCol.ColFormat = nfgBooleanCheckBox Then
               If IsNull(vValue) Or vValue = "" Then
                  bChecked = False
               Else
                  bChecked = CBool(vValue)
               End If
               .Col = pCol.index + .FixedCols
               .CellPictureAlignment = 4
               Call SetPicture(bChecked, pCol.Enabled)
            Else
               .TextMatrix(.Row, pCol.index + .FixedCols) = pCol.FormatValue(vValue)
            End If

         End If

      Next lonRowCounter

      .Row = m_CurrRow + Grid1.FixedRows
      .RowSel = m_CurrRowSel + Grid1.FixedRows
      .Col = m_CurrCol + Grid1.FixedCols
      .colSel = m_CurrColSel + Grid1.FixedCols

      .Redraw = m_blnRedraw

   End With

   m_blnIgnoreSel = False

   If pCol.ShowTotal Then
      lbltotal(pCol.index).Text = pCol.FormatValue(CDbl(lbltotal(pCol.index).Text) + (dblnewTotal - dblOldTotal))
   End If

   m_blnLoading = False

   RaiseEvent AfterEdit(ntEditField, vValue, pCol.Name, m_lonEditRowIDs)

   Set pCol = Nothing
   m_PrevEditVal = ""

End Sub

Private Sub SetTotalPositions()
   Dim loncols As Long
   Dim pCol As ntColumn
   Dim pTop As Long
   Dim pWidth As Long
   Dim pLeft As Long
   Dim pBorder As Long
   Dim i As Integer
   
On Error Resume Next

   If Ambient.UserMode Then

      If m_ntColumns Is Nothing Then Exit Sub
      If m_ntColumns.count = 0 Then Exit Sub
      
      If m_intTotalFloat Then
         pTop = Grid1.RowPos(Grid1.Rows - 1) + Grid1.RowHeight(Grid1.Rows - 1)
      Else
         pTop = (UserControl.ScaleHeight - m_def_ShowTotalHeight - (Abs(m_bln_NeedHorzScroll * 255)))
      End If

      With Grid1

         If m_ntColumns.count = 0 Then
            For loncols = 0 To lbltotal.UBound
               lbltotal(loncols).Left = -20000
               lbltotal(loncols).Visible = False
            Next loncols
         Else
            pLeft = 0
            If .FixedCols > 0 Then pLeft = .ColWidth(0)

            For loncols = .FixedCols To .Cols - 1

               If loncols - .FixedCols > lbltotal.UBound Then
               
                  Load lbltotal(loncols - .FixedCols)
                  lbltotal(loncols - .FixedCols).Text = ""
               End If
                             
               Set pCol = m_ntColumns(loncols - .FixedCols)

               If m_blnTotalRow And pCol.Visible = True Then
                     lbltotal(loncols - .FixedCols).Top = pTop
                     lbltotal(loncols - .FixedCols).Left = .ColPos(loncols)
                     lbltotal(loncols - .FixedCols).Width = .ColWidth(loncols)
                     lbltotal(loncols - .FixedCols).Height = 285
                     lbltotal(loncols - .FixedCols).Visible = lbltotal(loncols - .FixedCols).Left >= pLeft - 45
                     lbltotal(loncols - .FixedCols).ZOrder 0
               Else
                  lbltotal(loncols - .FixedCols).Visible = False
               End If
            Next loncols

         End If

      End With

   Else
             
      If m_blnTotalRow Then
         pTop = UserControl.ScaleHeight - m_def_ShowTotalHeight
         If m_eScrollBars = nfgScrollHorizontal Or m_eScrollBars = nfgScrollBoth Then
            pTop = pTop - 255
         End If
         If m_intTotalFloat Then pTop = pTop - 240
         pWidth = UserControl.ScaleWidth
         If m_eScrollBars = nfgScrollVertical Or m_eScrollBars = nfgScrollBoth Then
            pWidth = pWidth - 255
         End If
         lbltotal(0).Height = 285
         lbltotal(0).Width = pWidth
         lbltotal(0).Top = pTop
         lbltotal(0).Left = 0
         lbltotal(0).Visible = True
         lbltotal(0).ZOrder 0
      Else
         lbltotal(0).Left = -20000
      End If

   End If

End Sub

'Either pass in a Valid ADO Sort for the recordset, or the grid will build
' a sort based on the current Col and Colsel properties and the bAscending parameter
Public Sub SortGrid(Optional ByVal strsort As String = "", _
                    Optional ByVal lStartCol As Long = -1, _
                    Optional ByVal lEndCol As Long = -1, _
                    Optional ByVal bAscending As Boolean = True)
Attribute SortGrid.VB_HelpID = 1260

On Error GoTo Sort_Err
   
   Dim i As Integer
   Dim l As Long
   Dim bEvent As Boolean
   Dim arrSorts() As String
   Dim bCancel As Boolean
   Dim bSelChange As Boolean
   
   If IsUnbound Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
      
   bEvent = (Len(strsort) = 0)
   
   If m_blnShown Then If bEditing Then Grid1_GotFocus
    
   If Len(strsort) = 0 Then
      If lStartCol = -1 Then lStartCol = m_CurrCol
      If lEndCol = -1 Then lEndCol = m_CurrColSel
      If m_ntColumns.Exists(lStartCol) = False Or m_ntColumns.Exists(lEndCol) = False Then
         Err.Raise ERR_INVCOL, Ambient.DisplayName & ".SortGrid", MSG_INVCOL
         Exit Sub
      End If
      bCancel = False
      RaiseEvent BeforeSort(m_ntColumns(lStartCol).Name, m_ntColumns(lEndCol).Name, bCancel)
      If bCancel Then Exit Sub
      strsort = BuildGridRSSort(lStartCol, lEndCol, bAscending)
      m_CurrCol = m_ntColumns(lStartCol).index
      m_CurrColSel = m_ntColumns(lEndCol).index
   End If

   For i = 0 To m_ntColumns.count - 1
      If InStr(1, strsort, m_ntColumns(i).Name, vbTextCompare) = 0 Then
         m_ntColumns(i).Sorted = nfgSortNone
      End If
   Next i

   arrSorts = Split(strsort, ",", , vbTextCompare)

   For l = 0 To m_ntColumns.count - 1
      For i = 0 To UBound(arrSorts)
         If InStr(1, arrSorts(i), m_ntColumns(l).Name, vbTextCompare) > 0 Then
            If InStr(1, arrSorts(i), "DESC", vbTextCompare) > 0 Then
               m_ntColumns(l).Sorted = nfgSortDesc
            Else
               m_ntColumns(l).Sorted = nfgSortAsc
            End If
            Exit For
         End If
      Next i
   Next l

   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass

   m_RsFiltered.Sort = strsort
   
   m_blnDidSort = True
   
   Grid1.Redraw = False
   m_CurrRow = 0
   If m_CurrRowSel <> m_RsFiltered.RecordCount - 1 Then
      m_CurrRowSel = m_RsFiltered.RecordCount - 1
      bSelChange = True
   End If
   SetScrollValue 0
   m_LastRow = 0
   Grid1.Row = Grid1.FixedRows
   CalcPaintedArea
   Grid1.Redraw = m_blnRedraw
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
   
   If bSelChange Then RaiseEvent SelChange
   If bEvent Then RaiseEvent OnSort(m_ntColumns(lStartCol).Name, m_ntColumns(lEndCol).Name)
   
   m_blnDidSort = False
   
Exit Sub

Sort_Err:
   Err.Raise Err.Number, Ambient.DisplayName & ".SortGrid: " & Err.Source, "Invalid Sort"
End Sub

Private Function ValidateEditKey(ByVal KeyCode As Integer, ByRef Shift As Integer) As Boolean

   ValidateEditKey = False
     
   If Shift = 0 Then
      Select Case KeyCode
         Case 32, 45
            ValidateEditKey = (KeyCode = m_intEditKey)
         Case 113 To 121
            ValidateEditKey = (KeyCode = m_intEditKey)
      End Select
   ElseIf Shift = 2 Then
      If KeyCode = 69 Then
         ValidateEditKey = (KeyCode = m_intEditKey)
      End If
   End If

End Function
' Put the current column widths in the current layout
Friend Sub WriteColWidths()
    
   If m_blnMoving Then Exit Sub
    
   Dim pCol As ntColumn
        
   For Each pCol In m_ntColumns
      If pCol.Visible = True Then
         pCol.Width = Grid1.ColWidth(pCol.index + Grid1.FixedCols)
      End If
   Next pCol

End Sub

' Put the current column widths in the current layout
Friend Sub WriteRowHeights()
           
   If Grid1.RowHeight(Grid1.FixedRows) = m_sngRowHeight Then Exit Sub
   
   If m_sngRowHeightMin > 0 Then
      If Grid1.RowHeight(Grid1.FixedRows) < m_sngRowHeightMin Then
         m_sngRowHeight = m_sngRowHeightMin
      Else
         m_sngRowHeight = Grid1.RowHeight(Grid1.FixedRows)
      End If
   End If
   If m_sngRowHeightMax > 0 Then
      If Grid1.RowHeight(Grid1.FixedRows) > m_sngRowHeightMax Then
         m_sngRowHeight = m_sngRowHeightMax
      Else
         m_sngRowHeight = Grid1.RowHeight(Grid1.FixedRows)
      End If
   End If
         
   RecalcGrid
   
End Sub

Private Sub lbltotal_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim strText As String
On Error Resume Next
   If Button = vbRightButton Then
      ReleaseCapture
      strText = lbltotal(index).Text
      Do While TextWidth(strText) < lbltotal(index).Width - 700
         strText = strText & " "
      Loop
      mnuTotalShow.Caption = strText
      PopupMenu mnuTotal, , lbltotal(index).Left, lbltotal(index).Top + lbltotal(index).Height
   End If
End Sub

Private Sub mnuGridEdit_Click()
   Dim iShift As Integer
   m_blnDidMenu = True
   If m_intEditKey = 69 Then iShift = 2
   Grid1_KeyUp m_intEditKey, iShift
End Sub

Private Sub mnuGridFit_Click()
   Dim loncols As Long
   Dim lonFromCol As Long
   Dim lonToCol As Long

On Error Resume Next

   m_blnLoading = True

   With Grid1

      If .Col > .colSel Then
         lonFromCol = .colSel - .FixedCols
         lonToCol = .Col - .FixedCols
      Else
         lonFromCol = .Col - .FixedCols
         lonToCol = .colSel - .FixedCols
      End If

      For loncols = lonFromCol To lonToCol
         Call AutosizeGridColumns(loncols)
         m_ntColumns(loncols).Width = .ColWidth(loncols + .FixedCols)
      Next loncols

   End With

   m_blnDidMenu = True

   m_blnLoading = False

   Call SetGridRows
   Call SetTotalPositions

End Sub

Private Sub SubclassGrid()
   Dim msgIDs() As Long

On Error GoTo SubClass_Err

   If Not m_blnGridSubclassed = True Then

      m_blnGridSubclassed = True

      If Ambient.UserMode Then

         AttachMessage Me, Grid1.hwnd, WM_LBUTTONUP
         AttachMessage Me, Grid1.hwnd, WM_ERASEBKGND
         AttachMessage Me, Grid1.hwnd, WM_LBUTTONDOWN
         AttachMessage Me, Grid1.hwnd, WM_PAINT
                 
         AttachMessage Me, UserControl.hwnd, WM_TIMER

      End If

   End If

Exit Sub

SubClass_Err:
  Exit Sub
End Sub

'BEGIN EVENTS ========================================================================================
'             ========================================================================================
Private Sub txtEdit_Validate(Cancel As Boolean)
  RaiseEvent EditControlValidate(txtEdit.Text, m_ntColumns(Grid1.Col - Grid1.FixedCols).Name, Cancel)
End Sub

Private Sub txtEdit_LostFocus()
On Error Resume Next
    If txtEdit.Visible = False Then Exit Sub
    If Not m_blnCancelEdit Then
      txtEdit.Visible = False
      m_blnCancelEdit = True
      If Not ((Grid1.Row = Grid1.RowSel) And (m_PrevEditVal = txtEdit.Text)) Then
         Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, txtEdit.Text)
         Call ResetText
      End If
      m_blnCancelEdit = False
   End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
  Dim pCol As ntColumn
  Dim i As nfgValidationType

On Error Resume Next

  Set pCol = m_ntColumns(Grid1.Col - Grid1.FixedCols)

  Select Case pCol.ValidationType
    Case nfgValidateAlpha
      KeyAscii = ntGridValidate.ValidateAlphaKey(KeyAscii)
    Case nfgValidateNumeric
      KeyAscii = ntGridValidate.ValidateNumericKey((pCol.DecimalPlaces > 0), KeyAscii)
    Case nfgValidateAlphaNumeric
      KeyAscii = ntGridValidate.ValidateAlphaNumeric(KeyAscii)
  End Select

  Set pCol = Nothing

End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim pCol As ntColumn
   Dim pText As String
   Dim i As Integer
   Dim iCol As Integer
   Dim iRow As Long
   Dim bFoundCol As Boolean
   
On Error Resume Next
  
  If m_blnInitEdit Then
    m_blnInitEdit = False
    Exit Sub
  End If

   If KeyCode = vbKeyEscape Then
      m_blnCancelEdit = True
      txtEdit.Visible = False
      txtEdit.Text = vbNullString
      Call ResetText
      Grid1.SetFocus
      m_blnCancelEdit = False
      Exit Sub
   End If

  'Plus Key also?
  If KeyCode = 109 Then
      If IsNumeric(txtEdit.Text) Then
         txtEdit.Text = 0 - Abs(CSng(txtEdit.Text))
      End If
   End If

  Set pCol = m_ntColumns(Grid1.Col - Grid1.FixedCols)

   If pCol.ValidationType = nfgValidateNumeric Then
      Call ntGridValidate.EnforceNumericText(txtEdit, pCol.DecimalPlaces)
   End If

   Grid1.FillStyle = flexFillRepeat
   Grid1.Text = txtEdit.Text
   Grid1.FillStyle = flexFillSingle

   If KeyCode = vbKeyReturn Or _
      ((KeyCode = vbKeyUp) And (Grid1.Row >= Grid1.FixedRows)) Or _
      ((KeyCode = vbKeyDown) And (Grid1.Row < Grid1.Rows - 1)) Then
      m_blnCancelEdit = True
      Dim vValue As String
      vValue = txtEdit.Text
      txtEdit.Text = vbNullString
      txtEdit.Visible = False
      If Not ((Grid1.Row = Grid1.RowSel) And (m_PrevEditVal = vValue)) Then
         Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, vValue)
         Call ResetText
      End If
      Grid1.SetFocus
      m_blnCancelEdit = False
      
      iCol = -1
      
      If KeyCode = vbKeyDown Then
         If Not RowIsVisible(m_CurrRow + 1) Then
            vScroll.Value = vScroll.Value + 1
         End If
         Grid1.Row = Grid1.Row + 1
         Grid1_RowColChange
         Grid1_Click
      ElseIf KeyCode = vbKeyUp Then
         If Not RowIsVisible(m_CurrRow - 1) Then
            vScroll.Value = vScroll.Value - 1
         End If
         Grid1.Row = Grid1.Row - 1
         Grid1_RowColChange
         Grid1_Click
      ElseIf KeyCode = vbKeyReturn Then
         iRow = m_CurrRow
         If Shift = 0 Then
            If Grid1.Col < Grid1.Cols - 1 Then
               For i = Grid1.Col + 1 To Grid1.Cols - 1
                  If Grid1.ColWidth(i) > 0 Then
                     bFoundCol = True
                     iCol = i
                     Exit For
                  End If
               Next i
               If Not bFoundCol Then
                  If Grid1.Row < Grid1.Rows - 1 Then
                     For i = Grid1.FixedCols To Grid1.Col
                        If Grid1.ColWidth(i) > 0 Then
                           bFoundCol = True
                           iRow = iRow + 1
                           iCol = i
                           Exit For
                        End If
                     Next i
                  End If
               End If
               If bFoundCol Then
                  If Not RowIsVisible(iRow) Then
                     vScroll.Value = vScroll.Value + 1
                  End If
                  Grid1.Col = iCol
                  Grid1.colSel = iCol
                  Grid1_RowColChange
                  Grid1_Click
               End If
            End If
         Else
            If Grid1.Col >= Grid1.FixedCols Then
               For i = Grid1.Col - 1 To Grid1.FixedCols Step -1
                  If Grid1.ColWidth(i) > 0 Then
                     bFoundCol = True
                     iCol = i
                     Exit For
                  End If
               Next i
               If Not bFoundCol Then
                  If Grid1.Row >= Grid1.FixedRows Then
                     For i = Grid1.Cols - 1 To Grid1.Col Step -1
                        If Grid1.ColWidth(i) > 0 Then
                           bFoundCol = True
                           iRow = iRow - 1
                           iCol = i
                           Exit For
                        End If
                     Next i
                  End If
               End If
               If bFoundCol Then
                  If Not RowIsVisible(iRow) Then
                     vScroll.Value = vScroll.Value - 1
                  End If
                  Grid1.Col = iCol
                  Grid1.colSel = iCol
                  Grid1_RowColChange
                  Grid1_Click
               End If
            End If
         End If
      End If
   End If

End Sub

Private Sub cmbEdit_Validate(Cancel As Boolean)
  RaiseEvent EditControlValidate(cmbEdit.Text, m_ntColumns(Grid1.Col - Grid1.FixedCols).Name, Cancel)
End Sub

Private Sub cmbEdit_LostFocus()
   If Not cmbEdit.Visible Then Exit Sub
On Error Resume Next
   If Not m_blnCancelEdit Then
      cmbEdit.Visible = False
      m_blnCancelEdit = True
      If Not ((Grid1.Row = Grid1.RowSel) And (m_PrevEditVal = cmbEdit.Text)) Then
         Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, cmbEdit.Text)
      End If
      m_blnCancelEdit = False
   End If
End Sub

Private Sub cmbEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim iCol As Long
   Dim iRow As Long
   Dim i As Integer
   Dim bFoundCol As Boolean
   
On Error Resume Next

  If m_blnInitEdit Then
    m_blnInitEdit = False
    Exit Sub
  End If

  Select Case KeyCode
    Case vbKeyEscape
      m_blnCancelEdit = True
      m_blnLoading = True
      cmbEdit.Text = CStr(m_PrevEditVal & "")
      cmbEdit.Visible = False
      Call ResetText
      Grid1.SetFocus
      m_blnLoading = False
      m_blnCancelEdit = False
      Exit Sub
    Case vbKeyReturn
      m_blnCancelEdit = True
      cmbEdit.Visible = False
      Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, cmbEdit.Text)
      Call ResetText
      Grid1.SetFocus
      m_blnCancelEdit = False

      iCol = -1
      
      If KeyCode = vbKeyReturn Then
         iRow = m_CurrRow
         If Shift = 0 Then
            If Grid1.Col < Grid1.Cols - 1 Then
               For i = Grid1.Col + 1 To Grid1.Cols - 1
                  If Grid1.ColWidth(i) > 0 Then
                     bFoundCol = True
                     iCol = i
                     Exit For
                  End If
               Next i
               If Not bFoundCol Then
                  If Grid1.Row < Grid1.Rows - 1 Then
                     For i = Grid1.FixedCols To Grid1.Col
                        If Grid1.ColWidth(i) > 0 Then
                           bFoundCol = True
                           iRow = iRow + 1
                           iCol = i
                           Exit For
                        End If
                     Next i
                  End If
               End If
               If bFoundCol Then
                  If Not RowIsVisible(iRow) Then
                     vScroll.Value = vScroll.Value + 1
                  End If
                  Grid1.Col = iCol
                  Grid1.colSel = iCol
                  Grid1_RowColChange
                  Grid1_Click
               End If
            End If
         Else
            If Grid1.Col >= Grid1.FixedCols Then
               For i = Grid1.Col - 1 To Grid1.FixedCols Step -1
                  If Grid1.ColWidth(i) > 0 Then
                     bFoundCol = True
                     iCol = i
                     Exit For
                  End If
               Next i
               If Not bFoundCol Then
                  If Grid1.Row >= Grid1.FixedRows Then
                     For i = Grid1.Cols - 1 To Grid1.Col Step -1
                        If Grid1.ColWidth(i) > 0 Then
                           bFoundCol = True
                           iRow = iRow - 1
                           iCol = i
                           Exit For
                        End If
                     Next i
                  End If
               End If
               If bFoundCol Then
                  If Not RowIsVisible(iRow) Then
                     vScroll.Value = vScroll.Value - 1
                  End If
                  Grid1.Col = iCol
                  Grid1.colSel = iCol
                  Grid1_RowColChange
                  Grid1_Click
               End If
            End If
         End If
      End If
  End Select

End Sub

Private Sub cmbEdit_Click()

On Error Resume Next

   If m_blnInitEdit Then
      m_blnInitEdit = False
      Exit Sub
   End If

   If cmbEdit.ListIndex = -1 Then Exit Sub

   If m_blnLoading Then Exit Sub

   m_blnLoading = True

   Grid1.FillStyle = flexFillRepeat
   Grid1.Text = cmbEdit.Text
   Grid1.FillStyle = flexFillSingle

   m_blnLoading = False

End Sub

Private Sub Grid1_Click()
   Dim pCol As ntColumn
   Dim strList() As String
   Dim i As Long
   Dim j As Long
    
   If Not Ambient.UserMode Then Exit Sub
   
   If m_bln_NeedVertScroll Then vScroll.HideTip
   If m_bln_NeedHorzScroll Then hScroll.HideTip
   
On Error Resume Next

   RaiseEvent Click
      
   prevRow = Grid1.Row
   prevRowSel = Grid1.RowSel
   
   If m_RsFiltered Is Nothing Then Exit Sub
   If m_RsFiltered.RecordCount = 0 Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub

   If m_blnDidMenu Then
      m_blnDidMenu = False
      Exit Sub
   End If

   If txtEdit.Visible = True Or cmbEdit.Visible = True Then Exit Sub

   If m_blnLoading Then Exit Sub
   If Not m_blnLeftClick Then Exit Sub
   If Not m_blnAllowEdit Then Exit Sub
   If Grid1.MouseRow < Grid1.FixedRows Then Exit Sub
   If Grid1.MouseCol < Grid1.FixedCols Then Exit Sub
   If Grid1.Col <> Grid1.colSel Then Exit Sub
   If Grid1.Row <> Grid1.RowSel Then Exit Sub
   If m_ntColumns(Grid1.Col - Grid1.FixedCols).ColFormat = nfgBooleanCheckBox Then Exit Sub
   If Not m_ntColumns(Grid1.Col - Grid1.FixedCols).Enabled Then Exit Sub

   m_lonEditCol = Grid1.Col - Grid1.FixedCols
   m_lonEditRow = m_CurrRow
   m_lonEditRowSel = m_CurrRow

   Set pCol = m_ntColumns(m_lonEditCol)

   ' Build an array of all rows being edited to pass as Param in BeforeEdit event
   ReDim m_lonEditRowIDs(0)
   m_lonEditRowIDs(0) = m_CurrRow

   Select Case Grid1.ColAlignment(Grid1.MouseCol)
      Case 1
         txtEdit.Alignment = 0
      Case 4
         txtEdit.Alignment = 2
      Case 7
         txtEdit.Alignment = 1
   End Select

   Select Case pCol.EditType
      Case nfgTextBox
         Grid1.Highlight = flexHighlightAlways
         Call MSFlexGridEdit(txtEdit, 32)
      Case nfgComboBox
         Grid1.Highlight = flexHighlightAlways
         cmbEdit.Clear
         If Len(pCol.ComboList) > 0 Then
            strList = Split(pCol.ComboList, ",", , vbTextCompare)
            For i = 0 To UBound(strList)
               cmbEdit.AddItem strList(i)
            Next i
         End If
         Call MSFlexGridCombo(cmbEdit, 32)

   End Select

   Set pCol = Nothing

End Sub

Private Sub Grid1_DblClick()
   Dim pCol As ntColumn
   Dim strFilter As String
   Dim vBkmrk As Variant
   Dim i As Integer
   
On Error Resume Next
         
   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
   If m_blnLoading Then Exit Sub
   If IsMouseOutOfGrid Then Exit Sub

   With Grid1

      ' IF Dbl-Click was on Fixed Row then see if we should sort
      If .MouseRow < .FixedRows Then
         If .MouseCol >= (.FixedCols) Then
            If m_blnAllowSort Then
               Select Case m_ntColumns(.MouseCol - .FixedCols).Sorted
                  Case nfgSortNone, nfgSortDesc
                     Call SortGrid(, .MouseCol - .FixedCols)
                  Case nfgSortAsc
                     Call SortGrid(, .MouseCol - .FixedCols, , False)
               End Select
               m_blnDidSort = True
            End If
         End If
      Else
         ' IF Dbl-Click was not on Fixed Row or fixed Col, then we either filter or edit checkboxes
         If .MouseCol >= .FixedCols Then

            Set pCol = m_ntColumns(.MouseCol - .FixedCols)

            If m_blnAllowEdit = False Or pCol.Enabled = False Then

               If m_blnDblClickFilter Then

                  If m_blnAllowFilter = True Then

                     Dim arrFilters() As String

                     arrFilters = BuildGridRSFilterValues(pCol.Name, m_CurrRow, m_CurrRow, True)
                     GridMousePointer = vbHourglass
                     Screen.MousePointer = vbHourglass
                     
                     If ApplyFilter(pCol.Name, arrFilters, True) Then
                        Grid1.Redraw = False
                        Call SetGridRows
                        m_CurrColSel = m_CurrCol
                        m_CurrRow = 0
                        m_CurrRowSel = 0
                        SetScrollValue 0
                        m_LastRow = 0
                        Grid1.FocusRect = m_intFocusRect
                        Grid1.Highlight = m_intHighlight
                        Grid1.Redraw = m_blnRedraw
                        Grid1.SetFocus
                        Dim nf As ntFilter
                        Set nf = m_colFilters.NewFilter
                        nf.FieldName = pCol.Name
                        nf.Include = True
                        nf.Values = arrFilters
                        RaiseEvent OnFilter(nf)
                        Set nf = Nothing
                     End If
                     GridMousePointer = vbDefault
                     Screen.MousePointer = vbDefault

                  End If

               End If

            ElseIf m_blnAllowEdit = True Then

               If pCol.ColFormat = nfgBooleanCheckBox And pCol.Enabled = True Then
                  
                  Dim vOldValue As Variant
                  Dim blnChecked As Boolean
                  Dim bCancel As Boolean
                  Dim bValid As Boolean
                  Dim arrRows(0) As Long
                  Dim newValue As Variant

                  arrRows(0) = m_CurrRow
                  
                  m_RsFiltered.AbsolutePosition = m_CurrRow + 1
                  m_RSMaster.Bookmark = m_RsFiltered.Bookmark
                  
                  vOldValue = m_RSMaster.Fields(pCol.Name).Value
                  
                  If IsNumeric(vOldValue) Then
                     If vOldValue = 0 Then
                        newValue = -1
                     Else
                        newValue = 0
                     End If
                  ElseIf CBool(vOldValue) = True Then
                     newValue = False
                  Else
                     newValue = True
                  End If
                  
                  bValid = True

                  RaiseEvent BeforeEdit(ntEditField, CStr(vOldValue), newValue, pCol.Name, arrRows, bValid, bCancel)

                  If bCancel Then Exit Sub
                  
                  m_RSMaster.Fields(pCol.Name).Value = newValue
                  newValue = m_RSMaster.Fields(pCol.Name).Value
                  m_RsFiltered.Fields(pCol.Name).Value = newValue
                  
                  blnChecked = CBool(newValue)

                  .Redraw = False

                  Call SetPicture((blnChecked), True)
                                                                
                  If pCol.UseCriteria Then
                     If m_blnColorByRow = True Then
                        m_blnIgnoreSel = True
                        m_blnLoading = True
                        .Col = .FixedCols
                        .colSel = .Cols - 1
                        .FillStyle = flexFillRepeat
                        .CellBackColor = CheckRowColCriteria(m_CurrRow)
                        .FillStyle = flexFillSingle
                        .Col = m_CurrCol + .FixedCols
                        If pCol.BackColor <> -1 Or pCol.ForeColor <> -1 Then
                           .CellBackColor = pCol.BackColor
                           .CellForeColor = pCol.ForeColor
                        End If
                        m_blnIgnoreSel = False
                        m_blnLoading = False
                     Else
                        .CellBackColor = CheckRowColCriteria(m_CurrRow, m_CurrCol)
                     End If
                  End If

                  .Redraw = True

                  DoEvents

                  RaiseEvent AfterEdit(ntEditField, CStr(blnChecked), pCol.Name, arrRows)

               End If

            End If

         End If

      End If

   End With

   RaiseEvent DblClick

End Sub

Private Sub Grid1_EnterCell()
   RaiseEvent EnterCell
End Sub

Private Function bEditing() As Boolean
   bEditing = (txtEdit.Visible Or cmbEdit.Visible)
End Function

Private Sub Grid1_GotFocus()
   Dim vValue As Variant

On Error Resume Next
   
   If m_bln_NeedVertScroll Then vScroll.HideTip
   If m_bln_NeedHorzScroll Then hScroll.HideTip
   
   m_blnHasFocus = True

   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
   If m_blnLoading Then Exit Sub
   'If IsMouseOutOfGrid Then Exit Sub

   If bEditing Then

      Select Case True
         Case txtEdit.Visible
            vValue = txtEdit.Text
            txtEdit.Visible = False
         Case cmbEdit.Visible
            vValue = cmbEdit.Text
            cmbEdit.Visible = False
         Case Else
            Exit Sub
      End Select

      Grid1.Redraw = False

      GridMousePointer = vbHourglass
      Screen.MousePointer = vbHourglass

      Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, vValue)

      If m_bln_NeedVertScroll Then
         Call FillTextmatrix(GetScrollValue())
      Else
         Call FillTextmatrix(0)
      End If

      GridMousePointer = vbDefault
      Screen.MousePointer = vbDefault
      
      CalcPaintedArea
      Grid1.SetFocus
      
      Grid1.Redraw = m_blnRedraw

  Else
'      CalcPaintedArea
'      Grid1.SetFocus
  End If

End Sub

Private Sub Grid1_LeaveCell()
   RaiseEvent LeaveCell
End Sub

'RowColChange only happens when you are selecting something new
Private Sub Grid1_RowColChange()
   
On Error Resume Next
   
   Debug.Print "RowColChange"
   
   If m_blnLoading Then Exit Sub
   If m_blnIgnoreRCChange Then Exit Sub
   
   m_CurrRow = (GetScrollValue() + (Grid1.Row - Grid1.FixedRows))
   m_CurrRowSel = (GetScrollValue() + (Grid1.RowSel - Grid1.FixedRows))
   m_CurrCol = Grid1.Col - Grid1.FixedCols
   m_CurrColSel = Grid1.colSel - Grid1.FixedCols
   If m_CurrColSel < 0 Then m_CurrColSel = 0
   If m_CurrCol < 0 Then m_CurrCol = 0
   
   Grid1.Highlight = m_intHighlight
   Grid1.FocusRect = m_intFocusRect

   If (Grid1.Col - Grid1.FixedCols) <> m_LastCol Then

      If m_RsFiltered Is Nothing Or m_ntColumns.count = 0 Then

         RaiseEvent ColChange(Nothing)
         m_blnDoSel = True
         
      Else

         RaiseEvent ColChange(m_ntColumns(Grid1.Col - Grid1.FixedCols))
         m_blnDoSel = True
         
      End If

   End If

   If m_CurrRow <> m_LastRow Then
      RaiseEvent RowChange(m_CurrRow)
      m_blnDoSel = True
   End If
   
   m_LastRow = m_CurrRow
   m_LastCol = m_CurrCol
      
End Sub

Private Sub Grid1_SelChange()
            
   Debug.Print "SelChange"
   
   If m_blnIgnoreSel Then Exit Sub
        
On Error Resume Next
   
   m_blnDoSel = False
     
   If Grid1.colSel < Grid1.FixedCols Then Grid1.colSel = Grid1.FixedCols
   If Grid1.RowSel < Grid1.FixedRows Then Grid1.RowSel = Grid1.FixedRows
   m_CurrColSel = Grid1.colSel - Grid1.FixedCols
   If m_blnDidSort Then
      m_blnDidSort = False
   Else
      m_CurrRowSel = (GetScrollValue() + (Grid1.RowSel - Grid1.FixedRows))
   End If
   
   If m_CurrColSel < 0 Then m_CurrColSel = 0
   
   RaiseEvent SelChange
         
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
  
   RaiseEvent KeyDown(KeyCode, Shift)
   
On Error Resume Next

   If KeyCode = 255 Then Exit Sub
   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
   If m_blnLoading Then Exit Sub
   
   Select Case KeyCode

      Case 33  'PageUp

         If (GetScrollValue() - vScroll.LargeChange) >= 0 Then
            m_CurrRow = (GetScrollValue() - vScroll.LargeChange)
            m_CurrRowSel = m_CurrRow
            vScroll.Value = (GetScrollValue() - vScroll.LargeChange)
            RaiseEvent RowChange(m_CurrRow)
         Else
            m_CurrRow = 0
            m_CurrRowSel = 0
            vScroll.Value = 0
            RaiseEvent RowChange(m_CurrRow)
         End If
         
         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0
         
         m_blnPaging = True

      Case 34 'Page Down

         If (GetScrollValue() + vScroll.LargeChange) <= vScroll.Max Then
               m_CurrRow = (GetScrollValue() + vScroll.LargeChange)
               m_CurrRowSel = m_CurrRow
               vScroll.Value = (GetScrollValue() + vScroll.LargeChange)
               RaiseEvent RowChange(m_CurrRow)
           Else
               m_CurrRow = vScroll.Max
               m_CurrRowSel = m_CurrRow
               vScroll.Value = vScroll.Max
               RaiseEvent RowChange(m_CurrRow)
            End If

         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0

         m_blnPaging = True

      Case 35  'End
         If Shift = 0 Then
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Max
               hScroll.Value = hScroll.Max
            End If
            If m_CurrColSel <> m_CurrCol Then
               m_CurrColSel = m_CurrCol
               RaiseEvent SelChange
            End If
         ElseIf Shift = 1 Then
            m_CurrColSel = m_arrVisCols(UBound(m_arrVisCols))
            Grid1.colSel = Grid1.Cols - 1
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Max
               hScroll.Value = hScroll.Max
            End If
          
         'Ctrl + END puts you at bottom right corner
         ElseIf Shift = 2 Then
            m_CurrRow = m_RsFiltered.RecordCount - 1
            m_CurrRowSel = m_CurrRow
            m_CurrCol = m_arrVisCols(UBound(m_arrVisCols))
            m_CurrColSel = m_CurrCol
            If m_bln_NeedVertScroll Then vScroll.Value = vScroll.Max
            If m_bln_NeedHorzScroll Then
               'm_PrevHorzValue = hScroll.Max
               hScroll.Value = hScroll.Max
            End If
            Grid1.Row = Grid1.Rows - 1
            Grid1.RowSel = Grid1.Rows - 1
            Grid1.Col = m_CurrCol + Grid1.FixedCols
            Grid1.colSel = Grid1.Col
            RaiseEvent ColChange(m_ntColumns(m_CurrCol))
            RaiseEvent RowChange(m_CurrRow)
           
         'Shift + Control + End key
         ElseIf Shift = 3 Then
            m_CurrRowSel = m_RsFiltered.RecordCount - 1
            m_CurrColSel = m_ntColumns.count - 1
            If m_bln_NeedVertScroll Then vScroll.Value = vScroll.Max
            If m_bln_NeedHorzScroll Then
               'm_PrevHorzValue = hScroll.Max
               hScroll.Value = hScroll.Max
            End If
            Grid1.Row = Grid1.FixedRows
            Grid1.Col = Grid1.FixedCols
            Grid1.RowSel = Grid1.Rows - 1
            Grid1.colSel = Grid1.Cols - 1
            RaiseEvent SelChange
         End If
         
         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0

         m_blnPaging = True

      Case 36 'Home
         If Shift = 0 Then
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = 0
               hScroll.Value = 0
            End If
            If m_CurrColSel <> m_CurrCol Then
               m_CurrColSel = m_CurrCol
               RaiseEvent SelChange
            End If
         'Ctrl + Home puts you at top left corner
         ElseIf Shift = 2 Then
            m_CurrRow = 0
            m_CurrRowSel = m_CurrRow
            m_CurrCol = m_arrVisCols(LBound(m_arrVisCols))
            m_CurrColSel = m_CurrCol
            m_LastRow = 0
            m_LastCol = m_CurrCol
            If m_bln_NeedVertScroll Then vScroll.Value = 0
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Min
               hScroll.Value = hScroll.Min
            End If
            Grid1.Row = Grid1.FixedRows
            Grid1.RowSel = Grid1.FixedRows
            Grid1.Col = m_CurrCol + Grid1.FixedCols
            Grid1.colSel = Grid1.Col
            RaiseEvent ColChange(m_ntColumns(m_CurrCol))
            RaiseEvent RowChange(m_CurrRow)
            
         'Shift + Control + Home key
         ElseIf Shift = 3 Then
            m_CurrRowSel = 0
            m_CurrColSel = m_arrVisCols(LBound(m_arrVisCols))
            If m_bln_NeedVertScroll Then vScroll.Value = vScroll.Min
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Min
               hScroll.Value = hScroll.Min
            End If
            Grid1.RowSel = Grid1.FixedRows
            Grid1.colSel = Grid1.FixedCols
            RaiseEvent SelChange
         End If
         
         m_blnPaging = True
         
         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0
      
      Case vbKeyDown
                 
         If Shift = 0 Then
            If Not IsVisible(m_CurrRow) Then
               If m_CurrRow < m_RsFiltered.RecordCount - 1 Then
                  m_CurrRow = m_CurrRow + 1
               End If
               If m_CurrRowSel <> m_CurrRow Then
                  m_CurrRowSel = m_CurrRow
                  RaiseEvent SelChange
               End If
               vScroll.Value = (m_CurrRow - vScroll.LargeChange) + 1
               RaiseEvent RowChange(m_CurrRow)
            Else
               If Grid1.Row = Grid1.Rows - 1 And prevRow = Grid1.Rows - 1 Then
                  If (GetScrollValue() + 1) <= vScroll.Max Then
                     m_CurrRow = m_CurrRow + 1
                     m_CurrRowSel = m_CurrRow
                     vScroll.Value = (GetScrollValue() + 1)
                     m_blnIgnoreSel = False
                     If m_CurrRowSel <> m_CurrRow Then
                        m_CurrRowSel = m_CurrRow
                        RaiseEvent SelChange
                     End If
                     m_blnPaging = False
                     RaiseEvent RowChange(m_CurrRow)
                  End If
               End If
            End If
         'Ctrl + Down goes to bottom of column
         ElseIf Shift = 2 Then
            If m_CurrRow < m_RsFiltered.RecordCount - 1 Then
               m_CurrRow = m_RsFiltered.RecordCount - 1
               m_CurrRowSel = m_CurrRow
               vScroll.Value = vScroll.Max
               m_blnPaging = False
               RaiseEvent RowChange(m_CurrRow)
               RaiseEvent SelChange
            End If
         ElseIf Shift = 1 Then
            If Grid1.RowSel = Grid1.Rows - 1 And prevRowSel = Grid1.Rows - 1 Then
               If (GetScrollValue() + 1) <= vScroll.Max Then
                  m_CurrRowSel = m_CurrRowSel + 1
                  vScroll.Value = (GetScrollValue() + 1)
                  m_blnPaging = False
               End If
            End If
         
         ElseIf Shift = 3 Then
            If m_CurrRowSel < m_RsFiltered.RecordCount - 1 Then
               m_CurrRowSel = m_RsFiltered.RecordCount - 1
               RaiseEvent SelChange
            End If
         End If
         
         m_blnPaging = True
         
         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0
        
      Case vbKeyUp
                           
         If Shift = 0 Then
            If Not IsVisible(m_CurrRow) Then
               If m_CurrRow > Grid1.FixedRows Then
                  m_CurrRow = m_CurrRow - 1
               End If
               If m_CurrRowSel <> m_CurrRow Then
                  m_CurrRowSel = m_CurrRow
                  RaiseEvent SelChange
               End If
               vScroll.Value = m_CurrRow
            Else
               If Grid1.Row = Grid1.FixedRows And prevRow = Grid1.FixedRows Then
                  If (GetScrollValue() - 1) >= vScroll.Min Then
                     m_CurrRow = m_CurrRow - 1
                     vScroll.Value = (GetScrollValue() - 1)
                     RaiseEvent RowChange(m_CurrRow)
                  End If
                  m_blnIgnoreSel = False
                  If m_CurrRowSel <> m_CurrRow Then
                     m_CurrRowSel = m_CurrRow
                     RaiseEvent SelChange
                     m_blnPaging = False
                  End If
               End If
            End If
            
            m_iKeyCode = KeyCode
            m_iShift = Shift
            KeyCode = 0
            Shift = 0
                  
         ElseIf Shift = 1 Then
            If Grid1.RowSel = Grid1.FixedRows And prevRowSel = Grid1.FixedRows Then
               If (GetScrollValue() - 1) >= vScroll.Min Then
                  m_CurrRowSel = m_CurrRowSel - 1
                  vScroll.Value = (GetScrollValue() - 1)
                  m_blnIgnoreSel = False
                  RaiseEvent SelChange
                  m_blnPaging = True
                  m_iKeyCode = KeyCode
                  m_iShift = Shift
                  KeyCode = 0
                  Shift = 0
               End If
            End If
            
         'Ctrl + Up goes to Top of column
         ElseIf Shift = 2 Then
            If m_CurrRow > 0 Then
               m_CurrRow = 0
               m_CurrRowSel = m_CurrRow
               vScroll.Value = 0
               m_blnPaging = False
               RaiseEvent RowChange(m_CurrRow)
               RaiseEvent SelChange
            End If
            
         ElseIf Shift = 3 Then
            If m_CurrRowSel > 0 Then
               m_CurrRowSel = 0
               RaiseEvent SelChange
            End If
         End If
         
         m_blnPaging = True
         
      Case vbKeyRight
         
         If Shift = 0 Then
            If m_bln_NeedHorzScroll = True Then
               If Grid1.ColPos(Grid1.Col) + Grid1.ColWidth(Grid1.Col) > (Grid1.Width - 240) Then
                  If hScroll.Value < hScroll.Max Then
                     m_PrevHorzValue = hScroll.Value + 1
                     hScroll.Value = hScroll.Value + 1
                  End If
               End If
            End If
         ElseIf Shift = 2 Then
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Max
               hScroll.Value = hScroll.Max
            End If
            m_blnPaging = True
         End If
         
         m_iKeyCode = KeyCode
         m_iShift = Shift
         KeyCode = 0
         Shift = 0
            
      Case vbKeyLeft
                  
         If Shift = 0 Then
            If m_bln_NeedHorzScroll = True Then
               If Grid1.Col < Grid1.LeftCol Then
                  If hScroll.Value > hScroll.Min Then
                     m_blnLoading = True
                     Grid1.Redraw = False
                     hScroll.Value = hScroll.Value - 1
                     Grid1.Height = UserControl.Height
                     Grid1.ScrollBars = flexScrollBarHorizontal
                     Grid1.LeftCol = m_arrVisCols(hScroll.Value) + Grid1.FixedCols
                     Grid1.ScrollBars = flexScrollBarNone
                     Grid1.Height = UserControl.ScaleHeight
                     Grid1.Redraw = True
                     m_blnLoading = False
                     m_CurrCol = hScroll.Value
                     m_CurrColSel = m_CurrCol
                     Grid1.colSel = Grid1.Col
                     RaiseEvent SelChange
                  End If
               End If
            End If
            m_blnPaging = True
            
            m_iKeyCode = KeyCode
            m_iShift = Shift
            KeyCode = 0
            Shift = 0
         
         ElseIf (Shift And vbCtrlMask) Then
            If m_bln_NeedHorzScroll Then
               m_PrevHorzValue = hScroll.Min
               hScroll.Value = hScroll.Min
            End If
            m_blnPaging = True
            m_iKeyCode = KeyCode
            m_iShift = Shift
            KeyCode = 0
            Shift = 0
         End If
          
      Case vbKeyReturn
        
   End Select
   
   prevRow = Grid1.Row
   prevRowSel = Grid1.RowSel
   
   If m_iKeyCode <> 0 Then
      RaiseEvent KeyUp(m_iKeyCode, m_iShift)
      m_iKeyCode = 0
      m_iShift = 0
      Exit Sub
   End If
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lonPrevRow As Long
   Dim lonRows As Long
   Dim lonFromRow As Long
   Dim lonToRow As Long
   Dim blnChecked As Boolean
   Dim pCol As ntColumn
   Dim strList() As String
   Dim i As Integer
   Dim j As Long
   Dim vBkmrk As Variant
   Dim strFilter As String
   
   If m_iKeyCode <> 0 Then
      RaiseEvent KeyUp(m_iKeyCode, m_iShift)
      m_iKeyCode = 0
      m_iShift = 0
      Exit Sub
   End If
   
   RaiseEvent KeyUp(KeyCode, Shift)
   
On Error Resume Next

   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
   If Not HasRecords Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
   If m_blnLoading Then Exit Sub
   
   If Not m_blnAllowEdit Then Exit Sub

   If m_blnPaging Then
      m_blnPaging = False
      Exit Sub
   End If
   
   If KeyCode = 16 Then Exit Sub
      
   If Grid1.Col <> Grid1.colSel Then Exit Sub
   
   If m_ntColumns(Grid1.Col - Grid1.FixedCols).Enabled = False Then Exit Sub
      
   If (Grid1.Row <> Grid1.RowSel) Or (m_ntColumns(Grid1.Col - Grid1.FixedCols).ColFormat = nfgBooleanCheckBox) Then
      If (KeyCode <> 32) And (Not ValidateEditKey(KeyCode, Shift)) Then
         Exit Sub
      End If
   End If

   m_lonEditRow = m_CurrRow
   m_lonEditRowSel = m_CurrRowSel
   m_lonEditCol = Grid1.Col - Grid1.FixedCols

   If Not IsVisible(m_CurrRow) Then
      Grid1.Redraw = False
      If m_CurrRow <= m_CurrRowSel Then
         If m_CurrRow > vScroll.Max Then
            vScroll.Value = vScroll.Max
         Else
            vScroll.Value = m_CurrRow
         End If
      Else
         If m_CurrRow - (vScroll.LargeChange - 1) > 0 Then
            vScroll.Value = m_CurrRow - (vScroll.LargeChange - 1)
         Else
            vScroll.Value = 0
         End If
      End If
      Grid1.Redraw = m_blnRedraw
      Grid1.SetFocus
      DoEvents
   End If

   If Grid1.Col >= Grid1.FixedCols Then

      With Grid1

         .Redraw = False

         Set pCol = m_ntColumns(m_lonEditCol)

         
         If pCol.ColFormat = nfgBooleanCheckBox Then

            blnChecked = (.CellPicture <> picCheck.Image)

            If .Row <> .RowSel Then

               .FillStyle = flexFillRepeat
               Call SetPicture((blnChecked), True)
               .FillStyle = flexFillSingle

               Call SetRecordsetsOnEdit(m_lonEditRow, m_lonEditRowSel, blnChecked)

               If pCol.UseCriteria Then

                  m_blnIgnoreSel = True
                  m_blnLoading = True

                  If .Row < .RowSel Then
                     lonFromRow = .Row
                     lonToRow = .RowSel
                  Else
                     lonFromRow = .RowSel
                     lonToRow = .Row
                  End If

                  For lonRows = lonFromRow To lonToRow

                     .Row = lonRows

                     If m_blnColorByRow = True Then
                        .Col = .FixedCols
                        .colSel = .Cols - 1
                        .FillStyle = flexFillRepeat
                        .CellBackColor = CheckRowColCriteria((GetScrollValue() + lonRows) - .FixedRows)
                        .FillStyle = flexFillSingle
                        .Col = m_CurrCol + .FixedCols
                        If pCol.BackColor <> -1 Or pCol.ForeColor <> -1 Then
                           .CellBackColor = pCol.BackColor
                           .CellForeColor = pCol.ForeColor
                        End If
                     Else
                        .CellBackColor = CheckRowColCriteria((GetScrollValue() + lonRows) - .FixedRows, m_CurrCol)
                     End If

                  Next lonRows

                  m_blnIgnoreSel = False
                  m_blnLoading = False

                  Call CalcPaintedArea

               End If

            Else

               Call SetPicture((blnChecked), True)

               m_RsFiltered.AbsolutePosition = m_lonEditRow + 1
               m_RSMaster.Bookmark = m_RsFiltered.Bookmark

               m_RsFiltered.Fields(pCol.Name).Value = blnChecked
               m_RSMaster.Fields(pCol.Name).Value = blnChecked

               If pCol.UseCriteria Then
                  m_blnIgnoreSel = True
                  m_blnLoading = True
                  If m_blnColorByRow = True Then
                     .Col = .FixedCols
                     .colSel = .Cols - 1
                     .FillStyle = flexFillRepeat
                     .CellBackColor = CheckRowColCriteria((GetScrollValue() + .Row) - .FixedRows)
                     .FillStyle = flexFillSingle
                     .Col = pCol.index + .FixedCols
                  Else
                     .CellBackColor = CheckRowColCriteria((GetScrollValue() + .Row) - .FixedRows, pCol.index)
                  End If
                  m_blnIgnoreSel = False
                  m_blnLoading = False
               End If

            End If

         Else

            ' Build an array of all rows being edited to pass as Param in BeforeEdit event
            ReDim m_lonEditRowIDs(Abs(m_lonEditRow - m_lonEditRowSel))

            j = 0

            If m_lonEditRow > m_lonEditRowSel Then
               For i = m_lonEditRowSel To m_lonEditRow
                  m_lonEditRowIDs(j) = i
                  j = j + 1
               Next i
            Else
               For i = m_lonEditRow To m_lonEditRowSel
                  m_lonEditRowIDs(j) = i
                  j = j + 1
               Next i
            End If

            Select Case .ColAlignment(.MouseCol)
               Case 1
                  txtEdit.Alignment = 0
               Case 4
                  txtEdit.Alignment = 2
               Case 7
                  txtEdit.Alignment = 1
            End Select

            Select Case pCol.EditType
               Case nfgTextBox
                  Grid1.Highlight = flexHighlightAlways
                  Call MSFlexGridEdit(txtEdit, KeyCode)
               Case nfgComboBox
                  Grid1.Highlight = flexHighlightAlways
                  cmbEdit.Clear
                  If Len(pCol.ComboList) > 0 Then
                     strList = Split(pCol.ComboList, ",", , vbTextCompare)
                     For i = 0 To UBound(strList)
                        cmbEdit.AddItem strList(i)
                     Next i
                  End If
                  Call MSFlexGridCombo(cmbEdit, KeyCode)

            End Select

         End If
        
         .Redraw = m_blnRedraw

      End With

   End If


End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Not m_blnShown Then Exit Sub
   If IsUnbound Then Exit Sub
   If m_ntColumns.count = 0 Then Exit Sub
   If m_blnLoading Then Exit Sub
      
On Error Resume Next

   Dim bCancel As Boolean
   
   If m_blnAllowColMove Then
      If Grid1.FixedRows > 0 Then
         If Button = vbLeftButton Then
            If Grid1.Col = Grid1.colSel Then
               If Grid1.MouseCol = Grid1.Col And y < Grid1.RowHeight(0) Then
                  If m_CurrRow = 0 And m_CurrRowSel = m_RsFiltered.RecordCount - 1 Then
                     RaiseEvent ColBeginDrag(m_ntColumns(Grid1.MouseCol - Grid1.FixedCols).Name, bCancel)
                     If Not bCancel Then
                        m_lDragCol = Grid1.MouseCol
                        m_lLastLineCol = -1
                        SetDragLine x
                        Grid1.Drag vbBeginDrag
                     End If
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
   End If
     
   m_blnIgnoreSel = False

   RaiseEvent MouseDown(Button, Shift, x, y)
   
   If Not HasRecords Then Exit Sub
   
   If Button = vbLeftButton Then

      m_blnLButtonDown = True
      m_blnLeftClick = True
      m_blnResizing = False

      If Grid1.MouseRow < Grid1.FixedRows Then

         If Grid1.AllowBigSelection Then
            If Grid1.MouseCol < Grid1.FixedCols Then
               m_blnIgnoreSel = True
               m_CurrRow = 0
               m_CurrRowSel = m_RsFiltered.RecordCount - 1
               m_CurrCol = 0
               m_CurrColSel = m_RsFiltered.Fields.count - 1
               CalcPaintedArea
               m_blnIgnoreSel = True
               m_blnBigSel = True
               RaiseEvent SelChange
            Else
               m_blnIgnoreSel = True
               m_CurrRow = 0
               m_CurrRowSel = m_RsFiltered.RecordCount - 1
               m_CurrCol = Grid1.Col - Grid1.FixedCols
               m_CurrColSel = Grid1.colSel - Grid1.FixedCols
               If m_CurrColSel < 0 Then m_CurrColSel = 0
               If m_CurrCol < 0 Then m_CurrCol = 0
               CalcPaintedArea
               m_blnIgnoreSel = True
               RaiseEvent SelChange
            End If
         End If
      Else
         If Shift And vbShiftMask Then
            m_CurrRowSel = (GetScrollValue() + ((Grid1.MouseRow) - Grid1.FixedRows))
            m_CurrColSel = Grid1.colSel - Grid1.FixedCols
            If m_CurrRowSel < 0 Then m_CurrRowSel = 0
            If m_CurrColSel < 0 Then m_CurrColSel = 0
            CalcPaintedArea
         Else
            Grid1.FocusRect = m_intFocusRect
            Grid1.Highlight = m_intHighlight
            ' Set m_CurrRow to GridID where mouse went down
            m_CurrRow = (GetScrollValue() + ((Grid1.Row) - Grid1.FixedRows))
            If m_CurrRow < 0 Then m_CurrRow = 0
            m_CurrRowSel = m_CurrRow
            m_CurrCol = Grid1.Col - Grid1.FixedCols
            If m_CurrCol < 0 Then m_CurrCol = 0
            m_CurrColSel = Grid1.colSel - Grid1.FixedCols
         End If
      End If
   Else

      m_blnLButtonDown = False
      m_blnLeftClick = False

   End If
   
   Call SetCapture(Grid1.hwnd)

End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
   Picture2.Visible = False
   RaiseEvent MouseMove(Button, Shift, x, y)

On Error Resume Next
         
   Static vInt As Long
   Static hInt As Long

   Dim vInterval As Long
   Dim hInterval As Long
   Dim blnVTimer As Boolean
   Dim blnHTimer As Boolean
     
   If Button <> vbLeftButton Then
      ReleaseCapture
      Exit Sub
   End If
   
   If Not m_bln_NeedHorzScroll And Not m_bln_NeedVertScroll Then Exit Sub

' **************************
' VERTICAL

   ' If mouse is below the grid - means we are scrolling down
   If y > UserControl.Height And GetScrollValue() < vScroll.Max Then
      vInterval = ((y - Grid1.Height) - ((y - Grid1.Height) Mod 100))
      blnVTimer = True
      m_ScrollDown = True
   'If we are above the grid
   ElseIf y < 0 And GetScrollValue() > vScroll.Min Then
      vInterval = (Abs(y) - (Abs(y) Mod 100))
      blnVTimer = True
      m_ScrollDown = False
   Else
      blnVTimer = False
   End If


'*************************
' HORIZONTAL
   
   ' If the mouse is to the left of the grid
   If x < 0 And Grid1.LeftCol > Grid1.FixedCols Then
      hInterval = ((x - Grid1.Width) - ((x - Grid1.Width) Mod 200))
      blnHTimer = True
      m_ScrollLeft = True
   'if the mouse is to the right
   ElseIf x > UserControl.Width And _
      (Grid1.ColPos(Grid1.Cols - 1) + Grid1.ColWidth(Grid1.Cols - 1)) > Grid1.Width Then
      blnHTimer = True
      m_ScrollLeft = False
   Else
      blnHTimer = False
   End If

   If Not blnHTimer Then
      If m_HTimer <> 0 Then
         Call KillTimer(UserControl.hwnd, 2)
         m_HTimer = 0
      End If
   Else
      If hInterval > 1000 Then hInterval = 1000
      If hInterval < 200 Then hInterval = 200
      hInterval = (5000 \ hInterval)
      If hInterval <> hInt Then
         If m_HTimer <> 0 Then Call KillTimer(UserControl.hwnd, 2)
         m_HTimer = SetTimer(UserControl.hwnd, 2, hInterval, 0)
      End If
   End If

   hInt = hInterval

   If Not blnVTimer Then
      If m_VTimer <> 0 Then
         Call KillTimer(UserControl.hwnd, 1)
         m_VTimer = 0
      End If
   Else
      If vInterval > 4000 Then vInterval = 4000
      If vInterval < 100 Then vInterval = 100
      vInterval = (10000 \ vInterval)
      If vInterval <> vInt Then
         If m_VTimer <> 0 Then Call KillTimer(UserControl.hwnd, 1)
         m_VTimer = SetTimer(UserControl.hwnd, 1, vInterval, 0)
      End If
   End If

   vInt = vInterval

End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lMCol As Long
   Dim singleSel As Boolean
   
On Error Resume Next
      
   If m_blnDragging Then
      m_blnDragging = False
      Grid1.Drag vbEndDrag
      Picture2.Visible = False
   End If
         
   RaiseEvent MouseUp(Button, Shift, x, y)
   
   singleSel = ((m_CurrCol = m_CurrColSel) And (m_CurrRow = m_CurrRowSel))
      
   m_blnLButtonDown = False

   ' If the Recordset has not been set - do nothing
   If IsUnbound Then Exit Sub

   If Button = vbLeftButton Then

      If m_HTimer <> 0 Then
         Call KillTimer(UserControl.hwnd, 2)
         m_HTimer = 0
      End If

      If m_VTimer <> 0 Then
         Call KillTimer(UserControl.hwnd, 1)
         m_VTimer = 0
      End If

      'If m_blnIgnoreSel Or m_blnBigSel Then
         If Not m_blnBigSel Then
            m_CurrRowSel = (GetScrollValue() + ((Grid1.MouseRow) - Grid1.FixedRows))
            m_CurrColSel = Grid1.MouseCol - Grid1.FixedCols
            If m_CurrColSel < 0 Then m_CurrColSel = 0
         End If
         m_blnIgnoreSel = False
         m_blnBigSel = False
      'End If
            
       'm_CurrRowSel = (GetScrollValue() + ((Grid1.MouseRow) - Grid1.FixedRows))
       'm_CurrColSel = Grid1.MouseCol - Grid1.FixedCols
      
      Call ReleaseCapture
      
   Else

      If IsMouseOutOfGrid Then Exit Sub
      
      'Simulate a Left_Click
      If m_CurrCol = m_CurrColSel And m_CurrRow = m_CurrRowSel Then
         If Grid1.MouseRow < Grid1.FixedRows Then
            Grid1.Row = Grid1.FixedRows
            m_CurrRow = 0
         Else
            Grid1.Row = Grid1.MouseRow
            m_CurrRow = GetScrollValue() + Grid1.MouseRow - Grid1.FixedRows
         End If
         m_CurrRowSel = m_CurrRow
         m_CurrCol = Grid1.MouseCol - Grid1.FixedCols
         If m_CurrCol < 0 Then m_CurrCol = 0
         m_CurrColSel = m_CurrCol
         Grid1.RowSel = Grid1.Row
         Grid1.Col = Grid1.MouseCol
         Grid1.colSel = Grid1.Col
         Grid1.SetFocus
         If m_CurrRow <> m_LastRow Or m_CurrCol <> m_LastCol Then
            Grid1_RowColChange
            Grid1_SelChange
         End If
      End If
      
      If m_blnAllowMenu Then
       
         mnuGridFilterRemove.Enabled = IsFiltered
         mnuGridRefresh.Enabled = IsSorted
         mnuGridResetCols.Enabled = IsReordered
         mnuGridReset.Enabled = (m_blnAllowSort Or m_blnAllowFilter Or m_blnAllowColMove)
               
         mnuGridFit.Enabled = False
         
         mnuGridSortAsc.Enabled = False
         mnuGridSortDesc.Enabled = False
         
         mnuGridFilterBy.Enabled = False
         mnuGridFilterExclude.Enabled = False
         
         mnuGridEdit.Enabled = False
                     
         mnuGridFilterBy.Visible = m_blnAllowFilter And (Grid1.SelectionMode = flexSelectionFree)
         mnuGridFilterExclude.Visible = m_blnAllowFilter And (Grid1.SelectionMode = flexSelectionFree)
         mnuGridFilterRemove.Visible = m_blnAllowFilter And (Grid1.SelectionMode = flexSelectionFree)
         mnuGridBar1.Visible = m_blnAllowFilter And (Grid1.SelectionMode = flexSelectionFree)
         
         mnuGridSortAsc.Visible = m_blnAllowSort And (Grid1.SelectionMode <> flexSelectionByRow)
         mnuGridSortDesc.Visible = m_blnAllowSort And (Grid1.SelectionMode <> flexSelectionByRow)
         mnuGridRefresh.Visible = m_blnAllowSort
         mnuGridBar0.Visible = m_blnAllowSort
         
         mnuGridEdit.Visible = m_blnAllowEdit
         mnuGridBar4.Visible = m_blnAllowEdit
         
         mnuResetBar.Visible = m_blnAllowColMove
         mnuGridResetCols.Visible = m_blnAllowColMove
               
         ' If an entire column or columns are selected
         If Grid1.MouseRow < Grid1.FixedRows Then
   
            If Grid1.MouseCol >= Grid1.FixedCols Then
               
               mnuGridSortAsc.Enabled = (m_RsFiltered.RecordCount > 0) And m_blnAllowSort
               mnuGridSortDesc.Enabled = (m_RsFiltered.RecordCount > 0) And m_blnAllowSort
               
               mnuGridFit.Enabled = ((m_RsFiltered.RecordCount) > 0 And ((Grid1.AllowUserResizing = flexResizeColumns) Or (Grid1.AllowUserResizing = flexResizeBoth)))
               
               mnuGridEdit.Enabled = ((Grid1.Col = Grid1.colSel) And m_blnAllowEdit And m_ntColumns(Grid1.Col - Grid1.FixedCols).Enabled)
                     
            End If
   
         ' If an entire Row or Rows are selected
         ElseIf Grid1.MouseCol < Grid1.FixedCols Then
            'nothing
         Else
            
            mnuGridFit.Enabled = ((m_RsFiltered.RecordCount) > 0 And ((Grid1.AllowUserResizing = flexResizeColumns) Or (Grid1.AllowUserResizing = flexResizeBoth)))
            mnuGridFilterBy.Enabled = (m_RsFiltered.RecordCount > 0) And (m_CurrCol = m_CurrColSel)
            mnuGridFilterExclude.Enabled = (m_RsFiltered.RecordCount > 0) And (m_CurrCol = m_CurrColSel)
            mnuGridEdit.Enabled = ((Grid1.Col = Grid1.colSel) And m_blnAllowEdit And m_ntColumns(m_CurrCol).Enabled)
   
         End If
   
         RaiseEvent BeforeShowMenu
         PopupMenu mnuGrid
      
      End If
      
   End If
   
   prevRow = Grid1.Row
   prevRowSel = Grid1.RowSel
   
End Sub

'BEGIN MENUS ----------------------------------------------------------------------------------

Private Sub mnuCustom_Click(index As Integer)
   RaiseEvent CustomMenuClick(mnuCustom(index).Tag, mnuCustom(index).index)
End Sub

Private Sub mnuGridFilterBy_Click()
   Dim arrFilters() As String
   If Not m_blnAllowFilter Then Exit Sub
On Error Resume Next
   arrFilters = BuildGridRSFilterValues(m_CurrCol, m_CurrRow, m_CurrRowSel)
   Call AddFilter(m_ntColumns(m_CurrCol).Name, arrFilters, True)
End Sub

Private Sub mnuGridFilterExclude_Click()
   Dim arrFilters() As String
   If Not m_blnAllowFilter Then Exit Sub
On Error Resume Next
   arrFilters = BuildGridRSFilterValues(m_CurrCol, m_CurrRow, m_CurrRowSel)
   Call AddFilter(m_ntColumns(m_CurrCol).Name, arrFilters, False)
End Sub

Private Sub mnuGridFilterRemove_Click()
   Call RemoveFormattedFilter
End Sub

Public Sub RemoveFormattedFilter()
Attribute RemoveFormattedFilter.VB_HelpID = 1270

If Not m_blnShown Then Exit Sub
If IsUnbound Then Exit Sub

   Dim PrevSort As String
   
On Error GoTo F_Err
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Grid1.Redraw = False
   m_blnLoading = True
   PrevSort = m_RsFiltered.Sort
   Call GetUnFilteredRecords
   Call SetGridRows
   If Len(PrevSort) > 0 Then m_RsFiltered.Sort = PrevSort
   Call SetScrollValue(0)
   If m_bln_NeedVertScroll Then vScroll.Value = 0
   FillTextmatrix 0
   Grid1.Col = m_CurrCol + Grid1.FixedCols
   Grid1.colSel = Grid1.Col
   Grid1.Row = Grid1.FixedRows
   Grid1.RowSel = Grid1.FixedRows
   m_CurrColSel = m_CurrCol
   m_CurrRow = 0
   m_CurrRowSel = 0
   Grid1.FocusRect = m_intFocusRect
   Grid1.Highlight = m_intHighlight
   RaiseEvent OnFilterRemove
   m_blnLoading = False
   Grid1.Redraw = m_blnRedraw
   On Error Resume Next
   If m_blnHasFocus Then Grid1.SetFocus
   DoEvents
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
Exit Sub

F_Err:
   Err.Raise Err.Number, "RemoveFormattedFilter; " & Err.Source, Err.Description
End Sub

Public Property Let GridMode(ByVal eMode As nfgGridModeConstants)
Attribute GridMode.VB_HelpID = 1330
   m_blnGridMode = CBool(eMode)
   If Not IsUnbound Then Err.Raise vbObjectError + 50029, Ambient.DisplayName & ".Gridmode", "Cannot set GridMode while grid is bound. Set Recordset to nothing first."
   If Ambient.UserMode Then
      If Not m_blnGridMode Then
         Set m_colRowColors = New ntRowInfo
         Set m_ntColumns = New ntColumns
         Set m_colFilters = New ntColFilters
         Set m_colColpics = New ntColPics
      End If
   End If
   PropertyChanged "GridMode"
End Property

Public Property Get GridMode() As nfgGridModeConstants
   GridMode = Abs(m_blnGridMode)
End Property

Private Property Let GridMousePointer(ByVal New_Pointer As MousePointerConstants)
   Dim i As Integer

   Grid1.MousePointer = New_Pointer

   If lbltotal.UBound > 0 Then
      For i = 0 To lbltotal.UBound
         lbltotal(i).MousePointer = New_Pointer
      Next i
   End If

End Property

Private Sub mnuGridSortAsc_Click()
   If Not m_blnAllowSort Then Exit Sub
   Call SortGrid(, Grid1.Col - Grid1.FixedCols, Grid1.colSel - Grid1.FixedCols)
End Sub

Private Sub mnuGridSortDesc_Click()
   If Not m_blnAllowSort Then Exit Sub
   Call SortGrid(, Grid1.Col - Grid1.FixedCols, Grid1.colSel - Grid1.FixedCols, False)
End Sub

'Remove Sort
Private Sub mnuGridRefresh_Click()
   If Not m_blnAllowSort Then Exit Sub
   If Not m_blnRedraw Then Exit Sub
   Dim l As Long
On Error Resume Next
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Grid1.Redraw = False
   m_RsFiltered.Sort = ""
   For l = 0 To m_ntColumns.count - 1
      m_ntColumns(l).Sorted = nfgSortNone
   Next l
   SetScrollValue 0
   Grid1.Row = Grid1.FixedRows
   Grid1.Col = m_CurrCol + Grid1.FixedCols
   Grid1.RowSel = Grid1.FixedRows
   Grid1.colSel = m_CurrColSel + Grid1.FixedCols
   m_CurrRow = 0
   m_CurrRowSel = m_CurrRow
   Grid1.Redraw = m_blnRedraw
   If m_blnHasFocus Then Grid1.SetFocus
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
   RaiseEvent SelChange
   RaiseEvent OnSortRemove
End Sub

Private Sub mnuGridReset_Click()
On Error Resume Next
   Dim l As Long
   GridMousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   Grid1.Redraw = False
   If m_blnAllowSort Then
      m_RsFiltered.Sort = ""
      For l = 0 To m_ntColumns.count - 1
         m_ntColumns(l).Sorted = nfgSortNone
      Next l
   End If
   If m_blnAllowFilter Then
      Call GetUnFilteredRecords
   End If
   If m_blnAllowColMove And IsReordered Then
      m_ntColumns.Reset
      BuildGridFromColumns m_ntColumns
   End If
   Call SetGridRows
   SetScrollValue 0
   Grid1.Row = Grid1.FixedRows
   Grid1.Col = Grid1.FixedCols
   Grid1.RowSel = Grid1.FixedRows
   Grid1.colSel = Grid1.FixedCols
   m_CurrRow = 0
   m_CurrRowSel = 0
   m_CurrCol = 0
   m_CurrColSel = 0
   m_LastRow = 0
   m_LastCol = 0
   Grid1.FocusRect = m_intFocusRect
   Grid1.Highlight = m_intHighlight
   Grid1.Redraw = m_blnRedraw
   GridMousePointer = vbDefault
   Screen.MousePointer = vbDefault
   RaiseEvent SelChange
   RaiseEvent OnColReorder
   RaiseEvent OnFilterRemove
   RaiseEvent OnSortRemove
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
On Error Resume Next
   If PropertyName = "BackColor" Then
      UserControl.BackColor = Ambient.BackColor
      Grid1.BackColorBkg = Ambient.BackColor
      PropertyChanged ("BackColor")
   End If
End Sub

Private Sub UserControl_Click()
   If m_blnLoading Then Exit Sub
   RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
   Grid1.ScrollBars = flexScrollBarNone
   bRedrawFlag = True
   m_LastCol = -1
   m_LastRow = -1

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
   Static Resize As Boolean
   If Resize Then Exit Sub
   Resize = True
   If UserControl.Height < 1500 Then UserControl.Height = 1500
   If UserControl.Width < 1500 Then UserControl.Width = 1500
   vScroll.Resize
   hScroll.Resize
       
   Grid1.Move 0, 0, UserControl.ScaleWidth - (Abs(m_bln_NeedVertScroll) * 225), _
                  UserControl.ScaleHeight - (Abs(m_bln_NeedHorzScroll) * 255)
   Call RecalcGrid
   bRedrawFlag = True
   RaiseEvent Resize
   Resize = False
End Sub

Private Sub UserControl_Show()
On Error Resume Next
   If Ambient.UserMode And (IsUnbound Or Not HasRecords) Then
      m_bln_NeedVertScroll = False
      m_bln_NeedHorzScroll = False
      hScroll.Visible = False
      vScroll.Visible = False
   End If
End Sub

Private Sub UserControl_Terminate()
   DestroyRefs
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   On Error Resume Next
   
   Set m_picChecked = LoadResPicture(101, 1)
   Set m_picUnchecked = LoadResPicture(102, 1)
   Set m_picCheckedDis = LoadResPicture(103, 1)
   Set m_picUncheckedDis = LoadResPicture(104, 1)
   
   SetCheckBoxes

   'Defaults for FlexGrid
   With Grid1
      .AllowBigSelection = True
      .AllowUserResizing = 0
      .Appearance = flex3D
      .BackColor = &H80000005
      .BackColorBkg = Ambient.BackColor
      .BackColorFixed = &H8000000F
      .BackColorSel = &H8000000D
      Grid1.BorderStyle = flexBorderNone
      .CausesValidation = True
      .Cols = 2
      .ColWidth(0) = m_def_RecordSelWidth
      .Enabled = True
      Set .Font = Ambient.Font
      .FillStyle = flexFillSingle
      .FixedCols = 1
      .FixedRows = 1
      .FocusRect = flexFocusLight
      .ForeColor = &H80000012
      .ForeColorFixed = &H80000008
      .ForeColorSel = &H8000000E
      .FormatString = ""
      .GridColor = &HC0C0C0
      .GridColorFixed = &H0&
      .GridLines = flexGridFlat
      .GridLinesFixed = flexGridInset
      .GridLineWidth = 1
      .Highlight = flexHighlightWithFocus
      .Left = 0
      .MergeCells = flexMergeNever
      .Redraw = True
      m_blnRedraw = False
      .RightToLeft = False
      .RowHeightMin = m_def_RowHeightMin
      .Rows = 2
      .ScrollBars = flexScrollBarNone
      .ScrollTrack = True
      .SelectionMode = flexSelectionFree
      .TextStyle = 0
      .TextStyleFixed = 0
      .Visible = True
      .WordWrap = False

   End With
   
   m_eScrollBars = nfgScrollBoth
   m_sngRecordSelectorWidth = m_def_RecordSelWidth
   m_intTotalFloat = m_def_TotalFloat
   m_intFocusRect = m_def_FocusRect
   m_clrEnabledBackcolor = Grid1.BackColor
   m_blnEnabled = False
   m_Text = m_def_Text
   lbltotal(0).ForeColor = &H80000012
   lbltotal(0).BackColor = &H80000012
   Set txtEdit.Font = Ambient.Font
   Set cmbEdit.Font = Ambient.Font
   'Set dtEdit.Font = Ambient.Font
   m_blnAllowFilter = False
   m_blnAllowSort = False
   m_blnAllowEdit = False
   m_blnAutoSizeColumns = False
   m_sngMaxColWidth = m_def_MaxColWidth
   m_sngMinColWidth = m_def_MinColWidth
   m_sngRowHeight = m_def_Rowheight
   m_sngRowHeightMin = m_def_RowHeightMin
   m_sngRowHeightMax = m_def_Rowheight
   m_sngRowHeightFixed = m_def_RowHeightFixed
   m_blnUseFieldNamesAsHeader = True
   m_blnHeaderRow = True
   m_blnRecordSelectors = True
   m_blnColorByRow = False
   m_blnTotalRow = False
   m_clrGridTotalPlus = m_def_GRID_PLUSCOLOR
   m_clrGridTotalMinus = m_def_GRID_MINUSCOLOR
   m_clrDisBackcolor = m_def_DisabledColor
   m_clrPositive = m_def_GRID_PLUSCOLOR
   m_clrNegative = m_def_GRID_MINUSCOLOR
   m_intHighlight = m_def_Highlight
   m_blnDblClickFilter = False
   m_blnRedraw = False
   m_blnAllowMenu = True
   
   m_GradClrStart = &HFFFFFF
   m_GradClrEnd = &HF7D768
   m_GradType = ntFgGTHorizontal
   m_BackStyle = ntFgBsSolidColor
   m_GradTransColor = &H8000000C
   m_BackPicDrawMode = ntFgBPMNormal
   m_BackPic = Nothing
   m_FgrdDraw = ntFgFGMOpaque
   
   UserControl.BorderStyle = 1
   
   m_intEditKey = 32
   UserControl.BackColor = Ambient.BackColor

   Call SetTotalPositions
         
End Sub

Private Sub SetCheckBoxes()
   picCheck.Width = m_def_Rowheight - 30
   picCheck.Height = m_def_Rowheight - 30
   picCheck.AutoRedraw = True
   picUnCheck.Width = m_def_Rowheight - 30
   picUnCheck.Height = m_def_Rowheight - 30
   picUnCheck.AutoRedraw = True
   picCheckDis.Width = m_def_Rowheight - 30
   picCheckDis.Height = m_def_Rowheight - 30
   picCheckDis.AutoRedraw = True
   picUnCheckDis.Width = m_def_Rowheight - 30
   picUnCheckDis.Height = m_def_Rowheight - 30
   picUnCheckDis.AutoRedraw = True
   
   picCheck.PaintPicture m_picChecked, 0, 0, picCheck.ScaleWidth, picCheck.ScaleHeight
   picUnCheck.PaintPicture m_picUnchecked, 0, 0, picUnCheck.ScaleWidth, picUnCheck.ScaleHeight
   picCheckDis.PaintPicture m_picCheckedDis, 0, 0, picCheckDis.ScaleWidth, picCheckDis.ScaleHeight
   picUnCheckDis.PaintPicture m_picUncheckedDis, 0, 0, picUnCheckDis.ScaleWidth, picUnCheckDis.ScaleHeight

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   On Error Resume Next
         
   Set m_picChecked = PropBag.ReadProperty("CheckEnabledPic", LoadResPicture(101, 1))
   Set m_picUnchecked = PropBag.ReadProperty("UnCheckEnabledPic", LoadResPicture(102, 1))
   Set m_picCheckedDis = PropBag.ReadProperty("CheckDisabledPic", LoadResPicture(103, 1))
   Set m_picUncheckedDis = PropBag.ReadProperty("UnCheckDisabledPic", LoadResPicture(104, 1))
     
   'Load the picture boxes with the checkmarks(1-checked, 2-UnChecked, 3-Checked Disabled, 4-unchecked Disabled)
   SetCheckBoxes

   m_sngRecordSelectorWidth = PropBag.ReadProperty("RecordSelWidth", m_def_RecordSelWidth)

   With Grid1
      .RowHeightMin = PropBag.ReadProperty("RowHeightMin", m_def_RowHeightMin)
      .MousePointer = PropBag.ReadProperty("MousePointer", 0)
      .Rows = 2
      .FixedCols = PropBag.ReadProperty("FixedCols", 1)
      If .FixedCols > 0 Then .ColWidth(0) = m_sngRecordSelectorWidth
      .TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 3)
      .AllowBigSelection = True
      If Not Ambient.UserMode = True Then
        .AllowUserResizing = flexResizeBoth
      Else
        .AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 3)
      End If
      .BackColor = PropBag.ReadProperty("BackColor", &H80000005)
      .BackColorBkg = PropBag.ReadProperty("BackColorBkg", &H8000000F)
      .BackColorFixed = PropBag.ReadProperty("BackColorFixed", &H8000000F)
      .BackColorSel = PropBag.ReadProperty("BackColorSel", &H8000000D)
      Grid1.BorderStyle = flexBorderNone
      .CausesValidation = PropBag.ReadProperty("CausesValidation", True)
      .FillStyle = PropBag.ReadProperty("FillStyle", 0)
       Set .Font = PropBag.ReadProperty("Font", Ambient.Font)
      .ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
      .ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", &H80000008)
      .ForeColorSel = PropBag.ReadProperty("ForeColorSel", &H8000000E)
      .GridColor = PropBag.ReadProperty("GridColor", &HC0C0C0)
      .GridColorFixed = PropBag.ReadProperty("GridColorFixed", &H0&)
      .GridLines = PropBag.ReadProperty("GridLines", 1)
      .GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", 2)
      .GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
      .Redraw = True
      .ScrollBars = flexScrollBarNone
      .SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
      .ToolTipText = PropBag.ReadProperty("ToolTipText", "")
      .TextStyle = PropBag.ReadProperty("TextStyle", 0)
      .WordWrap = PropBag.ReadProperty("WordWrap", False)
      .WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
       ' ALWAYS SET THESE TO THESE SETTINGS
      .ScrollTrack = False
      .MergeCells = 0

   End With
   m_blnGridMode = PropBag.ReadProperty("GridMode", True)

   If Not m_blnGridMode Then
      Set m_colRowColors = New ntRowInfo
      Set m_ntColumns = New ntColumns
      Set m_colFilters = New ntColFilters
      Set m_colColpics = New ntColPics
   End If

   m_blnAllowColMove = PropBag.ReadProperty("AllowMoveCols", False)
   m_blnDblClickFilter = PropBag.ReadProperty("DblClickFilter", False)
   m_intEditKey = PropBag.ReadProperty("EditRangeKey", 32)
   m_intTotalFloat = PropBag.ReadProperty("TotalFloat", 0)
   m_blnEnabled = PropBag.ReadProperty("Enabled", False)
   Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   UserControl.Height = PropBag.ReadProperty("Height", 6255)
   UserControl.Width = PropBag.ReadProperty("Width", 8595)
   m_Text = PropBag.ReadProperty("Text", m_def_Text)
   m_blnAllowMenu = PropBag.ReadProperty("AllowMenu", True)
   m_blnAllowEdit = PropBag.ReadProperty("AllowEdit", False)
   m_blnAllowFilter = PropBag.ReadProperty("AllowFilter", False)
   m_blnAllowSort = PropBag.ReadProperty("AllowSort", False)
   m_blnAutoSizeColumns = PropBag.ReadProperty("AutoSizeColumns", False)
   m_sngMaxColWidth = PropBag.ReadProperty("MaxColWidth", m_def_MaxColWidth)
   m_sngMinColWidth = PropBag.ReadProperty("MinColWidth", m_def_MinColWidth)
   m_sngRowHeight = PropBag.ReadProperty("RowHeight", m_def_Rowheight)
   m_sngRowHeightMin = PropBag.ReadProperty("RowHeightMin", m_def_RowHeightMin)
   m_sngRowHeightMax = PropBag.ReadProperty("RowHeightMax", m_def_Rowheight)
   m_blnHeaderRow = PropBag.ReadProperty("FixedHeaderRow", True)
   Grid1.FixedRows = Abs(m_blnHeaderRow)
   m_sngRowHeightFixed = PropBag.ReadProperty("RowHeightFixed", m_def_RowHeightFixed)
   If m_blnHeaderRow Then
      If m_sngRowHeightFixed Then Grid1.RowHeight(0) = m_sngRowHeightFixed
   End If
   m_blnRecordSelectors = PropBag.ReadProperty("RecordSelectors", True)
   Grid1.FixedCols = Abs(m_blnRecordSelectors)
   m_intFocusRect = PropBag.ReadProperty("FocusRect", 1)
   Grid1.FocusRect = m_intFocusRect
   m_intHighlight = PropBag.ReadProperty("Highlight", flexHighlightWithFocus)
   Grid1.Highlight = m_intHighlight
   m_blnUseFieldNamesAsHeader = PropBag.ReadProperty("UseFieldNamesAsHeader", True)
   m_blnColorByRow = PropBag.ReadProperty("ColorByRow", False)
   m_blnTotalRow = PropBag.ReadProperty("TotalRow", False)
   m_clrGridTotalPlus = PropBag.ReadProperty("TotalPlusColor", m_def_GRID_PLUSCOLOR)
   m_clrGridTotalMinus = PropBag.ReadProperty("TotalMinusColor", m_def_GRID_MINUSCOLOR)
   m_clrPositive = PropBag.ReadProperty("PositiveColor", m_def_GRID_PLUSCOLOR)
   m_clrNegative = PropBag.ReadProperty("NegativeColor", m_def_GRID_MINUSCOLOR)
   Grid1.Appearance = PropBag.ReadProperty("Appearance", 1)
   m_clrEnabledBackcolor = Grid1.BackColor
   m_clrDisBackcolor = PropBag.ReadProperty("BackColorDisabled", m_def_DisabledColor)
   Set txtEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)
   Set cmbEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_blnRedraw = True
      
    m_GradClrStart = PropBag.ReadProperty("GradientStartColor", &HFFFFFF)
   m_GradClrEnd = PropBag.ReadProperty("GradientEndColor", &HF7D768)
   m_GradType = PropBag.ReadProperty("GradientType", ntFgGTHorizontal)
   m_BackStyle = PropBag.ReadProperty("BackStyle", 0)
   m_GradTransColor = PropBag.ReadProperty("BkgdMaskColor", &H8000000C)
   m_BackPicDrawMode = PropBag.ReadProperty("BkgdPictureDrawMode", ntFgBPMNormal)
   Set m_BackPic = PropBag.ReadProperty("BkgdPicture", Nothing)
   m_FgrdDraw = PropBag.ReadProperty("ForeGroundDrawStyle", ntFgFGMOpaque)

   UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   UserControl.BackColor = Ambient.BackColor
   
   m_eScrollBars = PropBag.ReadProperty("ScrollBars", nfgScrollBoth)
   
   If Not Ambient.UserMode Then
      Call SetTotalPositions
   Else
       Set m_hScroll = New ntFxGdScrollBar
       Set m_vScroll = New ntFxGdScrollBar
       m_hScroll.Bind hScroll
       m_vScroll.Bind vScroll
      'Create the scroll bars
      m_bln_NeedVertScroll = ((m_eScrollBars = nfgScrollVertical) Or (m_eScrollBars = nfgScrollBoth))
      m_bln_NeedHorzScroll = ((m_eScrollBars = nfgScrollHorizontal) Or (m_eScrollBars = nfgScrollBoth))
      vScroll.Visible = ((m_eScrollBars = nfgScrollVertical) Or (m_eScrollBars = nfgScrollBoth))
      hScroll.Visible = ((m_eScrollBars = nfgScrollHorizontal) Or (m_eScrollBars = nfgScrollBoth))
   End If
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next
    
   Call PropBag.WriteProperty("GradientStartColor", m_GradClrStart, &HFFFFFF)
   Call PropBag.WriteProperty("GradientEndColor", m_GradClrEnd, &HF7D768)
   Call PropBag.WriteProperty("GradientType", m_GradType, ntFgGTHorizontal)
   Call PropBag.WriteProperty("BackStyle", m_BackStyle, 0)
   Call PropBag.WriteProperty("BkgdMaskColor", m_GradTransColor, &H8000000C)
   Call PropBag.WriteProperty("BkgdPictureDrawMode", m_BackPicDrawMode, ntFgBPMNormal)
   Call PropBag.WriteProperty("BkgdPicture", m_BackPic, Nothing)
   Call PropBag.WriteProperty("ForeGroundDrawStyle", m_FgrdDraw, ntFgFGMOpaque)
   Call PropBag.WriteProperty("AllowMoveCols", m_blnAllowColMove, False)
   Call PropBag.WriteProperty("GridMode", m_blnGridMode, True)
   Call PropBag.WriteProperty("RecordSelWidth", m_sngRecordSelectorWidth, m_def_RecordSelWidth)
   Call PropBag.WriteProperty("EditRangeKey", m_intEditKey, 32)
   Call PropBag.WriteProperty("TotalFloat", m_intTotalFloat, 0)
   Call PropBag.WriteProperty("CheckEnabledPic", m_picChecked, LoadResPicture(101, 1))
   Call PropBag.WriteProperty("UnCheckEnabledPic", m_picUnchecked, LoadResPicture(102, 1))
   Call PropBag.WriteProperty("CheckDisabledPic", m_picCheckedDis, LoadResPicture(103, 1))
   Call PropBag.WriteProperty("UnCheckDisabledPic", m_picUncheckedDis, LoadResPicture(104, 1))
   Call PropBag.WriteProperty("Enabled", m_blnEnabled, False)
   Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 6255)
   Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
   Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
   Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
   Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 8595)
   Call PropBag.WriteProperty("Height", UserControl.Height, 6255)
   Call PropBag.WriteProperty("Width", UserControl.Width, 8595)
   Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
   Call PropBag.WriteProperty("AllowEdit", m_blnAllowEdit, False)
   Call PropBag.WriteProperty("AllowFilter", m_blnAllowFilter, False)
   Call PropBag.WriteProperty("AllowSort", m_blnAllowSort, False)
   Call PropBag.WriteProperty("AutoSizeColumns", m_blnAutoSizeColumns, False)
   Call PropBag.WriteProperty("MaxColWidth", m_sngMaxColWidth, m_def_MaxColWidth)
   Call PropBag.WriteProperty("MinColWidth", m_sngMinColWidth, m_def_MinColWidth)
   Call PropBag.WriteProperty("RowHeight", m_sngRowHeight, m_def_Rowheight)
   Call PropBag.WriteProperty("RowHeightMin", m_sngRowHeightMin, m_def_RowHeightMin)
   Call PropBag.WriteProperty("RowHeightMax", m_sngRowHeightMax, m_def_Rowheight)
   Call PropBag.WriteProperty("RowHeightFixed", m_sngRowHeightFixed, m_def_RowHeightFixed)
   Call PropBag.WriteProperty("FixedHeaderRow", m_blnHeaderRow, True)
   Call PropBag.WriteProperty("RecordSelectors", m_blnRecordSelectors, True)
   Call PropBag.WriteProperty("UseFieldNamesAsHeader", m_blnUseFieldNamesAsHeader, True)
   Call PropBag.WriteProperty("ColorByRow", m_blnColorByRow, False)
   Call PropBag.WriteProperty("TotalRow", m_blnTotalRow, False)
   Call PropBag.WriteProperty("TotalPlusColor", m_clrGridTotalPlus, m_def_GRID_PLUSCOLOR)
   Call PropBag.WriteProperty("TotalMinusColor", m_clrGridTotalMinus, m_def_GRID_MINUSCOLOR)
   Call PropBag.WriteProperty("AllowMenu", m_blnAllowMenu, True)
   Call PropBag.WriteProperty("AllowUserResizing", Grid1.AllowUserResizing, 3)
   Call PropBag.WriteProperty("Appearance", Grid1.Appearance, 1)
   Call PropBag.WriteProperty("BackColor", Grid1.BackColor, &H80000005)
   Call PropBag.WriteProperty("BackColorBkg", Grid1.BackColorBkg, 2147483663#)
   Call PropBag.WriteProperty("BackColorDisabled", m_clrDisBackcolor, m_def_DisabledColor)
   Call PropBag.WriteProperty("BackColorFixed", Grid1.BackColorFixed, 2147483663#)
   Call PropBag.WriteProperty("BackColorSel", Grid1.BackColorSel, 2147483661#)
   Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
   Call PropBag.WriteProperty("CausesValidation", Grid1.CausesValidation, True)
   Call PropBag.WriteProperty("FillStyle", Grid1.FillStyle, 0)
   Call PropBag.WriteProperty("FixedCols", Grid1.FixedCols, 1)
   Call PropBag.WriteProperty("FocusRect", m_intFocusRect, 1)
   Call PropBag.WriteProperty("Font", Grid1.Font, Ambient.Font)
   Call PropBag.WriteProperty("ForeColor", Grid1.ForeColor, &H80000008)
   Call PropBag.WriteProperty("ForeColorFixed", Grid1.ForeColorFixed, 2147483656#)
   Call PropBag.WriteProperty("ForeColorSel", Grid1.ForeColorSel, 2147483662#)
   Call PropBag.WriteProperty("GridColor", Grid1.GridColor, 12632256)
   Call PropBag.WriteProperty("GridColorFixed", Grid1.GridColorFixed, 0)
   Call PropBag.WriteProperty("GridLines", Grid1.GridLines, 1)
   Call PropBag.WriteProperty("GridLinesFixed", Grid1.GridLinesFixed, 2)
   Call PropBag.WriteProperty("GridLineWidth", Grid1.GridLineWidth, 1)
   Call PropBag.WriteProperty("HighLight", m_intHighlight, flexHighlightWithFocus)
   Call PropBag.WriteProperty("MergeCells", Grid1.MergeCells, 0)
   Call PropBag.WriteProperty("MousePointer", Grid1.MousePointer, 0)
   Call PropBag.WriteProperty("SelectionMode", Grid1.SelectionMode, 0)
   Call PropBag.WriteProperty("ToolTipText", Grid1.ToolTipText, "")
   Call PropBag.WriteProperty("TextStyle", Grid1.TextStyle, 0)
   Call PropBag.WriteProperty("TextStyleFixed", Grid1.TextStyleFixed, 3)
   Call PropBag.WriteProperty("WordWrap", Grid1.WordWrap, False)
   Call PropBag.WriteProperty("WhatsThisHelpID", Grid1.WhatsThisHelpID, 0)
   Call PropBag.WriteProperty("PositiveColor", m_clrPositive, m_def_GRID_PLUSCOLOR)
   Call PropBag.WriteProperty("NegativeColor", m_clrNegative, m_def_GRID_MINUSCOLOR)
   Call PropBag.WriteProperty("DblClickFilter", m_blnDblClickFilter, False)
   Call PropBag.WriteProperty("ScrollBars", m_eScrollBars, nfgScrollBoth)


End Sub

Private Function GetScrollValue() As Long
    
On Error Resume Next
   
   GetScrollValue = 0
   
   If m_bln_NeedVertScroll Then
      GetScrollValue = vScroll.Value
      If GetScrollValue > vScroll.Max Then
         GetScrollValue = vScroll.Max
      End If
   End If

End Function


























