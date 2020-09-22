VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2170
Attribute VB_Description = "ntFlexGridControl.ntColumns"
Option Explicit
Option Compare Text

Public Enum ESBCScrollTypes
   esbcHorizontal
   esbcVertical
   esbcSizeGripper
End Enum

Public Enum ESBCButtonPositionConstants
   esbcButtonPositionDefault
   esbcButtonPositionLeftTop
   esbcButtonPositionRightBottom
End Enum

Public Enum nfgValidationType
  nfgValidateNone = 0
  nfgValidateAlpha = 1
  nfgValidateNumeric = 2
  nfgValidateAlphaNumeric = 3
End Enum

Public Enum nfgRowColor
  nfgBlue = &HFFFFC0
  nfgGray = &HE0E0E0
  nfgGreen = &HC0FFC0
  nfgPeach = &HC0E0FF
  nfgYellow = &HC0FFFF  'default
End Enum

Public Enum nfgColFormat
  nfgGeneral = 0
  nfgNumeric = 1
  nfgCurrency = 2
  nfgPercentage = 3
  nfgBooleanTrueFalse = 4
  nfgBooleanCheckBox = 5
  nfg24hourTime = 6
  nfgShortDay = 7
  nfgLongDay = 8
  nfgShortDayTime = 9
  nfgLongDayTime = 10
  nfgShortDate = 11
  nfgShortDateTime = 12
  nfgMedDate = 13
  nfgMedDateTime = 14
  nfgLongDate = 15
  nfgLongDateTime = 16
  nfgMedDayDate = 17
  nfgMedDayDateTime = 18
  nfgLongDayDate = 19
  nfgLongDayDateTime = 20
  nfgCustomFormat = 21
  nfgPicture = 22
  nfgPictureText = 23
  nfgPictureCustFormatText = 24
End Enum

Public Enum nfgTotalPosition
   nfgBottom = 0
   nfgFloat = 1
End Enum

Public Enum nfgColSort
  nfgSortNone = 0
  nfgSortAsc = 1
  nfgSortDesc = 2
End Enum

Public Enum nfgTextStyleSettings
  nfgTextFlat = 0
  nfgTextRaised = 1
  nfgTextInset = 2
  nfgTextRaisedLight = 3
  nfgTextInsetLight = 4
End Enum

Public Enum nfgColEditType
  nfgTextBox = 0
  nfgComboBox = 1
End Enum

Public Enum nfgGridAlignment
  nfgAlignLeftTop = 0
  nfgAlignLeft = 1
  nfgAlignLeftBottom = 2
  nfgAlignCenterTop = 3
  nfgAlignCenter = 4
  nfgAlignCenterBottom = 5
  nfgAlignRightTop = 6
  nfgAlignRight = 7
  nfgAlignRightBottom = 8
End Enum

Public Enum nfgAllowUserResizeSettings
  nfgResizeNone = 0
  nfgResizeColumns = 1
  nfgResizeRows = 2
  nfgResizeBoth = 3
End Enum

Public Enum nfgAppearanceSettings
  nfgFlat = 0
  nfg3D = 1
End Enum

Public Enum nfgBorderStyleSettings
  nfgNoBorder = 0
  nfgSingle = 1
End Enum

Public Enum nfgFocusRectSettings
  nfgNone = 0
  nfgLight = 1
  nfgHeavy = 2
End Enum

Public Enum nfgGridLineSettings
  nfgGridNone = 0
  nfgGridFlat = 1
  nfgGridInset = 2
  nfgGridRaised = 3
End Enum

Public Enum nfgHighlightSettings
  nfgHighlightNever = 0
  nfgHighlightAlways = 1
  nfgHighlightWithFocus = 2
End Enum

Public Enum nfgScrollBarSettings
  nfgScrollNone = 0
  nfgScrollHorizontal = 1
  nfgScrollVertical = 2
  nfgScrollBoth = 3
End Enum

Public Enum nfgSelectionModeSettings
  nfgSelectionFree = 0
  nfgSelectionByRow = 1
  nfgSelectionByColumn = 2
End Enum

Private m_strName As String
Private m_arrIndexes() As String

Private m_colColumns As New Collection

Public Property Let Name(ByVal strName As String)
   m_strName = strName
End Property

Public Property Get Name() As String
Attribute Name.VB_HelpID = 2930
   Name = m_strName
End Property

' Creates a new column, inserts into collection, and returns the reference
Public Function NewColumn() As ntColumn
Attribute NewColumn.VB_HelpID = 2850
                  
    Set NewColumn = New ntColumn
        
End Function

' Creates a new column, inserts into collection, and returns the reference
Public Sub Insert(ByRef objCol As ntColumn, ColName As String)
Attribute Insert.VB_HelpID = 2860
           
    objCol.index = m_colColumns.count
    If objCol.ColID = -1 Then objCol.ColID = objCol.index
    
    m_colColumns.Add objCol, ColName
           
    SetIndexOrder
    
End Sub

Friend Sub SetIndexOrder()
   Dim l As Long
   Dim arrIndex() As Long
   Dim cc As ntColumn
   
   'First get a list of all indexes(in case some aren't there yet)
   ReDim arrIndex(m_colColumns.count - 1)
   ReDim m_arrIndexes(m_colColumns.count - 1)
   
   For l = 1 To m_colColumns.count
      arrIndex(l - 1) = m_colColumns.Item(l).index
   Next l
   
   'Now Sort the indexes
   Call BubbleSortLongArray(arrIndex, False)
         
   'Now add the sorted index col names to array
   For l = 0 To UBound(arrIndex)
      For Each cc In m_colColumns
         If cc.index = arrIndex(l) Then
            m_arrIndexes(l) = cc.Name
            Exit For
         End If
      Next cc
   Next l
   
End Sub

Public Sub Clear()
Attribute Clear.VB_HelpID = 2870
    
    Dim lonTotal As Long
    Dim lonCounter As Long
    
    lonTotal = m_colColumns.count
    lonCounter = 1
    
    Do While lonCounter <= lonTotal
      m_colColumns.Remove 1
      lonCounter = lonCounter + 1
    Loop
    
End Sub

Public Sub Remove(ByVal varKey As Variant)
Attribute Remove.VB_HelpID = 2880
  
  If IsNumeric(varKey) Then varKey = m_arrIndexes(varKey)
  
  m_colColumns.Remove varKey

End Sub

Public Function Exists(ByVal varKey As Variant) As Boolean
Attribute Exists.VB_HelpID = 2890
    
    On Error GoTo Exists_Error
    
    If IsNumeric(varKey) Then varKey = m_arrIndexes(varKey)
    
    If Not (m_colColumns.Item(varKey) Is Nothing) Then Exists = True
    
    Exit Function
    
Exists_Error:
    Exists = False
End Function

Public Property Get Column(ByVal varKey As Variant) As ntColumn
Attribute Column.VB_HelpID = 2940
Attribute Column.VB_UserMemId = 0
Attribute Column.VB_MemberFlags = "200"
  Dim pCol As ntColumn
  
On Error GoTo Item_Error

   If IsNumeric(varKey) Then varKey = m_arrIndexes(varKey)
   Set Column = m_colColumns.Item(varKey)
      
Exit Property

Item_Error:
    Err.Raise 10009, "ntGrid.Columns", "No such Column Exists"
End Property

' Sub to Resort all Columns and Reset Indexes back to start
Public Sub Reset()
Attribute Reset.VB_HelpID = 2900
   Dim pCol As ntColumn
   Dim l As Long
   Dim m_colTemp As New Collection
  
   For l = 0 To m_colColumns.count - 1
      For Each pCol In m_colColumns
         If pCol.ColID = l Then
            pCol.index = l
            m_colTemp.Add pCol, pCol.Name
         End If
      Next pCol
   Next l
    
   Set m_colColumns = m_colTemp
   Set m_colTemp = Nothing
     
   SetIndexOrder
   
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2180
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = m_colColumns.[_NewEnum]

End Property

Public Property Get count() As Long
Attribute count.VB_HelpID = 2960
    
    count = m_colColumns.count

End Property

Public Property Get HiddenColumns() As Long
Attribute HiddenColumns.VB_HelpID = 2970
  Dim i As Long
  Dim Cols As Long
  Dim pCol As ntColumn
  
  If m_colColumns.count = 0 Then
    HiddenColumns = 0
    Exit Property
  End If
  
  For Each pCol In m_colColumns
    If pCol.Visible = False Then Cols = Cols + 1
  Next pCol
  
  HiddenColumns = Cols
    
End Property

Public Function Clone() As ntColumns
Attribute Clone.VB_HelpID = 2910
   
   Dim pCols As New ntColumns
   
   pCols.Properties = Properties
   
   Set Clone = pCols
   
   Set pCols = Nothing
   
End Function

Public Property Get Properties() As Byte()
Attribute Properties.VB_HelpID = 2980
  
   Dim pCol As ntColumn
   Dim PropBag As New PropertyBag
   Dim l As Long
      
   PropBag.WriteProperty "Name", m_strName, ""
     
   For l = 1 To m_colColumns.count
      PropBag.WriteProperty "COL" & CStr(l), Column(l - 1).AllProperties, Nothing
   Next l

   Properties = PropBag.Contents

End Property

Public Property Let Properties(mCols() As Byte)
   
   Dim pCol As ntColumn
   Dim PropBag As New PropertyBag
   Dim l As Long
   
   If m_colColumns.count > 0 Then
      Err.Raise vbObjectError + 1019, "ntFlexGridBeta.Columns", _
         "This function can only be performed on a new empty layout."
      Exit Property
   End If
   
   PropBag.Contents = mCols
     
   m_strName = PropBag.ReadProperty("Name", "")
      
    l = 1
    
   Do While Not Len(PropBag.ReadProperty("COL" & CStr(l), "")) = 0
      Set pCol = New ntColumn
      pCol.AllProperties = PropBag.ReadProperty("COL" & CStr(l))
      m_colColumns.Add pCol, pCol.Name
      l = l + 1
   Loop
   
   If m_colColumns.count > 0 Then SetIndexOrder
   
End Property

Private Sub Class_Initialize()
   m_strName = ""
End Sub

Public Sub BubbleSortLongArray(arr() As Long, Optional ByVal descending As Boolean = False)

   Dim Value As Single
   Dim index As Long
   Dim firstItem As Long
   Dim indexLimit As Long, lastSwap As Long
   Dim inverseOrder As Boolean
   Dim numEls As Long
      
    ' account for optional arguments
    ' (delete if working with VB3)
    numEls = UBound(arr)
    inverseOrder = (descending <> False)

    firstItem = LBound(arr)

    lastSwap = numEls
    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For index = firstItem To indexLimit
            Value = arr(index)
            If (Value > arr(index + 1)) Xor inverseOrder Then
                ' if the items are not in order, swap them
                arr(index) = arr(index + 1)
                arr(index + 1) = Value
                lastSwap = index
            End If
        Next
    Loop While lastSwap
End Sub

