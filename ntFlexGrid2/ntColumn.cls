VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 100
Attribute VB_Description = "ntFlexGridControl.ntColumn"
Option Explicit
Option Compare Text

Private Const m_fmt_Numeric   As String = "#,##0"
Private Const m_fmt_Currency  As String = "$#,##0"
Private Const m_fmt_Percentage As String = "Percent"
Private Const m_fmt_blnTrueFalse As String = "True/False"
Private Const m_fmt_24hourTime = "HH:mm"
Private Const m_fmt_ShortDay = "ddd"
Private Const m_fmt_ShortDayTime = "ddd HH:mm"
Private Const m_fmt_LongDay = "dddd"
Private Const m_fmt_LongDayTime = "dddd HH:mm"
Private Const m_fmt_ShortDate = "MM/dd/yyyy"
Private Const m_fmt_ShortDateTime = "MM/dd/yyyy HH:mm"
Private Const m_fmt_MedDate = "MMM dd, yyyy"
Private Const m_fmt_MedDateTime = "MMM dd, yyyy HH:mm"
Private Const m_fmt_LongDate = "MMMM dd, yyyy"
Private Const m_fmt_LongDateTime = "MMMM dd, yyyy HH:mm"
Private Const m_fmt_MedDayDate = "ddd MMM dd, yyyy"
Private Const m_fmt_MedDayDateTime = "ddd MMM dd, yyyy HH:mm"
Private Const m_fmt_LongDayDate = "dddd MMMM dd, yyyy"
Private Const m_fmt_LongDayDateTime = "dddd MMMM dd, yyyy HH:mm"

Private Const m_def_Numeric As String = m_fmt_Numeric
Private Const m_def_Boolean As String = m_fmt_blnTrueFalse
Private Const m_def_DateTime As String = m_fmt_ShortDate
Private Const m_def_ColFormat As Integer = 0
Private Const m_def_ColAlignment As Integer = 4
Private Const m_def_enabled As Boolean = False
Private Const m_def_UseCriteria As Boolean = False
Private Const m_def_Sorted As Integer = 0
Private Const m_def_ColEditType As Integer = 0
Private Const m_def_Visible As Boolean = True
Private Const m_def_RowCriteriaColor As Long = &HC0FFFF
Private Const m_def_UseCustom As Boolean = False
Private Const m_def_ShowTotal As Boolean = False
Private Const m_def_Validate As Integer = 0

Private m_strFormat     As String                ' Default is "", when ColFormat is set, default is
Private m_strCustFormat As String
Private m_blnUseCustom  As Boolean
Private m_ColFormat     As nfgColFormat          ' Default is nfgGeneral = 0
Private m_intColAlignment As nfgGridAlignment
Private m_intColAlignmentHdr As nfgGridAlignment
Private m_lonColID      As Long                  ' Keep track of original Position in Grid
Private m_CurrIndex     As Integer               ' Track Current Position in Grid
Private m_intDecimal    As Integer               ' Default = 2 Decimal Places for numeric Formats
Private m_strColName    As String                ' Default = ""
Private m_strHeaderText As String                ' Text to display in header
Private m_intSorted     As nfgColSort            ' Default = 0
Private m_blnEnabled    As Boolean               ' Default = False
Private m_intEditType   As nfgColEditType        ' Default = TextBox
Private m_strText       As String                ' Text to User for Edit Controls
Private m_Datatype      As Long                  ' Default = adVarWChar
Private m_blnVisible    As Boolean               ' Is column currently Visible
Private m_sngColWidth   As Single                ' Base width of column
Private m_blnUseCriteria As Boolean
Private m_blnShowTotal  As Boolean
Private m_arrRowCriteria() As String
Private m_lonRowColor   As nfgRowColor
Private m_lonForecolor As OLE_COLOR
Private m_lonBackcolor As OLE_COLOR
Private m_intValidate   As nfgValidationType
Private m_blnColorNumbers As Boolean
Private m_ComboList As String

Public Property Get ColFormat() As nfgColFormat
Attribute ColFormat.VB_HelpID = 3010
  ColFormat = m_ColFormat
End Property

'**
'Returns or sets the current Format property setting for the Column.
'@param        fgFormat Integer. Required. A valid member of the nfgColFormat enumeration.
Public Property Let ColFormat(ByVal fgFormat As nfgColFormat)
  Dim i As Integer
  Dim strZeros As String
      
  If fgFormat < 0 Or fgFormat > 24 Then
    Err.Raise 10001, "Column " & Me.Name, "Invalid Property Value for Column Format"
    Exit Property
  End If
        
    m_ColFormat = fgFormat
    
    If m_intDecimal > 0 Then
      For i = 1 To m_intDecimal
        strZeros = strZeros & "0"
      Next i
      strZeros = "." & strZeros
    End If
    
    Select Case m_ColFormat
      Case nfgCustomFormat
      Case nfgGeneral
      Case nfgNumeric
        m_strFormat = m_fmt_Numeric & strZeros
      Case nfgCurrency
        m_strFormat = m_fmt_Currency & strZeros
      Case nfgPercentage
        m_strFormat = m_fmt_Percentage
      Case nfgBooleanTrueFalse
        m_strFormat = m_fmt_blnTrueFalse
        m_intEditType = nfgComboBox
        m_strText = "True,False"
      Case nfgBooleanCheckBox
        m_strFormat = m_fmt_blnTrueFalse
        m_intEditType = -1
      Case nfg24hourTime
        m_strFormat = m_fmt_24hourTime
      Case nfgShortDay
        m_strFormat = m_fmt_ShortDay
      Case nfgLongDay
        m_strFormat = m_fmt_LongDay
      Case nfgShortDayTime
        m_strFormat = m_fmt_ShortDayTime
      Case nfgLongDayTime
        m_strFormat = m_fmt_LongDayTime
      Case nfgShortDate
        m_strFormat = m_fmt_ShortDate
      Case nfgShortDateTime
        m_strFormat = m_fmt_ShortDateTime
      Case nfgMedDate
        m_strFormat = m_fmt_MedDate
      Case nfgMedDateTime
        m_strFormat = m_fmt_MedDateTime
      Case nfgLongDate
        m_strFormat = m_fmt_LongDate
      Case nfgLongDateTime
        m_strFormat = m_fmt_LongDateTime
      Case nfgMedDayDate = 13
        m_strFormat = m_fmt_MedDayDate
      Case nfgMedDayDateTime
        m_strFormat = m_fmt_MedDayDateTime
      Case nfgLongDayDate = 13
        m_strFormat = m_fmt_LongDayDate
      Case nfgLongDayDateTime
        m_strFormat = m_fmt_LongDayDateTime
    End Select

End Property

Public Property Get ColID() As Long
Attribute ColID.VB_HelpID = 3020
  ColID = m_lonColID
End Property

Friend Property Let ColID(ByVal lID As Long)
  m_lonColID = lID
End Property

Public Property Get ComboList() As String
Attribute ComboList.VB_HelpID = 3030
  ComboList = m_ComboList
End Property

Public Property Let ComboList(ByVal sList As String)
  m_ComboList = sList
End Property

Public Property Get CustomFormatString() As String
Attribute CustomFormatString.VB_HelpID = 3040
  CustomFormatString = m_strCustFormat
End Property

'**
'Returns or sets a custom format string for the column.
'@param        sFormat String. Required. A string value to be used with the internal Format function.
'@rem Note: This will only have effect if the UseCustomFormat property of the column is set to True.
Public Property Let CustomFormatString(ByVal sFormat As String)
  m_strCustFormat = sFormat
End Property

'**
'Returns the ADO Datatype of the column based on the recordset.
'@return        ADODB.DatatypeEnum.
Public Property Get DataType() As Long
Attribute DataType.VB_HelpID = 3050
  DataType = m_Datatype
End Property

Friend Property Let DataType(ByVal pType As Long)
  m_Datatype = pType
End Property

Public Property Get DecimalPlaces() As Integer
Attribute DecimalPlaces.VB_HelpID = 3060
  DecimalPlaces = m_intDecimal
End Property

'**
'Returns or sets the number of decimal places to display in a
'formatted numeric type in the Grid.
'@param        iDecimals Integer. Required. An integer value from 0 to 20.
Public Property Let DecimalPlaces(ByVal iDecimals As Integer)
  Dim strZeros As String
  Dim i As Integer
  If iDecimals >= 0 And iDecimals <= 20 Then
    m_intDecimal = iDecimals
    If m_intDecimal > 0 Then
      For i = 1 To m_intDecimal
        strZeros = strZeros & "0"
      Next i
      strZeros = "." & strZeros
    End If
    Select Case m_ColFormat
      Case nfgCurrency
        m_strFormat = m_fmt_Currency & strZeros
      Case nfgNumeric
        m_strFormat = m_fmt_Numeric & strZeros
    End Select
  Else
    Err.Raise 10001, "Column " & Me.Name, "Invalid Property Value for Decimal Places"
  End If
  
End Property

Public Property Get EditType() As nfgColEditType
Attribute EditType.VB_HelpID = 3070
  EditType = m_intEditType
End Property

'**
'Returns or sets the type of control displayed in the Grid
'when the column is enabled and the user clicks it.
'$EOL$
'Default is a textbox.
'@param        nType Integer. A valid member of the nfgColEditType enumeration.
Public Property Let EditType(ByVal nType As nfgColEditType)
  If m_ColFormat = nfgBooleanCheckBox Then
    m_intEditType = -1
    Exit Property
  End If
  If nType >= 0 And nType <= 3 Then
    m_intEditType = nType
  Else
    Err.Raise 10001, "Column " & Me.Name, "Invalid Property Value for Column Edit Type"
  End If
End Property

'**
'Returns or sets whether the column can be edited.
Public Property Let Enabled(ByVal pEnabled As Boolean)
Attribute Enabled.VB_HelpID = 3080
  m_blnEnabled = pEnabled
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_blnEnabled
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_HelpID = 3090
   ForeColor = m_lonForecolor
End Property

Public Property Let ForeColor(ByVal pColor As OLE_COLOR)
   m_lonForecolor = pColor
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_HelpID = 3100
   BackColor = m_lonBackcolor
End Property

Public Property Let BackColor(ByVal pColor As OLE_COLOR)
   m_lonBackcolor = pColor
End Property

'**
'This method is used by the Grid to format the recordset values,
'using the FormatString or CustomFormatString specified in the column properties.
'@param        vValue Variant. A value to format.
'@return       A formatted string value.
Public Function FormatValue(ByVal vValue As Variant) As String
Attribute FormatValue.VB_HelpID = 3000
   
   If IsNull(vValue) Then
      vValue = vbNullString
   Else
      vValue = CStr(vValue)
   End If
   
   Select Case m_ColFormat
      Case nfgGeneral, nfgPictureText
         FormatValue = CStr(vValue)
      Case nfgCustomFormat, nfgPictureCustFormatText
         If Len(m_strCustFormat) = 0 Then
            FormatValue = CStr(vValue)
         Else
            FormatValue = Format$(vValue, m_strCustFormat)
         End If
      Case Else
         If Len(m_strFormat) = 0 Then
            FormatValue = CStr(vValue)
         Else
            FormatValue = Format$(vValue, m_strFormat)
         End If
   End Select
   
End Function

'**
'Returns the actual string used by the format function in the column.
Public Property Get FormatString() As String
Attribute FormatString.VB_HelpID = 3110
  FormatString = m_strFormat
End Property

Friend Property Get CurrentFormatString() As String
  CurrentFormatString = ""
  Select Case m_ColFormat
    Case nfgGeneral
    Case nfgCustomFormat, nfgPictureCustFormatText
      If Len(m_strCustFormat) <> 0 Then CurrentFormatString = m_strCustFormat
    Case Else
      If Len(m_strFormat) <> 0 Then CurrentFormatString = m_strFormat
  End Select
End Property

Public Property Get HeaderText() As String
Attribute HeaderText.VB_HelpID = 3120
  HeaderText = m_strHeaderText
End Property

Public Property Let HeaderText(ByVal sText As String)
  m_strHeaderText = sText
End Property

'**
'Returns the current index in the Grid for the column. If the column has been moved,
' this will not match the ColID property, which references the columns original position.
Public Property Get index() As Integer
Attribute index.VB_HelpID = 3130
  index = m_CurrIndex
End Property

Public Property Let index(ByVal iIndex As Integer)
  m_CurrIndex = iIndex
End Property

Public Property Get Name() As String
Attribute Name.VB_HelpID = 3140
  Name = m_strColName
End Property

'**
'Returns or sets the name of the column.
'$EOL$
'$EOL$
'@param        sName String. Required.
'@rem Note: If setting the names programatically,
'they must match the field names in the recordset, or an error will be generated.
Public Property Let Name(ByVal sName As String)
  m_strColName = sName
  If Len(m_strHeaderText) = 0 Then m_strHeaderText = sName
End Property

Public Property Get RowCriteria() As Variant
Attribute RowCriteria.VB_HelpID = 3150
  RowCriteria = Join(m_arrRowCriteria, ",")
End Property

'**
'Returns or sets criteria to look for in the column. If criteria matching is enabled,
'the Grid will color a cell or row if the criteria is found, depending on the Grids
'ColorByRow property setting.
'@param        vValue Variant. Required. The value to try and match in the recordset.
Public Property Let RowCriteria(ByVal vValue As Variant)
  Dim arrtxt() As String
  Dim i As Integer
  If InStr(1, vValue, ",", vbTextCompare) <> 0 Then
    arrtxt = Split(CStr(vValue), ",", , vbTextCompare)
    ReDim m_arrRowCriteria(UBound(arrtxt))
    For i = 0 To UBound(arrtxt)
      m_arrRowCriteria(i) = CStr(arrtxt(i))
    Next i
  Else
    ReDim m_arrRowCriteria(0)
    m_arrRowCriteria(0) = CStr(vValue & "")
  End If
End Property

Friend Property Get RowCriteriaMatch(ByVal pValue As Variant) As Boolean
  Dim i As Integer
  
  If IsNull(pValue) Then pValue = vbNullString
  
  RowCriteriaMatch = False
  
  For i = 0 To UBound(m_arrRowCriteria)
    If StrComp(m_arrRowCriteria(i), CStr(pValue), vbTextCompare) = 0 Then
      RowCriteriaMatch = True
      Exit Property
    End If
  Next i
  
End Property

Public Property Get RowCriteriaColor() As nfgRowColor
Attribute RowCriteriaColor.VB_HelpID = 3160
  RowCriteriaColor = m_lonRowColor
End Property

'**
'Returns or sets the color used to display a criteria match in the Grid
'@param        cColor Long. A valid member of the nfgRowColor enumeration.
Public Property Let RowCriteriaColor(ByVal lColor As nfgRowColor)
  m_lonRowColor = lColor
End Property

Friend Property Let Sorted(ByVal pSort As nfgColSort)
  If pSort >= 0 And pSort <= 2 Then
    m_intSorted = pSort
  Else
    Err.Raise 10001, "Column " & Me.Name, "Invalid Property Value for Column Sort"
  End If
End Property
Friend Property Get Sorted() As nfgColSort
  Sorted = m_intSorted
End Property

Public Property Get UseCriteria() As Boolean
Attribute UseCriteria.VB_HelpID = 3170
  UseCriteria = m_blnUseCriteria
End Property

'**
'Returns or sets whether the column will look to match criteria.
'@param        bValue Boolean. Required.
'@rem Note: Be careful if setting this, as the default Criteria is an empty string (""),
'and it would be easy to get matches you didn't realize you were lookin for.
Public Property Let UseCriteria(ByVal bValue As Boolean)
  m_blnUseCriteria = bValue
End Property

Public Property Get UseColoredNumbers() As Boolean
Attribute UseColoredNumbers.VB_HelpID = 3180
  UseColoredNumbers = m_blnColorNumbers
End Property

Public Property Let UseColoredNumbers(ByVal bValue As Boolean)
  m_blnColorNumbers = bValue
End Property

Public Property Get ShowTotal() As Boolean
Attribute ShowTotal.VB_HelpID = 3190
  ShowTotal = m_blnShowTotal
End Property

'**
'Returns or sets whether the column will display a running total for the data in the Grid.
'@param        bValue Boolean. Required.
Public Property Let ShowTotal(ByVal bValue As Boolean)
  m_blnShowTotal = bValue
End Property

Public Property Get TextAlignment() As nfgGridAlignment
Attribute TextAlignment.VB_HelpID = 3200
  TextAlignment = m_intColAlignment
End Property

'**
' Returns or sets the alighment of the text in the column.
'@param        iAlignment Integer. Required. A valid member of the nfgGridAlignment enumeration.
Public Property Let TextAlignment(ByVal iAlignment As nfgGridAlignment)
  m_intColAlignment = iAlignment
End Property

Public Property Get TextAlignmentHeader() As nfgGridAlignment
Attribute TextAlignmentHeader.VB_HelpID = 3210
  TextAlignmentHeader = m_intColAlignmentHdr
End Property

'**
' Returns or sets the alighment of the text in the column.
'@param        iAlignment Integer. Required. A valid member of the nfgGridAlignment enumeration.
Public Property Let TextAlignmentHeader(ByVal iAlignment As nfgGridAlignment)
  m_intColAlignmentHdr = iAlignment
End Property

Public Property Get ValidationType() As nfgValidationType
Attribute ValidationType.VB_HelpID = 3220
  ValidationType = m_intValidate
End Property

'**
'Returns or sets the type of validation used for user input when editing a column.
'@param        iType Integer. A valid member of the nfgValidationType enumeration.
Public Property Let ValidationType(ByVal iType As nfgValidationType)
  m_intValidate = iType
End Property

'**
'Returns a value indicating if the column is visible.
Public Property Get Visible() As Boolean
Attribute Visible.VB_HelpID = 3230
  Visible = m_blnVisible
End Property

Public Property Let Visible(ByVal pVisible As Boolean)
  m_blnVisible = pVisible
End Property

Public Property Get Width() As Single
Attribute Width.VB_HelpID = 3240
  Width = m_sngColWidth
End Property

'**
'Returns or sets the width of the column, in twips. If the settings is outside of the range
'specified by the Grids ColWidthMin and ColWidthMax properties, it will have no effect.
'@param        sWidth Single. Required.
Public Property Let Width(ByVal sWidth As Single)
  m_sngColWidth = sWidth
End Property

Friend Property Get AllProperties() As Byte()
   
   Dim PropBag As New PropertyBag
         
   PropBag.WriteProperty "HeaderText", m_strHeaderText, ""
   PropBag.WriteProperty "TextAlign", m_intColAlignment, m_def_ColAlignment
   PropBag.WriteProperty "ColFormat", m_ColFormat, m_def_ColFormat
   PropBag.WriteProperty "FormatString", m_strFormat, ""
   PropBag.WriteProperty "CustomFormatString", m_strCustFormat, ""
   PropBag.WriteProperty "Visible", m_blnVisible, m_def_Visible
   PropBag.WriteProperty "Enabled", m_blnEnabled, m_def_enabled
   PropBag.WriteProperty "EditType", m_intEditType, m_def_ColEditType
   PropBag.WriteProperty "Decimals", m_intDecimal, 2
   PropBag.WriteProperty "ValidationType", m_intValidate, m_def_Validate
   PropBag.WriteProperty "ShowTotal", m_blnShowTotal, m_def_ShowTotal
   PropBag.WriteProperty "UseCriteria", m_blnUseCriteria, m_def_UseCriteria
   PropBag.WriteProperty "RowCriteria", RowCriteria, ""
   PropBag.WriteProperty "RowCriteriaColor", m_lonRowColor, m_def_RowCriteriaColor
   PropBag.WriteProperty "Width", m_sngColWidth, 800
   PropBag.WriteProperty "Name", m_strColName, ""
   PropBag.WriteProperty "Index", m_CurrIndex, -1
   PropBag.WriteProperty "ColID", m_lonColID, -1
   PropBag.WriteProperty "UseColors", m_blnColorNumbers, False
   PropBag.WriteProperty "FORECOLOR", m_lonForecolor, -1
   PropBag.WriteProperty "BACKCOLOR", m_lonBackcolor, -1
   PropBag.WriteProperty "COMBOLIST", m_ComboList, ""
   AllProperties = PropBag.Contents

End Property

Friend Property Let AllProperties(Value() As Byte)
   
   Dim PropBag As New PropertyBag
   
   PropBag.Contents = Value()
    
   m_ComboList = PropBag.ReadProperty("COMBOLIST", "")
   m_strHeaderText = PropBag.ReadProperty("HeaderText", "")
   m_intColAlignment = PropBag.ReadProperty("TextAlign", m_def_ColAlignment)
   m_ColFormat = PropBag.ReadProperty("ColFormat", m_def_ColFormat)
   m_strFormat = PropBag.ReadProperty("FormatString", "")
   m_strCustFormat = PropBag.ReadProperty("CustomFormatString", "")
   m_blnVisible = PropBag.ReadProperty("Visible", m_def_Visible)
   m_blnEnabled = PropBag.ReadProperty("Enabled", m_def_enabled)
   m_intEditType = PropBag.ReadProperty("EditType", m_def_ColEditType)
   m_intDecimal = PropBag.ReadProperty("Decimals", 2)
   m_intValidate = PropBag.ReadProperty("ValidationType", m_def_Validate)
   m_blnShowTotal = PropBag.ReadProperty("ShowTotal", m_def_ShowTotal)
   m_blnUseCriteria = PropBag.ReadProperty("UseCriteria", m_def_UseCriteria)
   RowCriteria = PropBag.ReadProperty("RowCriteria", "")
   m_lonRowColor = PropBag.ReadProperty("RowCriteriaColor", m_def_RowCriteriaColor)
   m_sngColWidth = PropBag.ReadProperty("Width", 800)
   m_strColName = PropBag.ReadProperty("Name", "")
   m_CurrIndex = PropBag.ReadProperty("Index", -1)
   m_lonColID = PropBag.ReadProperty("ColID", -1)
   m_blnColorNumbers = PropBag.ReadProperty("UseColors", False)
   m_lonForecolor = PropBag.ReadProperty("FORECOLOR", -1)
   m_lonBackcolor = PropBag.ReadProperty("BACKCOLOR", -1)
      
End Property

Private Sub Class_Initialize()
   m_intColAlignment = m_def_ColAlignment
   m_intColAlignmentHdr = m_def_ColAlignment
   m_strFormat = vbNullString
   m_ColFormat = m_def_ColFormat
   m_intDecimal = 2
   m_CurrIndex = -1
   m_lonColID = -1
   m_strColName = vbNullString
   m_intSorted = m_def_Sorted
   m_blnEnabled = m_def_enabled
   m_intEditType = m_def_ColEditType
   m_strText = vbNullString
   m_Datatype = 202
   m_blnVisible = m_def_Visible
   m_strHeaderText = vbNullString
   m_lonRowColor = m_def_RowCriteriaColor
   m_sngColWidth = 800
   m_blnUseCriteria = m_def_UseCriteria
   m_blnShowTotal = m_def_ShowTotal
   m_blnColorNumbers = False
   m_lonForecolor = -1
   m_lonBackcolor = -1
End Sub
