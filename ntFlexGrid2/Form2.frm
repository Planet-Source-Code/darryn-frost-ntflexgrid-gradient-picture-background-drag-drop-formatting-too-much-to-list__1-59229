VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "*\AntFxGd2.vbp"
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin ntFxGd2.ntFlexGrid2 gdUser 
      Height          =   5475
      Left            =   690
      TabIndex        =   21
      Top             =   1470
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   9657
      GradientEndColor=   16744576
      BackStyle       =   2
      AllowMoveCols   =   -1  'True
      GridMode        =   0   'False
      CheckEnabledPic =   "Form2.frx":0000
      UnCheckEnabledPic=   "Form2.frx":031A
      CheckDisabledPic=   "Form2.frx":0634
      UnCheckDisabledPic=   "Form2.frx":094E
      ScaleHeight     =   5415
      ScaleWidth      =   10485
      Object.Height          =   5475
      Object.Width           =   10545
      MouseIcon       =   "Form2.frx":0C68
      AllowFilter     =   -1  'True
      AllowSort       =   -1  'True
      ColorByRow      =   -1  'True
      BackColorBkg    =   -2147483633
      BackColorFixed  =   -2147483633
      BackColorSel    =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ForeColorFixed  =   -2147483640
      ForeColorSel    =   -2147483634
      TextStyleFixed  =   0
      DblClickFilter  =   -1  'True
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   435
      Left            =   10590
      TabIndex        =   20
      Top             =   3330
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   315
      Left            =   10920
      TabIndex        =   19
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   9540
      TabIndex        =   18
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   8520
      TabIndex        =   15
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   7170
      TabIndex        =   14
      Top             =   60
      Width           =   1275
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Debug.Print Filter Events"
      Height          =   285
      Left            =   4470
      TabIndex        =   13
      Top             =   300
      Width           =   2175
   End
   Begin VB.CheckBox chkDrag 
      Caption         =   "Debug.Print ColDrag Events"
      Height          =   285
      Left            =   4470
      TabIndex        =   12
      Top             =   0
      Width           =   2415
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Debug.Print Sort Events"
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   300
      Width           =   2175
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "Debug.Print SelChange"
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   0
      Width           =   2175
   End
   Begin VB.CheckBox chkCol 
      Caption         =   "Debug.Print ColChange"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   300
      Width           =   2175
   End
   Begin VB.CheckBox chkRow 
      Caption         =   "Debug.Print RowChange"
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txtColSel 
      Height          =   315
      Left            =   6030
      TabIndex        =   4
      Top             =   630
      Width           =   795
   End
   Begin VB.TextBox txtRowSel 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   630
      Width           =   795
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   630
      Width           =   795
   End
   Begin VB.TextBox txtRow 
      Height          =   315
      Left            =   570
      TabIndex        =   0
      Top             =   690
      Width           =   795
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   900
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0C84
            Key             =   "ICON"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":121E
            Key             =   "DOLLAR"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      Caption         =   "Label2"
      Height          =   435
      Index           =   1
      Left            =   9090
      TabIndex        =   17
      Top             =   570
      Width           =   1725
   End
   Begin VB.Label lblTotal 
      Caption         =   "Label2"
      Height          =   435
      Index           =   0
      Left            =   7170
      TabIndex        =   16
      Top             =   570
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ColSel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   660
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RowSel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   6
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Col"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   660
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   405
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveGrid(ByRef sGrid As ntFlexGrid2, ByVal sFileName As String)
   Dim myFile As Long
   Dim v As Variant
   
   
   Dim pb As PropertyBag
   
   Call pb.WriteProperty("AUTOSIZE", sGrid.AutoSizeColumns)
   
      
   v = sGrid.Columns.Properties
   myFile = FreeFile
   Open sFileName For Binary Access Write As myFile
   Put #myFile, , v
   Close myFile

End Sub

Private Sub Command2_Click()
  
   Dim myArray() As Byte
   Dim myFile As Long
   Dim v As Variant
   Dim i As Integer
   Dim j As Integer
   
   Dim pb As PropertyBag
   Set pb = New PropertyBag
      
   Set gdUser.Recordset = Nothing
   
   gdUser.GridMode = nfgGridmodeUnbound
   
   myFile = FreeFile
   Open App.Path & "\Grid.bin" For Binary Access Read As myFile
   Get #myFile, , v
   Close myFile
                         
   myArray = v
   
   pb.Contents = myArray
               
   gdUser.Columns.Properties = pb.Contents
   gdUser.Display
   gdUser.Redraw = False
   
  
  With gdUser
  
   Dim sVals() As Variant
      Dim arr() As Variant
      
      arr = Array("COL1", "COL2", "COL3", "COL4", "COL5", "COL6", "COL7", "COL8", "COL9", "COL10", "COL11", "COL12", "COL13", "COL14", "COL15", "COL16", "COL17", "COL18", "COL19", "COL20")
      
      ReDim sVals(0 To .Cols - 1)
           
      For i = 0 To 199
         For j = 0 To .Cols - 1
            If j = 1 Then
               sVals(j) = "True"
            ElseIf j = 3 Then
               sVals(j) = 50
            ElseIf j = 4 Then
               sVals(j) = 100
            Else
               sVals(j) = "Row" & CStr(i) & ":COL" & CStr(j)
            End If
         Next j
         .AddRow arr, sVals, True
      Next i
                                  
      
      '.Redraw = True
      
      gdUser.CustomMenuAddItem "Example menu item 1", "EX1"
      gdUser.CustomMenuAddItem "Example menu item 2", "EX1"
      
      
      '.Refresh
     .Redraw = True
   End With
      
  
  
  
End Sub

Private Sub Command3_Click()
         
   
 gdUser.CopySelectedCells False
        
End Sub

Private Sub Command5_Click()
Static change As Long
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkCol.Value = vbChecked Then
      change = change + 1
      Debug.Print CStr(change) & ": COLCHANGE"
   End If
End Sub

Private Sub Form_Load()
   gdUser.HScrollBar.ImageList = ImageList1
   gdUser.HScrollBar.AddButton "TEST", "TEST", 0, 0, esbcButtonPositionLeftTop
   Call setunboundgrid
End Sub

Private Sub setunboundgrid()
   Dim cCol As ntFxGd2.ntColumn
   Dim i As Integer
   Dim j As Integer
      
   With gdUser
      '.UseFieldNamesAsHeader = False
      '.AutoSizeColumns = True
      '.AllowDblClickFilter = True
      '.AllowUserResizing = nfgResizeNone
      '.SelectionMode = nfgSelectionFree
      '.Highlight = nfgHighlightAlways
      
      .GridMode = nfgGridmodeUnbound
      
      'Add columns to grid
      For i = 1 To 20
         Set cCol = .Columns.NewColumn
         cCol.Name = "COL" & CStr(i)
         cCol.ColFormat = nfgGeneral
         .Columns.Insert cCol, cCol.Name
        ' Set .colheaderPicture(i - 1, nfgAlignLeft) = Picture1.Image
      Next i
      
'      .Columns("COL2").Visible = False
'
'      .Columns("COL1").TextAlignment = nfgAlignRight
'      .Columns("COL1").ColFormat = nfgPictureText
'
      .Columns("COL2").Enabled = True
      .Columns("COL2").ColFormat = nfgBooleanCheckBox

      .Columns("COL4").UseCriteria = True
      .Columns("COL4").RowCriteria = "2,3,9"
      .Columns("COL4").RowCriteriaColor = &H8000000D
      
      .Columns("COL5").ColFormat = nfgCustomFormat
      .Columns("COL5").CustomFormatString = "#,##0.00;(#,##0.00)"
            
      '.Columns("COL10").ColFormat = nfgPicture
      
      .Columns("COL8").Enabled = True
      .Columns("COL8").EditType = nfgComboBox
      .Columns("COL8").ComboList = "Selection 1,Selection 2,Selection3,Selection4"
      
                  
      .Display
            
      '.Redraw = False
               
      Dim sVals() As Variant
      Dim arr() As Variant
      
      arr = Array("COL1", "COL2", "COL3", "COL4", "COL5", "COL6", "COL7", "COL8", "COL9", "COL10", "COL11", "COL12", "COL13", "COL14", "COL15", "COL16", "COL17", "COL18", "COL19", "COL20")
      
      ReDim sVals(0 To .Cols - 1)
           
      For i = 0 To 99
        AddARow
        'Set .CellPicture(9, i) = ImageList1.ListImages("DOLLAR").ExtractIcon
      Next i
           
      '.Redraw = True
      
      gdUser.CustomMenuAddItem "Example menu item 1", "EX1"
      gdUser.CustomMenuAddItem "Example menu item 2", "EX1"
      
      .Refresh
      
   End With
      
      
   gdUser.TopRow = 0
   gdUser.Row = 3
   gdUser.Col = 0
   gdUser.RowSel = 2
   'gduser.SetFocus
         
   lblTotal(0).Caption = gdUser.ColumnTotal("COL4")
   lblTotal(1).Caption = gdUser.ColumnTotal("COL5")
   
End Sub

Private Sub AddARow()
Dim i As Long
Dim j As Long

   Dim sVals() As Variant
      Dim arr() As Variant
      
      With gdUser
      arr = Array("COL1", "COL2", "COL3", "COL4", "COL5", "COL6", "COL7", "COL8", "COL9", "COL10", "COL11", "COL12", "COL13", "COL14", "COL15", "COL16", "COL17", "COL18", "COL19", "COL20")
      
      ReDim sVals(0 To .Cols - 1)
           
         For j = 0 To .Cols - 1
            If j = 1 Then
               sVals(j) = "False"
            ElseIf j = 3 Then
               sVals(j) = gdUser.Rows
            ElseIf j = 4 Then
               sVals(j) = gdUser.Rows * 10
            Else
               sVals(j) = "Row" & CStr(i) & ":COL" & CStr(j)
            End If
         Next j
         .AddRow arr, sVals, True
      End With
      
      
End Sub

Private Sub Form_Resize()
  gdUser.Move 30, 990, Me.ScaleWidth - 2000    ', Me.ScaleHeight - 1005
End Sub

Private Sub gduser_AfterEdit(ByVal pEditType As ntFxGd2.nfgEditType, ByVal vNewValue As Variant, ByVal pColName As String, arrlonRowID() As Long)
   lblTotal(0).Caption = gdUser.ColumnTotal("COL4")
   lblTotal(1).Caption = gdUser.ColumnTotal("COL5")
End Sub

Private Sub gduser_BeforeFilter(nFilter As ntFxGd2.ntFilter, bCancel As Boolean)
   If chkFilter.Value = vbChecked Then Debug.Print "BeforeFilter: " & nFilter.FieldName
End Sub

Private Sub gduser_KeyDown(KeyCode As Integer, Shift As Integer)
   Debug.Print "KD " & KeyCode
End Sub

Private Sub gduser_KeyPress(KeyAscii As Integer)
   Debug.Print "KP " & KeyAscii
End Sub

Private Sub gduser_KeyUp(KeyCode As Integer, Shift As Integer)
   Debug.Print "KU " & KeyCode
End Sub

Private Sub gduser_OnFilter(ByVal nFilter As ntFxGd2.ntFilter)
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkFilter.Value = vbChecked Then Debug.Print "OnFilter: " & nFilter.FieldName
   lblTotal(0).Caption = gdUser.ColumnTotal("COL4")
   lblTotal(1).Caption = gdUser.ColumnTotal("COL5")
End Sub

Private Sub gduser_OnFilterRemove()
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.MouseCol
   txtColSel.Text = gdUser.colSel
   If chkFilter.Value = vbChecked Then Debug.Print "OnFilterRemove"
   lblTotal(0).Caption = gdUser.ColumnTotal("COL4")
   lblTotal(1).Caption = gdUser.ColumnTotal("COL5")
End Sub


Private Sub gduser_Click()
'   Debug.Print gduser.Row
'   Debug.Print gduser.CellValue(gduser.Row, 1)
End Sub



'***************************************************************************************************
'COL MOVING -------------------------------------------------------------------------------
Private Sub gduser_ColBeginDrag(ByVal sDragColName As String, pCancel As Boolean)
  If chkDrag.Value = vbChecked Then Debug.Print "ColBeginDrag: " & sDragColName
End Sub

Private Sub gduser_ColDragDrop(ByVal sDragColName As String, ByVal lNewColPos As Long, pCancel As Boolean)
  If chkDrag.Value = vbChecked Then Debug.Print "ColDragDrop: " & sDragColName & " AT Col " & lNewColPos
End Sub

Private Sub gduser_ColDragOver(ByVal sDragColName As String, ByVal lNewColPos As Long, pCancel As Boolean)
  If chkDrag.Value = vbChecked Then Debug.Print "ColDragOver: " & sDragColName & " ON Col " & lNewColPos
End Sub

Private Sub gduser_ColEndDrag(ByVal sDragColName As String, ByVal lNewColPos As Long)
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkDrag.Value = vbChecked Then Debug.Print "ColEndDrag: " & sDragColName
End Sub

'******************************************************************************************************



Private Sub gduser_hScrollButtonClick(ByVal lButton As Long)
   'Debug.Print "BCLICK " & lButton
End Sub

Private Sub gduser_ColChange(ByVal ntCol As ntFxGd2.ntColumn)
Static change As Long
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkCol.Value = vbChecked Then
      change = change + 1
      Debug.Print CStr(change) & ": COLCHANGE"
   End If
End Sub

Private Sub gduser_RowChange(ByVal lRow As Long)
Static change As Long
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkRow.Value = vbChecked Then
      change = change + 1
      Debug.Print CStr(change) & ": ROWCHANGE"
   End If
   lblTotal(0).Caption = gdUser.ColumnTotal("COL4")
   lblTotal(1).Caption = gdUser.ColumnTotal("COL5")
End Sub

Private Sub gduser_SelChange()
   Static change As Long
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
   If chkSel.Value = vbChecked Then
      change = change + 1
      Debug.Print CStr(change) & ": SELCHANGE"
   End If
End Sub

Private Sub gduser_BeforeSort(ByVal sBeginCol As String, ByVal sEndCol As String, bCancel As Boolean)
   If chkSort.Value = vbChecked Then Debug.Print " BEFORESORT: BEGINCOL: " & sBeginCol & " : ENDCOL: " & sEndCol
End Sub

Private Sub gduser_OnSort(ByVal sBeginCol As String, ByVal sEndCol As String)
   If chkSort.Value = vbChecked Then Debug.Print " ONSORT: BEGINCOL: " & sBeginCol & " : ENDCOL: " & sEndCol
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
End Sub

Private Sub gduser_OnSortRemove()
   If chkSort.Value = vbChecked Then Debug.Print " ONSORTREMOVE"
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colSel
End Sub

