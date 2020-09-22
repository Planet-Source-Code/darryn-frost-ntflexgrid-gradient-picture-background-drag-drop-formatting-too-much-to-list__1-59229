VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   9120
      TabIndex        =   15
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   7170
      TabIndex        =   14
      Top             =   60
      Width           =   1755
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "ICON"
         EndProperty
      EndProperty
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Dim arrVals() As String
   Dim i As Integer
   Dim Counter As Integer
   
   ReDim arrVals(gdUser.RowSel - gdUser.Row)
   
   Counter = 0
   
   For i = gdUser.Row To gdUser.RowSel
      arrVals(Counter) = gdUser.CellValue(i, gdUser.Col)
      Counter = Counter + 1
   Next i
   
   gdUser.ApplyFormattedFilter gdUser.Columns(gdUser.Col).Name, arrVals, True, True
   
End Sub

Private Sub Command2_Click()
   gdUser.resetcolumns
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
      .UseFieldNamesAsHeader = False
      .AutoSizeColumns = True
      .AllowDblClickFilter = True
      '.AllowUserResizing = nfgResizeNone
      .SelectionMode = nfgSelectionFree
      
      .GridMode = nfgGridmodeUnbound
      
      'Add columns to grid
      For i = 1 To 10
         Set cCol = .Columns.NewColumn
         cCol.Name = "COL" & CStr(i)
         cCol.ColFormat = nfgGeneral
         .Columns.Insert cCol, cCol.Name
        ' Set .colheaderPicture(i - 1, nfgAlignLeft) = Picture1.Image
      Next i
      
      .Columns("COL2").Visible = False
      
      .Columns("COL1").TextAlignment = nfgAlignRight
      .Columns("COL1").ColFormat = nfgPictureText
      .Columns("COL2").Enabled = True
      .Columns("COL2").ColFormat = nfgBooleanCheckBox
      
      .Columns("COL4").Enabled = True
      .Columns("COL4").EditType = nfgTextBox
                 
      .Display
            
      .Redraw = False
               
      Dim sVals() As Variant
      Dim arr() As Variant
      
      arr = Array("COL1", "COL2", "COL3", "COL4", "COL5", "COL6", "COL7", "COL8", "COL9", "COL10")
      
      ReDim sVals(0 To .Cols - 1)
           
      For i = 0 To 999
         For j = 0 To .Cols - 1
            If j = 1 Then
               sVals(j) = "True"
            Else
               sVals(j) = "Row" & CStr(i) & ":COL" & CStr(j)
            End If
         Next j
         .AddRow arr, sVals, True
      Next i
                                  
      ' gduser.AddBlankRows 6
      .Redraw = True
      
      gdUser.CustomMenuAddItem "Example menu item 1", "EX1"
      gdUser.CustomMenuAddItem "Example menu item 2", "EX1"
      .Refresh
      
   End With
   
End Sub

Private Sub Form_Resize()
  gdUser.Move 30, 990, Me.ScaleWidth - 60, Me.ScaleHeight - 1005
End Sub

Private Sub gduser_BeforeFilter(nFilter As ntFxGd2.ntFilter, bCancel As Boolean)
   If chkFilter.Value = vbChecked Then Debug.Print "BeforeFilter: " & nFilter.FieldName
End Sub

Private Sub gduser_OnFilter(ByVal nFilter As ntFxGd2.ntFilter)
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colsel
   If chkFilter.Value = vbChecked Then Debug.Print "OnFilter: " & nFilter.FieldName
End Sub

Private Sub gduser_OnFilterRemove()
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.mousecol
   txtColSel.Text = gdUser.colsel
   If chkFilter.Value = vbChecked Then Debug.Print "OnFilterRemove"
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
   txtColSel.Text = gdUser.colsel
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
   txtColSel.Text = gdUser.colsel
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
   txtColSel.Text = gdUser.colsel
   If chkRow.Value = vbChecked Then
      change = change + 1
      Debug.Print CStr(change) & ": ROWCHANGE"
   End If
End Sub

Private Sub gduser_SelChange()
   Static change As Long
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colsel
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
   txtColSel.Text = gdUser.colsel
End Sub

Private Sub gduser_OnSortRemove()
   If chkSort.Value = vbChecked Then Debug.Print " ONSORTREMOVE"
   txtRow.Text = gdUser.Row
   txtCol.Text = gdUser.Col
   txtRowSel.Text = gdUser.RowSel
   txtColSel.Text = gdUser.colsel
End Sub

