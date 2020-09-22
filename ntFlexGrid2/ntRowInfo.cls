VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntRowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2163
Attribute VB_Description = "ntFlexGridControl.ntRowInfo"
Option Explicit
Option Compare Text


Private m_colColumns As Collection

' Creates a new RowInfo obj, inserts into collection, and returns the reference
Friend Function NewRowInfo() As RowInfo
                  
    Set NewRowInfo = New RowInfo
        
End Function

' Creates a new column, inserts into collection
Friend Sub Insert(ByRef objRowinfo As RowInfo, sKey As String)
  
    m_colColumns.Add objRowinfo, sKey
            
End Sub

Public Sub Clear()
Attribute Clear.VB_HelpID = 2740
    
    Dim lonTotal As Long
    Dim lonCounter As Long
    
    lonTotal = m_colColumns.count
    lonCounter = 1
    
    Do While lonCounter <= lonTotal
      Remove 0
      lonCounter = lonCounter + 1
    Loop
         
End Sub

Public Sub Delete(ByVal vBookmark As Variant)
Attribute Delete.VB_HelpID = 2750
   Dim l As Long
      
On Error GoTo Delete_Error

   m_colColumns.Remove "ID" & CStr(vBookmark)
 
Exit Sub

Delete_Error:
    Err.Raise 10009, "ntGrid.Columns", "No such Column Exists"
    
End Sub

Friend Sub Remove(ByVal varKey As Variant)
 
  If IsNumeric(varKey) Then varKey = varKey + 1
  
  m_colColumns.Remove varKey
    
End Sub

Public Function Exists(ByVal vBookmark As Variant) As Boolean
Attribute Exists.VB_HelpID = 2760
    
On Error GoTo Exists_Error
       
   If Not (m_colColumns.Item("ID" & CStr(vBookmark)) Is Nothing) Then Exists = True
    
   Exit Function
    
Exists_Error:
    Exists = False
End Function

Friend Function RowItemExists(ByVal varKey As Variant) As Boolean
    
    On Error GoTo Exists_Error
    
    If IsNumeric(varKey) Then varKey = varKey + 1
    
    If Not (m_colColumns.Item(varKey) Is Nothing) Then RowItemExists = True
    
    Exit Function
    
Exists_Error:
    RowItemExists = False
End Function

Friend Property Get RowItem(ByVal varKey As Variant) As RowInfo
   
   On Error GoTo Item_Error
   
   If IsNumeric(varKey) Then varKey = varKey + 1
   
   Set RowItem = m_colColumns.Item(varKey)
   
Exit Property

Item_Error:
    Err.Raise 10009, "ntGrid.Columns", "No such Column Exists"

End Property

Public Property Get Item(ByVal vBookmark As Variant) As RowInfo
Attribute Item.VB_HelpID = 2770
Attribute Item.VB_MemberFlags = "200"

On Error GoTo Item_Error

   Set Item = m_colColumns.Item("ID" & CStr(vBookmark))

Exit Property

Item_Error:
    Err.Raise 10009, "ntGrid.Columns", "No such Column Exists"
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2167
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = m_colColumns.[_NewEnum]

End Property

Public Property Get count() As Long
Attribute count.VB_HelpID = 2790
    
    count = m_colColumns.count

End Property

Private Sub Class_Initialize()
   Set m_colColumns = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_colColumns = Nothing
End Sub
