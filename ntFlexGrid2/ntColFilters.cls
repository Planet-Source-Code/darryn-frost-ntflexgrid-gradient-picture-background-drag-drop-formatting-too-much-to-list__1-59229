VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntColFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2219
Attribute VB_Description = "ntFlexGridControl.ntColFilters"
Option Explicit
Option Compare Text

Private m_colFilters As Collection
Private m_sName As String

Public Property Let FilterName(ByVal sName As String)
   m_sName = sName
End Property

Public Property Get FilterName() As String
   FilterName = m_sName
End Property

Public Function NewFilter() As ntFilter
Attribute NewFilter.VB_HelpID = 3260
   Dim objFilter As New ntFilter
   
   Set NewFilter = objFilter
   
   Set objFilter = Nothing
   
End Function

' Creates a new column, inserts into collection
Public Sub Add(ByRef nFilter As ntFilter, sKey As String)
Attribute Add.VB_HelpID = 3270
  
    m_colFilters.Add nFilter, sKey
            
End Sub

Public Sub Clear()
Attribute Clear.VB_HelpID = 3280
    
    Dim lonTotal As Long
    Dim lonCounter As Long
    
    lonTotal = m_colFilters.count
    lonCounter = 1
    
    Do While lonCounter <= lonTotal
      Remove 0
      lonCounter = lonCounter + 1
    Loop
         
End Sub

Public Sub Remove(ByVal varKey As Variant)
Attribute Remove.VB_HelpID = 3290
       
  If IsNumeric(varKey) Then varKey = varKey + 1
  
  m_colFilters.Remove varKey

End Sub

Public Function Exists(ByVal varKey As Variant) As Boolean
Attribute Exists.VB_HelpID = 3300
    
    On Error GoTo Exists_Error
            
    If IsNumeric(varKey) Then varKey = varKey + 1
    
    If Not (m_colFilters.Item(varKey) Is Nothing) Then Exists = True
    
    Exit Function
    
Exists_Error:
    Exists = False
End Function

Public Property Get Item(ByVal varKey As Variant) As ntFilter
Attribute Item.VB_HelpID = 3310
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    
On Error GoTo Item_Error
     
   If IsNumeric(varKey) Then varKey = varKey + 1
   
   Item = m_colFilters.Item(varKey)
   
Exit Property

Item_Error:
    Err.Raise 10009, "ntGrid.Filter", "No such Filter Exists"
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2226
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = m_colFilters.[_NewEnum]

End Property

Public Property Get count() As Long
Attribute count.VB_HelpID = 3330
    
    count = m_colFilters.count

End Property

Public Property Get Properties() As Byte()

   Dim nFlt As ntFilter
   Dim PropBag As New PropertyBag
   Dim l As Long

   PropBag.WriteProperty "Name", m_sName, ""

   For l = 1 To m_colFilters.count
      Set nFlt = m_colFilters(l)
      PropBag.WriteProperty "FILTER" & CStr(l), nFlt.Properties, Nothing
   Next l

   Properties = PropBag.Contents

End Property

Public Property Let Properties(mCols() As Byte)

   Dim nFlt As ntFilter
   Dim PropBag As New PropertyBag
   Dim l As Long
   
   Set m_colFilters = New Collection
   
   PropBag.Contents = mCols

   m_sName = PropBag.ReadProperty("Name", "")

   l = 1

   Do While Not Len(PropBag.ReadProperty("FILTER" & CStr(l), "")) = 0
      Set nFlt = New ntFilter
      nFlt.Properties = PropBag.ReadProperty("FILTER" & CStr(l))
      m_colFilters.Add nFlt, "FILTER" & CStr(l)
      l = l + 1
   Loop
  
End Property

Private Sub Class_Initialize()
   Set m_colFilters = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_colFilters = Nothing
End Sub
