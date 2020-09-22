VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntColPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private m_colPics As New Collection

' Creates a new column, inserts into collection, and returns the reference
Friend Function Add(ByVal sColName As String, ByVal lColID As Long, _
                     ByVal vRowBkmrk As Variant) As cPic
    
    Dim newpic As cPic
    Set newpic = New cPic
        
    newpic.ColID = lColID
    newpic.ColName = sColName
    newpic.RowBookmark = vRowBkmrk
            
    m_colPics.Add newpic, sColName & CStr(vRowBkmrk)
        
    Set Add = newpic
    Set newpic = Nothing
    
End Function

Public Property Get Item(sColName As String, vBkmrk As Variant) As cPic
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
   Set Item = m_colPics.Item(sColName & CStr(vBkmrk))
End Property

Public Sub Clear()
    
    Dim lonTotal As Long
    Dim lonCounter As Long
    
    lonTotal = m_colPics.Count
    lonCounter = 1
    
    Do While lonCounter <= lonTotal
      m_colPics.Remove 1
      lonCounter = lonCounter + 1
    Loop
    
End Sub

Public Sub Remove(sColName As String, vBkmrk As Variant)
      
  m_colPics.Remove sColName & CStr(vBkmrk)

End Sub

Public Function Exists(sColName As String, vBkmrk As Variant) As Boolean
    
    On Error GoTo Exists_Error
            
    If Not (m_colPics.Item(sColName & CStr(vBkmrk)) Is Nothing) Then Exists = True
    
    Exit Function
    
Exists_Error:
    Exists = False
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = m_colPics.[_NewEnum]

End Property

Public Property Get Count() As Long
    
    Count = m_colPics.Count

End Property


