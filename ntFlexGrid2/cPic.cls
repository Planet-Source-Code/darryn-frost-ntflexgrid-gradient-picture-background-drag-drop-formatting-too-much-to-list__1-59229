VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Private m_strName          As String
Private m_Alignment        As nfgGridAlignment
Private m_vBookmark        As Variant
Private m_lCol             As Long
Private m_pic              As StdPicture
Private m_bHasPic          As Boolean

Public Property Get HasPic() As Boolean
   HasPic = m_bHasPic
End Property

Friend Property Let ColName(ByVal strName As String)
   m_strName = strName
End Property

Public Property Get ColName() As String
   ColName = m_strName
End Property

Friend Property Let ColID(ByVal lID As Long)
   m_lCol = lID
End Property

Public Property Get ColID() As Long
   ColID = m_lCol
End Property

Friend Property Let RowBookmark(ByVal vBookmark As Variant)
   m_vBookmark = vBookmark
End Property

Public Property Get RowBookmark() As Variant
   RowBookmark = m_vBookmark
End Property

Friend Property Let PictureAlignment(ByVal eAlign As nfgGridAlignment)
   m_Alignment = eAlign
End Property

Public Property Get PictureAlignment() As nfgGridAlignment
   PictureAlignment = m_Alignment
End Property

Friend Property Set CellPicture(ByVal stdPic As Picture)
   Set m_pic = stdPic
   m_bHasPic = Not (m_pic Is Nothing)
End Property

Public Property Get CellPicture() As Picture
   Set CellPicture = m_pic
End Property
