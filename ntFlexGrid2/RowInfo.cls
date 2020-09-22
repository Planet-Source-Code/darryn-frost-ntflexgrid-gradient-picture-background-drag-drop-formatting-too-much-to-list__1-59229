VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2185
Attribute VB_Description = "ntFlexGridControl.RowInfo"
Option Explicit
Option Compare Text

Private m_lonForecolor As Long
Private m_lonBackcolor As Long
Private m_varBookmark As Variant
Private m_varItemData As Variant

Public Property Let ForeColor(ByVal pColor As OLE_COLOR)
Attribute ForeColor.VB_HelpID = 2690
   m_lonForecolor = pColor
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = m_lonForecolor
End Property

Public Property Let BackColor(ByVal pColor As OLE_COLOR)
Attribute BackColor.VB_HelpID = 2700
   m_lonBackcolor = pColor
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_lonBackcolor
End Property

Public Property Get Bookmark() As Variant
Attribute Bookmark.VB_HelpID = 2710
   Bookmark = m_varBookmark
End Property

Friend Property Let Bookmark(ByVal vBookmark As Variant)
   m_varBookmark = vBookmark
End Property

Public Property Get ItemData() As Variant
Attribute ItemData.VB_HelpID = 2720
   ItemData = m_varItemData
End Property

Public Property Let ItemData(ByVal vItemData As Variant)
   m_varItemData = vItemData
End Property

Private Sub Class_Initialize()
   m_varItemData = -1 ' Nothing
   m_varBookmark = -1
   m_lonBackcolor = -1
   m_lonForecolor = -1
End Sub
