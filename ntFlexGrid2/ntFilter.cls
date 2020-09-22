VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2212
Attribute VB_Description = "ntFlexGridControl.ntFilter"
Option Explicit
Option Compare Text


Public Enum FilterType
   ntFilterInclude = -1
   ntFilterExclude = 0
End Enum

Private m_arrValues() As String
Private m_strField As String
Private m_Include As Integer

Public Property Let Values(ByRef arrValues() As String)
Attribute Values.VB_HelpID = 2810
   If UBound(arrValues) = -1 Then Err.Raise vbObjectError + 1006, "ntFilter", "Filter array must contain values"
   m_arrValues = arrValues
End Property

Public Property Get Values() As String()
   Values = m_arrValues
End Property

Public Property Let FieldName(ByVal strName As String)
Attribute FieldName.VB_HelpID = 2820
   m_strField = strName
End Property

Public Property Get FieldName() As String
   FieldName = m_strField
End Property

Public Property Let Include(ByVal bInclude As FilterType)
Attribute Include.VB_HelpID = 2830
   m_Include = bInclude
End Property

Public Property Get Include() As FilterType
  Include = m_Include
End Property

Public Property Get Properties() As Byte()

   Dim PropBag As New PropertyBag
   Dim l As Long

   PropBag.WriteProperty "FIELD", m_strField
   PropBag.WriteProperty "INCLUDE", m_Include
   
   For l = 0 To UBound(m_arrValues)
      PropBag.WriteProperty "VALUE" & CStr(l), m_arrValues(l)
   Next l

   Properties = PropBag.Contents

End Property

Public Property Let Properties(mCols() As Byte)

   Dim PropBag As New PropertyBag
   Dim l As Long
      
   PropBag.Contents = mCols

   m_strField = PropBag.ReadProperty("FIELD", "")
   m_Include = PropBag.ReadProperty("INCLUDE", True)
  
   l = 0

   Do While Not PropBag.ReadProperty("VALUE" & CStr(l), "NOFILTERVALUE") = "NOFILTERVALUE"
      ReDim Preserve m_arrValues(l)
      m_arrValues(l) = PropBag.ReadProperty("VALUE" & CStr(l), "")
      l = l + 1
   Loop
  
End Property
