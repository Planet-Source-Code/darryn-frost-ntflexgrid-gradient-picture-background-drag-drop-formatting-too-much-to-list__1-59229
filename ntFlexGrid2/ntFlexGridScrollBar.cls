VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ntFxGdScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ctrl As vbalScrollButtonCtl

Friend Sub Bind(ByRef cCtl As vbalScrollButtonCtl)
   Set m_ctrl = cCtl
End Sub

Public Property Get ButtonKey(ByVal lButton As Long) As String
Attribute ButtonKey.VB_HelpID = 3500
   ButtonKey = m_ctrl.ButtonKey(lButton)
End Property

Public Property Get ButtonToolTipText(ByVal vKey As Variant) As String
Attribute ButtonToolTipText.VB_HelpID = 3510
   ButtonToolTipText = m_ctrl.ButtonToolTipText(vKey)
End Property

Public Property Let ButtonToolTipText(ByVal vKey As Variant, ByVal sText As String)
   m_ctrl.ButtonToolTipText(vKey) = sText
End Property

Public Property Get ButtonVisible(ByVal vKey As Variant) As Boolean
Attribute ButtonVisible.VB_HelpID = 3520
   ButtonVisible = m_ctrl.ButtonVisible(vKey)
End Property

Public Property Let ButtonVisible(ByVal vKey As Variant, ByVal bState As Boolean)
   m_ctrl.ButtonVisible(vKey) = bState
End Property

Public Property Get ButtonEnabled(ByVal vKey As Variant) As Boolean
Attribute ButtonEnabled.VB_HelpID = 3530
   ButtonEnabled = m_ctrl.ButtonEnabled(vKey)
End Property

Public Property Let ButtonEnabled(ByVal vKey As Variant, ByVal bState As Boolean)
   m_ctrl.ButtonEnabled(vKey) = bState
End Property

Public Property Get ButtonValue(ByVal vKey As Variant) As OLE_TRISTATE
Attribute ButtonValue.VB_HelpID = 3540
  ButtonValue = m_ctrl.ButtonValue(vKey)
End Property

Public Property Let ButtonValue(ByVal vKey As Variant, oValue As OLE_TRISTATE)
   m_ctrl.ButtonValue(vKey) = oValue
End Property

Public Property Let ImageList(vList As Variant)
Attribute ImageList.VB_HelpID = 3550
   m_ctrl.ImageList = vList
End Property

Public Sub AddButton( _
      Optional ByVal sKey As String = "", _
      Optional ByVal sToolTipText As String = "", _
      Optional ByVal lIconIndexUp As Long = -1, _
      Optional ByVal lIconIndexDown As Long = -1, _
      Optional ByVal ePosition As ESBCButtonPositionConstants = esbcButtonPositionDefault, _
      Optional ByVal bCheck As Boolean = False, _
      Optional ByVal sCheckGroup As String = "", _
      Optional ByVal bVisible As Boolean = True, _
      Optional ByVal vKeyBefore As Variant)
Attribute AddButton.VB_HelpID = 3490

      m_ctrl.AddButton sKey, sToolTipText, lIconIndexUp, lIconIndexDown, ePosition, _
                     bCheck, sCheckGroup, bVisible, vKeyBefore
End Sub
      
Public Property Get ButtonCount() As Long
Attribute ButtonCount.VB_HelpID = 3560
   ButtonCount = m_ctrl.ButtonCount
End Property

Public Property Get ButtonIndex(ByVal vKey As Variant) As Long
Attribute ButtonIndex.VB_HelpID = 3570
   ButtonIndex = m_ctrl.ButtonIndex(vKey)
End Property

Friend Property Get Visible() As Boolean
   Visible = m_ctrl.Visible
End Property

Friend Property Let Visible(ByVal bState As Boolean)
   m_ctrl.Visible = bState
End Property

Public Property Get SmallChange() As Long
   SmallChange = m_ctrl.SmallChange
End Property

Property Let SmallChange(ByVal lSmallChange As Long)
   m_ctrl.SmallChange = lSmallChange
End Property

Property Get ScrollEnabled() As Boolean
   ScrollEnabled = m_ctrl.ScrollEnabled
End Property

Property Let ScrollEnabled(ByVal bEnabled As Boolean)
   m_ctrl.ScrollEnabled = bEnabled
End Property

Property Get Min() As Long
   Min = m_ctrl.Min
End Property
Property Let Min(ByVal iMin As Long)
   m_ctrl.Min = iMin
End Property

Property Get Max() As Long
    Max = m_ctrl.Max
End Property
Property Let Max(ByVal iMax As Long)
  m_ctrl.Max = iMax
End Property

Property Get Value() As Long
   Value = m_ctrl.Value
End Property
Property Let Value(ByVal iValue As Long)
   m_ctrl.Value = iValue
End Property

Property Get LargeChange() As Long
   LargeChange = m_ctrl.LargeChange
End Property

Property Let LargeChange(ByVal iLargeChange As Long)
   m_ctrl.LargeChange = iLargeChange
End Property


Private Sub Class_Terminate()
   Set m_ctrl = Nothing
End Sub
