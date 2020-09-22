Attribute VB_Name = "ntGridValidate"
Option Explicit

Public Function ValidateNumericText(ByRef TextControl As Control, ByVal intDecimals As Integer) As Boolean
  Dim blnValid As Boolean
  Dim intDec As Integer
  Dim intStart As Integer
      
  blnValid = True
  
  'dont check if nothing entered
  If Len(TextControl.Text) = 0 Then
    ValidateNumericText = False
    Exit Function
  End If
  
  If InStr(InStr(1, TextControl.Text, ".", vbTextCompare) + 1, _
        TextControl.Text, ".", vbTextCompare) > 0 Then blnValid = False
                 
  intDec = InStr(1, TextControl.Text, ".", vbTextCompare)
  
  If (Len(TextControl.Text) - intDec) > intDecimals Then blnValid = False
      
  If Not (IsNumeric(TextControl.Text) = True) Then
    blnValid = False
  End If
        
  ValidateNumericText = blnValid
  
End Function
 
Public Sub EnforceNumericText(ByRef TextControl As Control, ByVal intDecimals As Integer)
  Dim blnValid As Boolean
  Dim intDec As Integer
  Dim intStart As Integer
  
  intStart = TextControl.SelStart
  
  'dont check if nothing entered
  If Len(TextControl.Text) = 0 Then Exit Sub
  
  intDec = InStr(InStr(1, TextControl.Text, ".", vbTextCompare) + 1, _
        TextControl.Text, ".", vbTextCompare)
        
  If intDec > 0 Then TextControl.Text = Left$(TextControl.Text, intDec - 1)
             
  intDec = InStr(1, TextControl.Text, ".", vbTextCompare)
  
  If intDec > 0 Then TextControl.Text = Left$(TextControl.Text, intDec + intDecimals)
      
  TextControl.SelStart = intStart
   
End Sub
 
Public Function ValidateNumericKey(ByVal pAllowDecimal As Boolean, ByRef KeyAscii As Integer) As Integer
  If KeyAscii = 46 And pAllowDecimal = False Then
    ValidateNumericKey = 0
    Exit Function
  End If
  If Not (IsNumeric(Chr$(KeyAscii)) Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 13) Then
    ValidateNumericKey = 0
  Else
    ValidateNumericKey = KeyAscii
  End If
End Function

Public Function ValidateAlphaKey(ByVal intKeyascii As Integer) As Integer
    
  If ((intKeyascii >= 65 And intKeyascii <= 90) Or (intKeyascii >= 96 And intKeyascii <= 122) Or intKeyascii = 8) Then
    ValidateAlphaKey = intKeyascii
  Else
    ValidateAlphaKey = 0
  End If
  
End Function

Public Function ValidateAlphaNumeric(ByVal intKeyascii As Integer) As Integer
  If ValidateAlphaKey(intKeyascii) = 0 And ValidateNumericKey(False, intKeyascii) = 0 Then
     ValidateAlphaNumeric = 0
  Else
     ValidateAlphaNumeric = intKeyascii
  End If
End Function

Public Function ValidateColLayoutNames(ByVal intKeyascii As Integer) As Integer
   ValidateColLayoutNames = intKeyascii
   If ValidateAlphaKey(intKeyascii) = 0 And ValidateNumericKey(False, intKeyascii) = 0 Then
      If Chr$(intKeyascii) <> "_" Then
         ValidateColLayoutNames = 0
      End If
   End If
End Function

