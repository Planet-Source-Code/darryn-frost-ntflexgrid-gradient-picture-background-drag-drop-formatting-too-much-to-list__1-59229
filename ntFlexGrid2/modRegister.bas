Attribute VB_Name = "modRegister"
Option Explicit

'--------------------------------------------------------------------------
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'   If Not bIsOK Then Exit Sub
'
'--------------------------------------------------------------------------
'Private Sub UserControl_InitProperties()
'   modRegister.bIsOK = Not modRegister.IS_DEMO
'   If modRegister.IS_DEMO Then Call modRegister.CheckReg
'
'   If Not modRegister.bIsOK Then
'        'Do some stuff
'
'---------------------------------------------------------------------------
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'   modRegister.bIsOK = Not modRegister.IS_DEMO
'   If modRegister.IS_DEMO Then Call modRegister.CheckReg
'
'   If Not bIsOK Then
'     'Do some stuff
'----------------------------------------------------------------------------
'Then do whatever in your modules with bIsOk
'----------------------------------------------------------------------------
    
Public Const IS_DEMO = False

Private Const NT_EXT As String = "006"
Private Const NT_EVAL As String = "CntFlexGrid2"
Private Const NT_MSGTITLE As String = "ntFlexGrid2"
Public Const NT_PRODTITLE = "nth Technologies ntFlexGrid2"
'This is the name used to register with - needs to match Database table tag
Public Const NT_ControlTag As String = "ntflexgrid2"

Private Const S_TRIALEXP As String = "Your 30 day registration period has expired. To continue using this product, you must purchase it from www.nthtechnologies.com"

Public bIsOK As Boolean
Public bIsRegistered As Boolean

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Function GetSysDir() As String
   Dim s As String
   s = Space(256)
   GetSystemDirectory s, 255
   s = Trim$(s)
   GetSysDir = Left$(s, Len(s) - 1)
End Function

Private Function GetFile(ByVal sPath As String) As String
      Dim MyFile As Long
      MyFile = FreeFile
      Open sPath For Input As MyFile
      GetFile = Input(19, #MyFile)
      GetFile = Trim$(GetFile)
      Close MyFile
End Function

Public Function CheckReg() As Boolean
   Dim MyFile As Long, i As Integer
   Dim sFile As String, s As String
   Dim sDate As String
   Dim sFName As String, sLName As String, sKey As String
   Dim arrFile() As String, arrBin() As String, arrnth() As String
        
   bIsOK = False
     
   s = GetSysDir
      
   Dim sCtrl As String, sNth As String
   
   sCtrl = GetSetting("Windows", "Controls", "Date" & NT_EXT, "")
   sNth = GetSetting("VBControls", "NthTech", "Control" & NT_EXT, "")
   
   If Dir(s & "\windll." & NT_EXT) = "" And sCtrl = "" And sNth = "" Then
      
      sDate = Format$(Now(), "mm/dd/yyyy hh:nn:ss")
            
      MyFile = FreeFile
      Open s & "\windll." & NT_EXT For Output As MyFile
      Print #MyFile, sDate
      Close #MyFile
       
      SaveSetting "Windows", "Controls", "Date" & NT_EXT, sDate
      SaveSetting "VBControls", "NthTech", "Control" & NT_EXT, sDate
      
      MsgBox "Thank you for trying the " & NT_PRODTITLE & " evalution. You have thirty days to use it as you wish.", vbInformation + vbOKOnly, "New Trial"
      bIsOK = True
      
   Else
            
      'read the file to see what date it has in it
      If Dir(s & "\windll." & NT_EXT) = "" Or sCtrl = "" Or sNth = "" Then
         SaveSetting "Windows", "Controls", "Date" & NT_EXT, DateAdd("y", 1, Now)
         SaveSetting "VBControls", "NthTech", "Control" & NT_EXT, DateAdd("y", -1, Now)
         MsgBox "Your 30 day registration period has expired. To continue using the " & NT_PRODTITLE & ", you must purchase it from www.nthtechnologies.com", vbInformation + vbOKOnly
      Else
         MyFile = FreeFile
         sFile = GetFile(s & "\windll." & NT_EXT)
         If StrComp(sCtrl, sNth, vbTextCompare) <> 0 Or StrComp(sFile, sCtrl, vbTextCompare) <> 0 _
            Or StrComp(sFile, sNth, vbTextCompare) <> 0 Or Abs(DateDiff("d", CDate(sFile), Now())) > 30 Then
            SaveSetting "Windows", "Controls", "Date" & NT_EXT, DateAdd("y", 1, Now)
            SaveSetting "VBControls", "NthTech", "Control" & NT_EXT, DateAdd("y", -1, Now)
            MsgBox "Your 30 day registration period has expired. To continue using the " & NT_PRODTITLE & ", you must purchase it from www.nthtechnologies.com", vbInformation + vbOKOnly
         Else
             bIsOK = True
         End If
      End If
   End If

End Function


