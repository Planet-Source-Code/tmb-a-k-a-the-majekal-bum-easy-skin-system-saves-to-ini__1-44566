Attribute VB_Name = "modINIEdit"

Option Explicit
' there are three functions GetINISetting,
' SaveINISetting as DeleteINIsetting.
' They roughfuly correspond to
' the VB GetSetteing,SaveSetting and DeleteSetting for the registry
' main diffrence is the substution of the path to the
' ini file for the Application name.

' Use
' Dim lStatus As Long
'  lStatus = SaveINISetting(location of ini file, _
'                           section name, _
'                           name of object, _
'                           source of value)

'  txtValue.Text = GetINISetting(dlgB.FileName, _
'                                txtSectionName.Text, _
'                                txtKeyName.Text, _
'                                sDefault) 'Optional

'  lStatus = DeleteINISetting(dlgB.FileName, _
'                             txtSectionName.Text, _
'                             txtKeyName.Text)

' Louis Boldt 2/28/2002
' -------API functions for editing private ini files ---------------------------------------------------------------
' there are others but these two provide the
' all the basics nedded here
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpSection As String, _
     ByVal lpKey As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturned As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

' ----------------------------------------------------------------------
Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
   (ByVal lpSection As String, _
    ByVal lpKey As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long


Public Const MAX_PATH As Long = 256
Public Const MAX_RETURNED As Long = 512

Public gstrKeyValue As String * MAX_PATH

' --------------------------------------------------------------------
Public Function GetINISetting(ByVal sPath As String, _
                              ByVal sSection As String, _
                              ByVal sKey As String, _
                           Optional sDefault As String = "") As String
' --------------------------------------------------------------------
                     
  Dim lLenOfReturned As Long 'Length of the string returned
  Dim sReturned As String * MAX_RETURNED
  Dim value As String
  ' add null char to stringd passed to api call
  sSection = FixSectionName(sSection) & vbNullChar 'remove [] if any
  sKey = sKey & vbNullChar
  sPath = sPath & vbNullChar
  
  If Len(sDefault) = 0 Then
    sDefault = " "
  End If
  
  sDefault = sDefault & vbNullChar
  
  lLenOfReturned = GetPrivateProfileString(sSection, _
                                           sKey, _
                                           sDefault, _
                                           sReturned, _
                                           MAX_RETURNED, _
                                           sPath)
  GetINISetting = Left$(sReturned, lLenOfReturned)
End Function
' --------------------------------------------------------------------
Public Function SaveINISetting(ByVal sPath As String, _
                               ByVal sSection As String, _
                               ByVal sKey As String, _
                               ByVal sValue As String) As Long
' --------------------------------------------------------------------
Dim lStatus As Long

' lStatus is Returns: Long—Nonzero on success, zero on failure.
' Sets GetLastError.
' The main diffrence between the save and delete functions
' the save will not delete a key or section
' add null char to stringd passed to api call
  sSection = FixSectionName(sSection) & vbNullChar 'remove [] if any
  sKey = sKey & vbNullChar
  sPath = sPath & vbNullChar
  sValue = sValue & vbNullChar
 
  lStatus = WritePrivateProfileString(sSection, _
                                      sKey, _
                                      sValue, _
                                      sPath)
  
  SaveINISetting = lStatus

End Function
' --------------------------------------------------------------------
Public Function DeleteINISetting(ByVal sPath As String, _
                               ByVal sSection As String, _
                               ByVal sKey As String) As Long
' --------------------------------------------------------------------
Dim lStatus As Long
Const sValue As Long = 0
' lStatus is Returns: Long—Nonzero on success, zero on failure.
' Sets GetLastError.
' add null char to string passed to api call
  sSection = FixSectionName(sSection) & vbNullChar 'remove [] if any
  If Len(sKey) > 0 Then
    sKey = sKey & vbNullChar
  End If
  sPath = sPath & vbNullChar
 
  lStatus = WritePrivateProfileString(sSection, _
                                      sKey, _
                                      sValue, _
                                      sPath)
  If lStatus = 0 Then
    MsgBox "Error code is :" & Err.LastDllError
  End If
  DeleteINISetting = lStatus

End Function

' ----------------------------------------------------------------------
Private Function FixSectionName(ByVal sOldName As String) As String
' ----------------------------------------------------------------------
  Dim strNewName As String
  
  ' Remove the [brackets] around the section name
  ' field. The Windows API functions add the brackets for you.
  
  If Left$(sOldName, 1) = "[" Then
    strNewName = Mid$(sOldName, 2)
  Else
    strNewName = sOldName
  End If
  
  If Right$(sOldName, 1) = "]" Then
    strNewName = Left$(strNewName, Len(strNewName) - 1)
  End If
  
  FixSectionName = strNewName

End Function
'Public Declare Function GetPrivateProfileSection Lib "kernel32" _
'    Alias "GetPrivateProfileSectionA" _
'    (ByVal lpSection As String, _
'     ByVal lpReturned As String, _
'     ByVal nSize As Long, _
'     ByVal lpFileName As String) As Long
'' ----------------------------------------------------------------------
'Public Declare Function WritePrivateProfileSection Lib "kernel32" _
'    Alias "WritePrivateProfileSectionA" _
'   (ByVal lpSection As String, _
'    ByVal lpString As String, _
'    ByVal lpFileName As String) As Long
'



