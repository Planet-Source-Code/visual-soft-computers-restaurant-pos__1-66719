Attribute VB_Name = "iniMod"
Public Arq0 As String
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Function ReadINI(strFile As String, strKey As String, strName As String) As String
Dim intLen As Integer
Dim strText As String
'strText = Space(255)
strText = "                                                                                                    "
intLen = GetPrivateProfileString(strKey, strName, "", strText, Len(strText), strFile)
If intLen > -1 Then
    strText = Left(strText, intLen)
Else
MsgBox "INI skedari nuk u gjet. Kontrolloni përsëri"
End
End If
ReadINI = strText
End Function
Public Sub WriteINI(strFile As String, strKey As String, strName As String, strText As String)
Dim intLen As Integer
intLen = WritePrivateProfileString(strKey, strName, strText, strFile)
End Sub


Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "dd/mm/yyyy")
End Function
