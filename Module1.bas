Attribute VB_Name = "Module1"
Option Explicit


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'--------------------------------------------------
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'--------------------------------------------------
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  
End If

lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetAllKeys(hKey As Long, strPath As String) As Variant


Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    
    ReDim Preserve strNames(lCounter) As String
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strBuffer
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

GetAllKeys = strNames
End Function


Public Function GetAllValues(hKey As Long, strPath As String) As Variant

Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do
  
  lValueNameSize = 255
  strValueName = String$(lValueNameSize, " ")
  lDataBufferSize = 4000
  
  lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
  
  If lRegResult = ERROR_SUCCESS Then
    
    
    ReDim Preserve strNames(lCounter) As String
    ReDim Preserve lTypes(lCounter) As Long
    lTypes(UBound(lTypes)) = lValueType
    
    
    intZeroPos = InStr(strValueName, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strValueName
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop


Dim Finisheddata() As Variant
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
  Finisheddata(lCounter, 0) = strNames(lCounter)
  Finisheddata(lCounter, 1) = lTypes(lCounter)
Next

GetAllValues = Finisheddata

End Function

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long


If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If


lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_SZ Then
    
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left$(strBuffer, intZeroPos - 1)
    Else
      GetSettingString = strBuffer
    End If

  End If

Else
  
End If

lRegResult = RegCloseKey(hCurKey)
End Function


