VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpbData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    
'Structures needed for Registry API

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

'Registry Specific Access Rights
Const KEY_ALL_ACCESS = &H3F
Const KEY_READ = &H20019

'Open/Create Options
Const REG_OPTION_NON_VOLATILE = 0&
Const REG_OPTION_VOLATILE = &H1

'Registry Value Type
Const REG_SZ = 1

'Registry Base Key Enums
Public Enum RootKeyEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA_WIN2K_ONLY = &H80000004 'Windows 2000 only
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum
Private Const ERR_NONE = 0
Function SetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Variant) As Boolean
   
   Dim lRegHWND As Long
   Dim lngValue As Long
   Dim strValue As String
   Dim retVal As Long
   
    Dim SecAttr As SECURITY_ATTRIBUTES
    SecAttr.nLength = Len(SecAttr)
    SecAttr.lpSecurityDescriptor = 0
    SecAttr.bInheritHandle = True
    retVal = RegCreateKeyEx(hKey, KeyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, lRegHWND, retVal)
    If retVal Then Exit Function
    strValue = Value
    retVal = RegSetValueEx(lRegHWND, ValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
    RegCloseKey lRegHWND
    SetRegistryValue = (retVal = 0)
End Function
Public Function GetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim lRegHWND As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim retVal As Long
    Dim lValueType As Long
   
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, lRegHWND) Then Exit Function
    retVal = RegQueryValueEx(lRegHWND, ValueName, 0, lValueType, ByVal 0, lDataBufferSize)
    If lValueType = REG_SZ Then
       strBuffer = String(lDataBufferSize, " ")
       retVal = RegQueryValueEx(lRegHWND, ValueName, 0, 0, ByVal strBuffer, lDataBufferSize)
       If InStr(strBuffer, vbNullChar) > 0 Then
          GetRegistryValue = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
       Else
          GetRegistryValue = strBuffer
       End If
    End If
    RegCloseKey lRegHWND
End Function
Public Function DeleteRegistryValue(ByVal hKey As RootKeyEnum, RegKeyName As String, ValueName As String) As Boolean
    Dim lRetval As Long
    Dim lRegHWND As Long
   
    lRetval = RegOpenKeyEx(hKey, RegKeyName, 0, KEY_ALL_ACCESS, lRegHWND)
    If lRetval = ERR_NONE Then
       lRetval = RegDeleteValue(lRegHWND, ValueName)
       If lRetval <> ERR_NONE Then Exit Function
       lRetval = RegCloseKey(lRegHWND)
       If lRetval = ERR_NONE Then DeleteRegistryValue = True
   End If
End Function
Public Function DeleteRegistryKey(ByVal hKey As RootKeyEnum, RegKeyName As String, ValueName As String) As Boolean
    Dim lRetval As Long
    Dim lRegHWND As Long
   
    lRetval = RegOpenKeyEx(hKey, RegKeyName, 0, KEY_ALL_ACCESS, lRegHWND)
    If lRetval = ERR_NONE Then
       lRetval = RegDeleteKey(lRegHWND, ValueName)
       If lRetval <> ERR_NONE Then Exit Function
       lRetval = RegCloseKey(lRegHWND)
       If lRetval = ERR_NONE Then DeleteRegistryKey = True
   End If
End Function
