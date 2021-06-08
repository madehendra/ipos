Attribute VB_Name = "modRegEdit"
''''''''            ''''''''        ''''     ''''
'''    ''           '''    ''         '''   '''
'''    '' '''    '' '''    '' ''''''''  '''''
''''''''   ''    '' ''''''''  '''    ''  '''
'''        ''    '' '''       '''    ''  '''
'''          '''' ' '''       ''''''''   '''
                              '''
                              '''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This module calls the necessary api functions    '
'to read/write to the registry                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Programed by Frankie Miklos -aKa- PuPpY          '
'Credit: http://vbnet.mvps.org/                   '
'Last update: Thursday, June 24, 2004 (13:29)     '
'''''''''''''''''''''''''''''''''''''''''''''''''''


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal KeyRoot As kRoot, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Enum regType
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
End Enum

Const REG_OPTION_NON_VOLATILE = 0

Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
Public Enum kRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Function WriteDWORD(ByVal KeyRoot As kRoot, ByVal KeyName As String, ByVal SubKeyName As String, ByVal SubKeyValue As Long) As Boolean
Dim r As Long, hkey As Long
    
    r = RegCreateKey(KeyRoot, KeyName, hkey)
    
    If (r <> ERROR_SUCCESS) Then GoTo Err_Hnd
    
    r = RegSetValueEx(hkey, SubKeyName, 0, REG_DWORD, SubKeyValue, 4)
                       
    If (r <> ERROR_SUCCESS) Then GoTo Err_Hnd

    RegCloseKey hkey
    
    WriteDWORD = True

Exit Function
Err_Hnd:
    
    WriteDWORD = False
    RegCloseKey hkey
    
End Function

Public Function RegRead(KeyRoot As kRoot, KeyName As String, SubKeyName As String) As String
Dim i As Long, r As Long, hkey As Long, hDepth As Long, lKeyValType As Long, KeyValSize As Long
Dim sKeyVal As String, tmpVal As String
    
    r = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hkey)
    
    If (r <> ERROR_SUCCESS) Then GoTo Err_Hnd
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
    r = RegQueryValueEx(hkey, SubKeyName, 0, lKeyValType, tmpVal, KeyValSize)
                        
    If (r <> ERROR_SUCCESS) Then GoTo Err_Hnd
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    Select Case lKeyValType
        Case REG_SZ, REG_EXPAND_SZ
            sKeyVal = tmpVal
        Case REG_DWORD
            For i = Len(tmpVal) To 1 Step -1
                sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
            Next
            sKeyVal = Val(Format$("&h" + sKeyVal))
    End Select
    
    RegRead = sKeyVal
    RegCloseKey hkey

Exit Function
Err_Hnd:
    
    RegRead = vbNullString
    RegCloseKey hkey
    
End Function
